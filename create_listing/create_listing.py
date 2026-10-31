#!/usr/bin/env python3
"""
Create Reverb listings in Chrome (Selenium), one store at a time.

Excel columns:
  store_name, title, description, remove_branding, price, quantity,
  Category, Subcategory, sku, photos, status, upc, free_shipping

Per store: SoftEther VPN -> matching Chrome profile -> verify IP
on nordvpn.com/what-is-my-ip -> open the Reverb sell form and save.

Usage, from the project folder:
  python create_listing/create_listing.py create_listing/Create.xlsx
  python create_listing/create_listing.py --store GG --limit 1
  python create_listing/create_listing.py --no-vpn --skip-ip-check
"""
from __future__ import annotations

import argparse
import html
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Optional

import requests
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

BOT_DIR = Path(__file__).resolve().parent
ROOT = BOT_DIR.parent
STORES_PATH = ROOT / "stores.json"
LOG_PATH = BOT_DIR / "create_listing.log"
OUTPUT_DIR = ROOT / "output"
FAILURES_DIR = OUTPUT_DIR / "failures"
PHOTO_DIR = OUTPUT_DIR / "photos"
DEFAULT_EXCEL = Path(r"G:\All Tools Created\Reverb Lisitng Automation\test_listings.xlsx")

IP_CHECK_URL = "https://nordvpn.com/what-is-my-ip/"
IPIFY_URL = "https://api.ipify.org"
SELL_URL = "https://reverb.com/my/selling/listings/new"
SIGN_IN_URL = "https://reverb.com/signin?redirect_to=%2Fmy%2Fselling%2Flistings%2Fnew"
IPV4_RE = re.compile(r"\b(\d{1,3}(?:\.\d{1,3}){3})\b")
URL_RE = re.compile(r"https?://[^\s|,]+")

VPN_CONNECT_TIMEOUT = 45
VPN_STATUS_POLL = 2
IP_SETTLE_TIMEOUT = 60
IP_SETTLE_POLL = 4
ROW_DELAY = 3
CHROME_CONNECT_TIMEOUT = 30
HUMAN_CHECK_TIMEOUT = 180
SIGN_IN_TIMEOUT = 300
MAX_PHOTOS = 8

# Same default the API script uses when the Excel condition cell is this UUID.
CONDITION_BY_UUID = {
    "7c3f45de-2ae0-4c81-8400-fdb6b1d74890": "Brand New",
}

HEADER_MAP = {
    "store name": "store",
    "store_name": "store",
    "storename": "store",
    "store": "store",
    "title": "title",
    "description": "description",
    "remove_branding": "remove_branding",
    "price": "price",
    "quantity": "quantity",
    "category": "category",
    "subcategory": "subcategory",
    "sub category": "subcategory",
    "sku": "sku",
    "photos": "photos",
    "status": "status",
    "publish": "publish",
    "upc": "upc",
    "free_shipping": "free_shipping",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("create_listing")


@dataclass
class ListingRow:
    row_number: int
    store: str
    title: str
    description: str
    price: str
    quantity: str
    condition: str
    category: str
    subcategory: str
    sku: str
    photos: list[str]
    publish: bool
    upc: str
    free_shipping: bool
    remove_branding: list[str] = field(default_factory=list)
    status: str = ""
    detail: str = ""
    photos_done: bool = False


def _load_config() -> dict:
    if not STORES_PATH.exists():
        raise SystemExit(f"Missing {STORES_PATH}")
    with STORES_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def _chrome_candidates() -> list[Path]:
    pf = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
    pf86 = Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
    return [
        pf / "Google" / "Chrome" / "Application" / "chrome.exe",
        pf86 / "Google" / "Chrome" / "Application" / "chrome.exe",
        Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "Application" / "chrome.exe",
    ]


def _vpncmd_candidates(configured: str) -> list[Path]:
    paths = []
    if configured:
        paths.append(Path(configured))
    env = os.environ.get("VPNCMD_PATH", "").strip()
    if env:
        paths.append(Path(env))
    pf = Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
    paths.append(pf / "SoftEther VPN Client" / "vpncmd.exe")
    return paths


def resolve_paths(cfg: dict) -> tuple[Path, Path, Path]:
    chrome_raw = (cfg.get("chrome_path") or "").strip()
    chrome = Path(chrome_raw) if chrome_raw else Path()
    if not chrome.is_file():
        chrome = next((p for p in _chrome_candidates() if p.is_file()), chrome)
    if not chrome.is_file():
        raise SystemExit("Chrome not found. Set chrome_path in stores.json")

    user_raw = (cfg.get("chrome_user_data_dir") or "").strip()
    user_data = Path(user_raw) if user_raw else Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"
    if not user_data.is_dir():
        raise SystemExit("Chrome User Data folder not found. Set chrome_user_data_dir in stores.json")

    vpncmd = Path("vpncmd")
    for candidate in _vpncmd_candidates(cfg.get("vpncmd_path") or ""):
        if candidate.is_file():
            vpncmd = candidate
            break
    return chrome, user_data, vpncmd


def store_lookup(cfg: dict, name: str) -> Optional[dict]:
    stores = cfg.get("stores") or {}
    key = (name or "").strip().upper()
    if key not in stores:
        return None
    row = dict(stores[key])
    row["name"] = key
    return row


def coerce_code(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", text):
        try:
            return format(Decimal(text), "f").split(".", 1)[0]
        except (InvalidOperation, ValueError):
            pass
    return text


def coerce_money(value: Any) -> str:
    text = coerce_code(value)
    if not text:
        return ""
    try:
        return f"{float(text):.2f}"
    except ValueError:
        return text


def parse_bool(value: Any) -> bool:
    text = coerce_code(value).lower()
    return text in {"true", "1", "yes", "y", "on"}


def _plain(fragment: str) -> str:
    text = re.sub(r"<[^>]+>", "", fragment or "")
    return re.sub(r"[ \t]+", " ", html.unescape(text)).strip()


def html_to_text(raw: str) -> str:
    text = re.sub(r"<!--.*?-->", "", raw or "", flags=re.S)
    text = re.sub(
        r"(?is)<h[1-6][^>]*>(.*?)</h[1-6]>",
        lambda m: f"\n{_plain(m.group(1))}\n",
        text,
    )
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n\n", text)

    def list_item(match: re.Match) -> str:
        inner = _plain(match.group(1))
        if not inner:
            return ""
        head, sep, body = inner.partition(":")
        if sep and body.strip() and len(head.strip()) <= 80:
            return f"\n{head.strip()}\n- {body.strip()}\n"
        return f"\n- {inner}\n"

    text = re.sub(r"(?is)<li[^>]*>(.*?)</li>", list_item, text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    lines: list[str] = []
    blank = False
    for line in text.splitlines():
        line = re.sub(r"[ \t]+", " ", line).strip()
        if not line:
            if lines and not blank:
                lines.append("")
                blank = True
            continue
        blank = False
        lines.append(line)
    return "\n".join(lines).strip()


def strip_brands(text: str, brands: list[str]) -> str:
    cleaned = text or ""
    for brand in sorted(brands, key=len, reverse=True):
        if not brand or brand.lower() == "generic":
            continue
        cleaned = re.compile(re.escape(brand), re.I).sub("", cleaned)
    lines = [re.sub(r"[ \t]{2,}", " ", line).strip() for line in cleaned.splitlines()]
    return "\n".join(lines).strip()


def parse_photos(value: Any) -> list[str]:
    text = "" if value is None else str(value)
    if not text or text.lower() == "nan":
        return []
    found: list[str] = []
    for part in re.split(r"[\n|]+", text):
        part = part.strip().strip(",")
        if not part:
            continue
        if part.lower().startswith("http"):
            found.extend(URL_RE.findall(part))
        elif Path(part).is_file():
            found.append(str(Path(part)))
    return found or URL_RE.findall(text)


def condition_label(value: str) -> str:
    text = (value or "").strip()
    if text in CONDITION_BY_UUID:
        return CONDITION_BY_UUID[text]
    if re.fullmatch(r"[0-9a-fA-F-]{36}", text):
        return "Brand New"
    return text or "Brand New"


def read_excel(path: Path) -> list[ListingRow]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_cells:
        raise SystemExit("Excel has no header row")
    index: dict[str, int] = {}
    for i, col in enumerate(header_cells):
        mapped = HEADER_MAP.get(str(col or "").strip().lower())
        if mapped and mapped not in index:
            index[mapped] = i
    missing = [c for c in ("store", "sku", "title") if c not in index]
    if missing:
        raise SystemExit(f"Excel is missing columns: {', '.join(missing)}")

    def cell(values: tuple, key: str) -> Any:
        pos = index.get(key)
        if pos is None or pos >= len(values):
            return None
        return values[pos]

    rows: list[ListingRow] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not values:
            continue
        store = coerce_code(cell(values, "store")).upper()
        sku = coerce_code(cell(values, "sku"))
        title = coerce_code(cell(values, "title"))
        if not store and not sku and not title:
            continue
        brands = [b.strip() for b in coerce_code(cell(values, "remove_branding")).split(",") if b.strip()]
        description = html_to_text(coerce_code(cell(values, "description"))) or title
        title = strip_brands(title, brands) or title or "Untitled"
        description = strip_brands(description, brands) or title
        status_text = coerce_code(cell(values, "status")).lower()
        publish_cell = coerce_code(cell(values, "publish")).lower()
        if status_text in {"draft"}:
            publish = False
        elif status_text in {"live", "published", "publish"}:
            publish = True
        elif publish_cell:
            publish = publish_cell in {"true", "1", "yes", "y", "on"}
        else:
            publish = False
        qty = coerce_code(cell(values, "quantity")) or "1"
        rows.append(
            ListingRow(
                row_number=row_number,
                store=store,
                title=title[:140],
                description=description[:8000],
                price=coerce_money(cell(values, "price")) or "1.00",
                quantity=qty.split(".", 1)[0] or "1",
                condition="Brand New",
                category=coerce_code(cell(values, "category")),
                subcategory=coerce_code(cell(values, "subcategory")),
                sku=sku,
                photos=parse_photos(cell(values, "photos"))[:MAX_PHOTOS],
                publish=publish,
                upc=coerce_code(cell(values, "upc")),
                free_shipping=parse_bool(cell(values, "free_shipping")),
                remove_branding=brands,
            )
        )
    return rows


def write_results(rows: list[ListingRow], dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "results"
    ws.append(["STORE NAME", "SKU", "TITLE", "STATUS", "DETAIL"])
    for r in rows:
        ws.append([r.store, r.sku, r.title, r.status, r.detail])
    wb.save(dest)


def run_vpncmd(vpncmd: Path, parts: list[str], timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        [str(vpncmd), "localhost", "/CLIENT", "/CMD", *parts],
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=False,
    )


def vpn_account_status(vpncmd: Path, account: str) -> str:
    try:
        result = run_vpncmd(vpncmd, ["AccountStatusGet", account], timeout=20)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return "error"
    out = ((result.stdout or "") + "\n" + (result.stderr or "")).lower()
    if "not connected" in out or "error code: 37" in out:
        return "down"
    if "does not exist" in out or "error code: 36" in out:
        return "error"
    if "connection completed" in out or "session established" in out or "established" in out:
        return "up"
    if "connecting" in out:
        return "connecting"
    return "error" if result.returncode != 0 else "down"


def disconnect_all_vpns(vpncmd: Path, cfg: dict) -> None:
    seen = set()
    for info in (cfg.get("stores") or {}).values():
        account = (info.get("vpn") or "").strip()
        if not account or account in seen:
            continue
        seen.add(account)
        try:
            run_vpncmd(vpncmd, ["AccountDisconnect", account], timeout=20)
        except Exception:
            pass


def connect_store_vpn(vpncmd: Path, account: str) -> bool:
    try:
        result = run_vpncmd(vpncmd, ["AccountConnect", account], timeout=30)
    except FileNotFoundError:
        log.error("vpncmd not found: %s", vpncmd)
        return False
    except subprocess.TimeoutExpired:
        log.error("vpncmd timed out connecting %s", account)
        return False
    out = ((result.stdout or "") + "\n" + (result.stderr or "")).lower()
    if result.returncode == 0 or "command completed successfully" in out:
        log.info("VPN connect issued for %s", account)
        return True
    log.error("VPN connect failed for %s: %s", account, (result.stderr or result.stdout or "").strip())
    return False


def wait_for_vpn(vpncmd: Path, account: str) -> bool:
    start = time.time()
    while time.time() - start < VPN_CONNECT_TIMEOUT:
        state = vpn_account_status(vpncmd, account)
        if state == "up":
            log.info("VPN session established for %s", account)
            return True
        if state == "error":
            log.error("VPN status error for account %s", account)
            return False
        time.sleep(VPN_STATUS_POLL)
    log.error("VPN did not come up for %s", account)
    return False


def close_chrome() -> None:
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe", "/T"], capture_output=True, text=True, shell=False)
    time.sleep(3)


DEBUG_USER_DATA = ROOT / "chrome_debug_data"
_SKIP_PROFILE_DIRS = {
    "Cache",
    "Code Cache",
    "GPUCache",
    "DawnGraphiteCache",
    "DawnWebGPUCache",
    "GrShaderCache",
    "ShaderCache",
    "Service Worker",
    "CacheStorage",
    "Crashpad",
    "BrowserMetrics",
    "blob_storage",
}


def _ignore_profile_junk(_directory: str, names: list[str]) -> list[str]:
    return [name for name in names if name in _SKIP_PROFILE_DIRS]


def automation_user_data(user_data: Path, profile: str) -> Path:
    """Chrome 153 blocks debugging on the normal User Data folder.

    A separate folder can be controlled, but its cookies are encrypted
    separately, so the normal Chrome login is not reused. Copy once, then
    keep whatever login the user adds in that window.
    """
    default = Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"
    try:
        is_default = user_data.resolve().samefile(default)
    except OSError:
        is_default = os.path.normcase(str(user_data)) == os.path.normcase(str(default))
    if not is_default:
        return user_data
    source = user_data / profile
    if not source.is_dir():
        raise RuntimeError(f"Chrome profile folder not found: {source}")
    dest_root = DEBUG_USER_DATA
    dest_profile = dest_root / profile
    if dest_profile.exists():
        log.info("Using saved automation profile %s", dest_profile)
        return dest_root
    dest_root.mkdir(parents=True, exist_ok=True)
    local_state = user_data / "Local State"
    if local_state.is_file():
        shutil.copy2(local_state, dest_root / "Local State")
    log.info("Copying %s once into %s", profile, dest_root)
    shutil.copytree(source, dest_profile, dirs_exist_ok=True, ignore=_ignore_profile_junk)
    return dest_root


def launch_for_user(chrome: Path, user_data: Path, profile: str, url: str, port: int) -> subprocess.Popen:
    """Debug port stays open so we can see the page, without attaching Selenium yet."""
    args = [
        str(chrome),
        f"--remote-debugging-port={port}",
        "--remote-allow-origins=*",
        f"--user-data-dir={user_data}",
        f"--profile-directory={profile}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble",
        url,
    ]
    log.info("Opening Chrome for Reverb sign-in")
    return subprocess.Popen(args)


def chrome_tabs(port: int) -> list[dict]:
    try:
        resp = requests.get(f"http://127.0.0.1:{port}/json/list", timeout=2)
        if resp.ok:
            data = resp.json()
            if isinstance(data, list):
                return data
    except Exception:
        return []
    return []


def let_user_finish_login(chrome: Path, user_data: Path, profile: str, port: int) -> None:
    del chrome, user_data, profile
    log.warning("Solve the check in the Chrome window that is already open. The tabs will stay open for %s seconds.", SIGN_IN_TIMEOUT)
    deadline = time.time() + SIGN_IN_TIMEOUT
    last = ""
    while time.time() < deadline:
        for tab in chrome_tabs(port):
            url = str(tab.get("url") or "")
            title = str(tab.get("title") or "")
            shown = f"{title} | {url}".strip()
            if shown and shown != last:
                log.info("Chrome page: %s", shown[:180])
                last = shown
            if re.search(r"are you human|just a moment|verifying", title, re.I):
                continue
            if "listings/new" in url or re.search(r"product info|tell us about your gear", f"{title} {url}", re.I):
                log.info("Listing form is open. Starting the listing.")
                return
        time.sleep(2)
    raise RuntimeError("Captcha was not solved in time. Chrome was left open.")


def start_chrome(chrome: Path, user_data: Path, profile: str, port: int) -> subprocess.Popen:
    args = [
        str(chrome),
        f"--remote-debugging-port={port}",
        "--remote-allow-origins=*",
        "--disable-features=DevToolsDebuggingRestrictions",
        f"--user-data-dir={user_data}",
        f"--profile-directory={profile}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble",
        IP_CHECK_URL,
    ]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    err_path = OUTPUT_DIR / "chrome_stderr.log"
    err_file = err_path.open("w", encoding="utf-8", errors="replace")
    log.info("Launching Chrome %s from %s", profile, user_data)
    proc = subprocess.Popen(args, stdout=err_file, stderr=subprocess.STDOUT)
    proc.stderr_log = err_file  # type: ignore[attr-defined]
    return proc


def wait_for_debug_port(port: int) -> None:
    url = f"http://127.0.0.1:{port}/json/version"
    deadline = time.time() + CHROME_CONNECT_TIMEOUT
    last = "port closed"
    while time.time() < deadline:
        try:
            resp = requests.get(url, timeout=2)
            if resp.ok and "webSocketDebuggerUrl" in resp.text:
                log.info("Chrome debug port %s is open", port)
                return
            last = f"HTTP {resp.status_code}"
        except Exception as exc:
            last = str(exc)
        time.sleep(1)
    hint = ""
    err_path = OUTPUT_DIR / "chrome_stderr.log"
    if err_path.is_file():
        text = err_path.read_text(encoding="utf-8", errors="replace")
        if "non-default data directory" in text:
            hint = " Chrome refused debugging on this User Data folder."
    raise RuntimeError(f"Chrome did not open debug port {port} ({last}).{hint}")


def local_chromedriver(browser_version: str) -> Path:
    cache = Path.home() / ".cache" / "selenium" / "chromedriver" / "win64"
    drivers = [path for path in cache.glob("*/chromedriver.exe") if path.is_file()]
    if not drivers:
        raise RuntimeError("ChromeDriver is not installed in the Selenium cache.")
    major = (browser_version or "").split(".", 1)[0]
    same_major = [path for path in drivers if path.parent.name.startswith(f"{major}.")]
    pool = same_major or drivers
    return max(pool, key=lambda path: [int(part) if part.isdigit() else 0 for part in path.parent.name.split(".")])


def connect_selenium(port: int) -> WebDriver:
    wait_for_debug_port(port)
    browser = ""
    try:
        meta = requests.get(f"http://127.0.0.1:{port}/json/version", timeout=2).json()
        browser = str(meta.get("Browser") or "").split("/", 1)[-1]
    except Exception:
        browser = ""
    driver_path = local_chromedriver(browser)
    log.info("Attaching Selenium with ChromeDriver %s", driver_path.parent.name)
    options = Options()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    service = Service(executable_path=str(driver_path))
    holder: dict[str, Any] = {}

    def _attach() -> None:
        try:
            holder["driver"] = webdriver.Chrome(service=service, options=options)
        except Exception as exc:
            holder["error"] = exc

    worker = threading.Thread(target=_attach, daemon=True)
    worker.start()
    worker.join(45)
    if worker.is_alive():
        service.stop()
        raise RuntimeError(
            "Attaching Selenium timed out. Close the automation Chrome window and run the script again."
        )
    if "error" in holder:
        raise RuntimeError(f"Could not attach Selenium to Chrome on port {port}: {holder['error']}") from holder["error"]
    log.info("Selenium attached")
    return holder["driver"]


def screenshot(driver: WebDriver, name: str) -> None:
    FAILURES_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:80]
    dest = FAILURES_DIR / f"{datetime.now().strftime('%H%M%S')}_{safe}.png"
    try:
        driver.save_screenshot(str(dest))
        log.info("Saved screenshot %s", dest.name)
    except Exception as exc:
        log.debug("Screenshot failed: %s", exc)


def page_text(driver: WebDriver) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def ip_from_text(text: str) -> str:
    match = IPV4_RE.search(text or "")
    return match.group(1) if match else ""


def read_ip(driver: WebDriver) -> str:
    try:
        driver.get(IP_CHECK_URL)
        time.sleep(4)
        for css in ("p.heading-xl", "p[data-hk]", "[class*='heading-xl']"):
            found = driver.find_elements(By.CSS_SELECTOR, css)
            if found:
                ip = ip_from_text(found[0].text)
                if ip:
                    return ip
        ip = ip_from_text(page_text(driver))
        if ip:
            return ip
    except Exception as exc:
        log.warning("NordVPN IP page failed: %s", exc)
    try:
        driver.get(IPIFY_URL)
        time.sleep(2)
        return ip_from_text(driver.find_element(By.TAG_NAME, "body").text)
    except Exception:
        return ""


def wait_for_expected_ip(driver: WebDriver, expected: str) -> tuple[bool, str]:
    shown = ""
    deadline = time.time() + IP_SETTLE_TIMEOUT
    while True:
        shown = read_ip(driver)
        if expected and shown == expected:
            log.info("IP verified: %s", shown)
            return True, shown
        if time.time() >= deadline:
            break
        log.info("Current IP %s, waiting for %s ...", shown or "(unknown)", expected)
        time.sleep(IP_SETTLE_POLL)
    return False, shown


def click_text(driver: WebDriver, pattern: str, timeout: float = 2) -> bool:
    xpath = (
        "//button|//a|//*[@role='button']|//*[@role='option']|//*[@role='menuitem']|"
        "//*[@role='combobox']|//li"
    )
    end = time.time() + timeout
    rx = re.compile(pattern, re.I)
    while time.time() < end:
        for el in driver.find_elements(By.XPATH, xpath):
            try:
                label = (el.text or el.get_attribute("aria-label") or "").strip()
                if label and rx.search(label) and el.is_displayed() and el.is_enabled():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    el.click()
                    time.sleep(0.6)
                    return True
            except Exception:
                continue
        time.sleep(0.3)
    return False


def visible_inputs(driver: WebDriver) -> list[WebElement]:
    out = []
    for el in driver.find_elements(By.CSS_SELECTOR, "input, textarea, [contenteditable='true']"):
        try:
            if el.is_displayed():
                out.append(el)
        except Exception:
            continue
    return out


def fill_element(el: WebElement, value: str) -> None:
    tag = (el.tag_name or "").lower()
    if el.get_attribute("contenteditable") == "true" or tag == "div":
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(value)
        return
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    el.send_keys(value)


def input_matches(el: WebElement, words: tuple[str, ...]) -> bool:
    blob = " ".join(
        filter(
            None,
            [
                el.get_attribute("name"),
                el.get_attribute("id"),
                el.get_attribute("placeholder"),
                el.get_attribute("aria-label"),
                el.get_attribute("data-testid"),
            ],
        )
    ).lower()
    return any(word in blob for word in words)


def fill_matching(driver: WebDriver, words: tuple[str, ...], value: str) -> bool:
    if not value:
        return False
    for el in visible_inputs(driver):
        try:
            if el.get_attribute("type") in {"file", "hidden", "checkbox", "radio"}:
                continue
            if input_matches(el, words):
                fill_element(el, value)
                time.sleep(0.3)
                return True
        except Exception:
            continue
    return False


def pick_suggestion(driver: WebDriver, text: str) -> bool:
    time.sleep(0.8)
    return click_text(driver, re.escape(text), timeout=3) or click_text(
        driver, re.escape(text.split()[0]) if text.split() else text, timeout=2
    )


def download_photos(row: ListingRow) -> list[str]:
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    paths: list[str] = []
    for i, url in enumerate(row.photos, start=1):
        if Path(url).is_file():
            paths.append(str(Path(url)))
            continue
        try:
            resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            resp.raise_for_status()
            ext = "jpg"
            ctype = (resp.headers.get("Content-Type") or "").lower()
            if "png" in ctype:
                ext = "png"
            elif "webp" in ctype:
                ext = "webp"
            dest = PHOTO_DIR / f"{re.sub(r'[^A-Za-z0-9._-]+', '_', row.sku)[:40]}_{i}.{ext}"
            dest.write_bytes(resp.content)
            paths.append(str(dest))
        except Exception as exc:
            log.warning("Photo download failed %s: %s", url[:80], exc)
    return paths


def upload_photos(driver: WebDriver, paths: list[str]) -> bool:
    if not paths:
        return False
    inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    for el in inputs:
        try:
            el.send_keys("\n".join(paths))
            time.sleep(2)
            return True
        except Exception:
            continue
    return False


def human_check_visible(driver: WebDriver) -> bool:
    body = page_text(driver)
    return bool(re.search(r"are you human|i'm not a robot|i’m not a robot", body, re.I))


def challenge_visible(driver: WebDriver) -> bool:
    if human_check_visible(driver) or cloudflare_widget(driver):
        return True
    try:
        shown = f"{driver.title or ''} {driver.current_url or ''}"
    except Exception:
        return False
    return bool(re.search(r"are you human|just a moment|verifying|challenge", shown, re.I))


def wait_for_human_check(driver: WebDriver) -> None:
    if not challenge_visible(driver):
        return
    log.warning(
        "Captcha is on the screen. Solve it in this Chrome window. The bot will wait %s seconds and will not close the tabs.",
        SIGN_IN_TIMEOUT,
    )
    deadline = time.time() + SIGN_IN_TIMEOUT
    while time.time() < deadline:
        time.sleep(2)
        try:
            if not challenge_visible(driver):
                log.info("Captcha is gone. Continuing the listing.")
                return
        except Exception:
            log.info("Chrome is still open. Finish the captcha in that window.")
    raise RuntimeError("Captcha was not solved in time. Chrome was left open.")


def on_auth_page(driver: WebDriver) -> bool:
    return bool(re.search(r"/(signin|login|session|signup)", driver.current_url or "", re.I))


def cloudflare_widget(driver: WebDriver) -> bool:
    try:
        html = driver.page_source or ""
    except Exception:
        html = ""
    if re.search(r"challenges\.cloudflare\.com|cf-turnstile|verification failed", html, re.I):
        return True
    for frame in driver.find_elements(By.CSS_SELECTOR, "iframe"):
        blob = " ".join(
            filter(None, [frame.get_attribute("src"), frame.get_attribute("title")])
        )
        if re.search(r"cloudflare|turnstile|challenge", blob, re.I):
            return True
    return False


def open_sell_form(driver: WebDriver) -> None:
    current = driver.current_url or ""
    if "listings/new" in current and not on_auth_page(driver):
        log.info("Already on the listing form")
        return
    log.info("Opening %s", SELL_URL)
    driver.get(SELL_URL)
    time.sleep(3)
    wait_for_human_check(driver)
    if re.search(r"/signup", driver.current_url or "", re.I):
        log.info("Signup page opened. Switching to sign in.")
        driver.get(SIGN_IN_URL)
        time.sleep(2)
        click_text(driver, r"^sign in\.?$", timeout=2)
        wait_for_human_check(driver)
    if challenge_visible(driver):
        wait_for_human_check(driver)
    if challenge_visible(driver):
        raise RuntimeError("Captcha was not solved in time. Chrome was left open.")
    if on_auth_page(driver):
        raise RuntimeError("Reverb is not signed in")
    if "listings/new" not in (driver.current_url or ""):
        driver.get(SELL_URL)
        time.sleep(3)
        wait_for_human_check(driver)
        if challenge_visible(driver):
            raise RuntimeError("Captcha was not solved in time. Chrome was left open.")
        if on_auth_page(driver):
            raise RuntimeError("Reverb is not signed in")
    if "listings/new" not in (driver.current_url or ""):
        click_text(driver, r"from scratch|create a new listing|list an item|sell your gear", timeout=4)
        time.sleep(2)


def labeled_input(driver: WebDriver, label: str) -> Optional[WebElement]:
    return driver.execute_script(
        """
        const want = arguments[0].toLowerCase();
        const norm = (s) => (s || '').replace(/\\*/g, '').trim().toLowerCase();
        const nodes = [...document.querySelectorAll('label, span, p, div')];
        for (const node of nodes) {
          const raw = (node.childNodes.length
            ? [...node.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('')
            : node.innerText || '');
          const text = norm(raw.split('\\n')[0]);
          if (!text || (text !== want && !text.startsWith(want + ' '))) continue;
          let el = node;
          for (let i = 0; i < 6 && el; i++) {
            const sib = el.nextElementSibling;
            if (sib) {
              const input = sib.matches('input,textarea')
                ? sib
                : sib.querySelector('input:not([type=hidden]):not([type=file]), textarea');
              if (input) return input;
            }
            el = el.parentElement;
          }
        }
        return null;
        """,
        label,
    )


def fill_labeled(driver: WebDriver, label: str, value: str) -> bool:
    if not value:
        return False
    try:
        el = labeled_input(driver, label)
        if el is None:
            return False
        fill_element(el, value)
        log.info("Filled %s", label)
        time.sleep(0.3)
        return True
    except Exception as exc:
        log.debug("Could not fill %s: %s", label, exc)
        return False


_CATEGORY_BY_UUID: Optional[dict[str, str]] = None


def category_full_name(value: str) -> str:
    text = (value or "").strip()
    if not re.fullmatch(r"[0-9a-fA-F-]{36}", text):
        return text
    global _CATEGORY_BY_UUID
    if _CATEGORY_BY_UUID is None:
        _CATEGORY_BY_UUID = {}
        candidates = [
            ROOT / "reverb_categories_cache.json",
            Path(r"G:\All Tools Created\Bulk Listing Formatter\output\reverb_categories_cache.json"),
            Path.home() / "Downloads" / "category uuid.json",
        ]
        for path in candidates:
            if not path.is_file():
                continue
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except Exception as exc:
                log.warning("Could not read categories from %s: %s", path, exc)
                continue
            items = data.get("categories") if isinstance(data, dict) else data
            for item in items or []:
                if not isinstance(item, dict):
                    continue
                uuid = str(item.get("uuid") or "").strip().lower()
                full = str(item.get("full_name") or item.get("name") or "").strip()
                if uuid and full:
                    _CATEGORY_BY_UUID[uuid] = full
            if _CATEGORY_BY_UUID:
                log.info("Loaded %s categories from %s", len(_CATEGORY_BY_UUID), path)
                break
    found = _CATEGORY_BY_UUID.get(text.lower(), "")
    if not found:
        log.warning("No category name for UUID %s", text)
    return found


def click_exact_label(driver: WebDriver, text: str) -> bool:
    if click_text(driver, rf"^{re.escape(text)}$", timeout=2):
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const want = arguments[0].toLowerCase();
                const visible = (n) => {
                  const r = n.getBoundingClientRect();
                  return r.width > 2 && r.height > 2;
                };
                const nodes = [...document.querySelectorAll('[role=option], li, button, [role=menuitem], div, span')];
                const hits = nodes.filter(n => visible(n) && (n.innerText || '').trim().toLowerCase() === want);
                hits.sort((a, b) => a.getBoundingClientRect().height - b.getBoundingClientRect().height);
                const hit = hits[0];
                if (!hit) return false;
                hit.scrollIntoView({block:'center'});
                hit.click();
                return true;
                """,
                text,
            )
        )
    except Exception:
        return False


def open_category_dropdown(driver: WebDriver) -> bool:
    if open_select_one(driver):
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const labels = [...document.querySelectorAll('label, span, p, div')];
                const label = labels.find(n => /^category\\b/i.test(((n.innerText || '').trim().split('\\n')[0] || '').replace('*','').trim()));
                if (!label) return false;
                let scope = label.parentElement;
                for (let i = 0; i < 6 && scope; i++) {
                  const control = [...scope.querySelectorAll('button, [role=combobox], [role=button], input')]
                    .find(n => n !== label && n.getBoundingClientRect().height > 2);
                  if (control) {
                    control.scrollIntoView({block:'center'});
                    control.click();
                    control.focus();
                    return true;
                  }
                  scope = scope.parentElement;
                }
                return false;
                """
            )
        )
    except Exception:
        return False


def open_select_one(driver: WebDriver) -> bool:
    if click_text(driver, r"^select one\.{0,3}$", timeout=2):
        return True
    try:
        return bool(
            driver.execute_script(
                """
                const nodes = [...document.querySelectorAll('button, [role=combobox], [role=button]')];
                const hit = nodes.find(n => /^select one\\.{0,3}$/i.test((n.innerText || '').trim()) && n.offsetParent);
                if (!hit) return false;
                hit.scrollIntoView({block:'center'});
                hit.click();
                return true;
                """
            )
        )
    except Exception:
        return False


def category_parts(raw: str) -> list[str]:
    full = category_full_name(raw)
    if not full:
        return []
    direct = [part.strip() for part in full.split("/") if part.strip()]
    if len(direct) > 1:
        return direct
    leaf = direct[0].lower() if direct else ""
    for name in (_CATEGORY_BY_UUID or {}).values():
        parts = [part.strip() for part in name.split("/") if part.strip()]
        if parts and parts[-1].lower() == leaf and len(parts) > 1:
            return parts
    return direct


def _is_country_select(labels: list[str]) -> bool:
    lowered = {label.lower() for label in labels}
    return "united states" in lowered and "none" in lowered


def _select_options(driver: WebDriver) -> list[tuple[WebElement, list[str]]]:
    found: list[tuple[WebElement, list[str]]] = []
    for el in driver.find_elements(By.CSS_SELECTOR, "select"):
        try:
            labels = [opt.text.strip() for opt in Select(el).options if opt.text.strip()]
        except Exception:
            continue
        if labels and not _is_country_select(labels):
            found.append((el, labels))
    return found


def leave_country_none(driver: WebDriver) -> None:
    for el in driver.find_elements(By.CSS_SELECTOR, "select"):
        try:
            labels = [opt.text.strip() for opt in Select(el).options if opt.text.strip()]
        except Exception:
            continue
        if not _is_country_select(labels):
            continue
        current = ""
        try:
            current = (Select(el).first_selected_option.text or "").strip()
        except Exception:
            current = ""
        if current.lower() == "none":
            return
        none_label = next(label for label in labels if label.lower() == "none")
        _apply_option(driver, el, none_label)
        log.info("Manufacturer's Country left as None")
        return


def _apply_option(driver: WebDriver, el: WebElement, label: str) -> None:
    Select(el).select_by_visible_text(label)
    driver.execute_script(
        """
        const sel = arguments[0];
        const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set;
        setter.call(sel, sel.value);
        sel.dispatchEvent(new Event('input', {bubbles:true}));
        sel.dispatchEvent(new Event('change', {bubbles:true}));
        """,
        el,
    )


def _select_kind(labels: list[str]) -> str:
    low = {label.lower() for label in labels}
    if "united states" in low and "none" in low:
        return "country"
    if "i have more than 1" in low or "this is a unique item" in low:
        return "inventory"
    if "excellent" in low and "fair" in low:
        return "condition"
    if "subcategory" in low:
        return "subcategory"
    if "accessories" in low or "acoustic guitars" in low or "electric guitars" in low:
        return "category"
    return ""


def _exact_option(labels: list[str], want: str) -> str:
    target = want.strip().lower()
    if len(target) < 3:
        return ""
    exact = next((label for label in labels if label.lower() == target), "")
    if exact:
        return exact
    return next((label for label in labels if label.lower().startswith(target)), "")


def choose_in_select(driver: WebDriver, option_text: str, *, kind: str, force: bool = False) -> str:
    menus = [(el, labels) for el, labels in _select_options(driver) if _select_kind(labels) == kind]
    if not menus:
        return ""
    el, labels = menus[0]
    match = _exact_option(labels, option_text)
    if not match:
        return ""
    current = ""
    try:
        current = (Select(el).first_selected_option.text or "").strip()
    except Exception:
        current = ""
    if current.lower() == match.lower():
        return match
    if not force and current and not re.search(r"select one|subcategory", current, re.I):
        return ""
    try:
        _apply_option(driver, el, match)
    except Exception:
        return ""
    return match


def category_aliases(name: str) -> list[str]:
    names = [name]
    low = name.lower()
    if low.endswith(" kits"):
        stem = name[: -len(" kits")].strip()
        names.extend([f"{stem}s", stem])
    elif low.endswith(" kit"):
        stem = name[: -len(" kit")].strip()
        names.extend([f"{stem}s", stem])
    return names


def select_category(driver: WebDriver, category: str, subcategory: str) -> None:
    if category:
        log.info("Selecting category %s", category)
        parent = ""
        for name in category_aliases(category):
            parent = choose_in_select(driver, name, kind="category")
            if parent:
                break
        if not parent:
            log.warning("Could not choose category option %s", category)
        else:
            log.info("Chose category %s", parent)
            time.sleep(0.8)
    if subcategory:
        child = choose_in_select(driver, subcategory, kind="subcategory")
        if child:
            log.info("Chose subcategory %s", child)
        else:
            log.warning("Could not choose subcategory %s", subcategory)


def fill_step(driver: WebDriver, row: ListingRow) -> None:
    body = page_text(driver)
    log.info("Listing step: %s", (body.splitlines() or ["?"])[0][:80])
    if fill_labeled(driver, "Brand", "Unbranded"):
        pick_suggestion(driver, "Unbranded")
    elif fill_matching(driver, ("make", "brand"), "Unbranded"):
        pick_suggestion(driver, "Unbranded")
    if not fill_labeled(driver, "Model", row.title[:80]):
        fill_matching(driver, ("model",), row.title[:80])
    if not fill_labeled(driver, "Listing Title", row.title):
        fill_matching(driver, ("title",), row.title)
    if row.category or row.subcategory:
        select_category(driver, row.category, row.subcategory)
    leave_country_none(driver)
    fill_condition(driver, row)
    fill_inventory(driver, row)
    if not fill_labeled(driver, "Description", row.description[:4000]):
        fill_description(driver, row.description[:4000])
    if not fill_labeled(driver, "Price", row.price):
        fill_matching(driver, ("price", "asking"), row.price)
    if not fill_labeled(driver, "SKU", row.sku):
        fill_matching(driver, ("sku",), row.sku)
    if not row.photos_done and row.photos:
        paths = download_photos(row)
        if paths and upload_photos(driver, paths):
            row.photos_done = True


def condition_option(label: str) -> str:
    return "Brand New - Dealers and manufacturers only"


def click_choice(driver: WebDriver, text: str) -> bool:
    try:
        checked = driver.execute_script(
            """
            const want = arguments[0].toLowerCase();
            const labels = [...document.querySelectorAll('label')];
            const lab = labels.find(n => (n.innerText || '').trim().toLowerCase().includes(want));
            if (!lab) return false;
            let input = lab.htmlFor ? document.getElementById(lab.htmlFor) : null;
            if (!input) input = lab.querySelector('input[type="radio"], input[type="checkbox"]');
            if (!input) {
                const wrap = lab.closest('div') || lab.parentElement;
                input = wrap ? wrap.querySelector('input[type="radio"], input[type="checkbox"]') : null;
            }
            lab.scrollIntoView({block: 'center'});
            if (input && !input.checked) {
                input.click();
                input.dispatchEvent(new Event('input', {bubbles: true}));
                input.dispatchEvent(new Event('change', {bubbles: true}));
            } else if (!input) {
                lab.click();
            }
            return input ? !!input.checked : true;
            """,
            text,
        )
        return bool(checked)
    except Exception:
        return False


def fill_condition(driver: WebDriver, row: ListingRow) -> None:
    option = condition_option(row.condition)
    if not option:
        return
    chosen = choose_in_select(driver, option, kind="condition")
    if chosen:
        log.info("Chose condition %s", chosen)
    else:
        log.warning("Could not choose condition %s", option)
        return
    if "brand new" in chosen.lower():
        time.sleep(0.5)
        if click_choice(driver, "I'm authorized to sell Brand New"):
            log.info("Checked authorization to sell Brand New")
        else:
            log.warning("Could not check authorization to sell Brand New")


def fill_description(driver: WebDriver, text: str) -> bool:
    if not text:
        return False
    try:
        box = driver.execute_script(
            """
            const boxes = [...document.querySelectorAll('[contenteditable="true"]')];
            return boxes.find(n => n.getBoundingClientRect().height > 40) || null;
            """
        )
        if box is None:
            return False
        box.click()
        box.send_keys(Keys.CONTROL, "a")
        box.send_keys(text[:4000])
        log.info("Filled description")
        return True
    except Exception:
        return False


def fill_inventory(driver: WebDriver, row: ListingRow) -> None:
    try:
        qty = int(float(row.quantity or "1"))
    except ValueError:
        qty = 1
    if qty > 1:
        chosen = choose_in_select(driver, "I have more than 1", kind="inventory", force=True)
        if chosen:
            log.info("Chose inventory %s", chosen)
            time.sleep(0.4)
            if not fill_labeled(driver, "Number in stock", str(qty)):
                fill_matching(driver, ("number in stock", "stock", "quantity"), str(qty))
    if row.upc:
        if click_choice(driver, "This item has a UPC or EAN"):
            log.info("Selected UPC or EAN")
            time.sleep(0.6)
        else:
            log.warning("Could not select This item has a UPC or EAN")
        if not fill_labeled(driver, "UPC or EAN code", row.upc):
            fill_matching(driver, ("upc", "ean", "barcode"), row.upc)


def click_button_label(driver: WebDriver, text: str, timeout: float = 2) -> bool:
    if click_text(driver, rf"^{re.escape(text)}$", timeout=timeout):
        return True
    end = time.time() + timeout
    while time.time() < end:
        try:
            hit = driver.execute_script(
                """
                const want = arguments[0].toLowerCase();
                const nodes = [...document.querySelectorAll('button, a, [role="button"]')];
                const el = nodes.find(n => (n.innerText || '').trim().toLowerCase() === want && n.getBoundingClientRect().height > 2);
                if (!el) return false;
                el.scrollIntoView({block: 'center'});
                el.click();
                return true;
                """,
                text,
            )
            if hit:
                time.sleep(0.6)
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def apply_bump(driver: WebDriver) -> bool:
    body = page_text(driver).lower()
    if "promote your listing" not in body and "save and apply" not in body:
        return False
    if click_button_label(driver, "Save and apply", timeout=3):
        log.info("Clicked Save and apply")
        time.sleep(1)
        return True
    log.warning("Promote popup was open but Save and apply was not clicked")
    return False


def fill_shipping(driver: WebDriver, row: ListingRow) -> None:
    body = page_text(driver).lower()
    if "offer free shipping" not in body:
        return
    if row.free_shipping:
        if click_choice(driver, "Offer free shipping"):
            log.info("Selected Offer free shipping")
        else:
            log.warning("Could not select Offer free shipping")
        time.sleep(0.4)


def publish_ready(driver: WebDriver) -> bool:
    if "promote your listing" in page_text(driver).lower():
        return False
    try:
        return bool(
            driver.execute_script(
                """
                const nodes = [...document.querySelectorAll('button, a, [role="button"]')];
                return nodes.some(n => (n.innerText || '').trim().toLowerCase() === 'publish' && n.getBoundingClientRect().height > 2);
                """
            )
        )
    except Exception:
        return False


def save_listing(driver: WebDriver, publish: bool) -> bool:
    if not publish_ready(driver):
        return False
    if not publish:
        log.info("Draft only — leaving this listing unpublished")
        return True
    if click_button_label(driver, "Publish", timeout=3):
        log.info("Clicked Publish")
        return True
    return False


def advance(driver: WebDriver) -> bool:
    return click_button_label(driver, "Continue", timeout=2) or click_text(
        driver, r"^next$|^done$", timeout=1
    )


def create_one(driver: WebDriver, row: ListingRow) -> None:
    open_sell_form(driver)
    for _ in range(8):
        wait_for_human_check(driver)
        if apply_bump(driver):
            continue
        fill_step(driver, row)
        fill_shipping(driver, row)
        if save_listing(driver, row.publish):
            time.sleep(2)
            return
        if not advance(driver):
            break
        time.sleep(1.2)
        apply_bump(driver)
    raise RuntimeError("Sell form did not reach Save/Publish")


def mark(rows: list[ListingRow], status: str, detail: str) -> None:
    for row in rows:
        if not row.status:
            row.status = status
            row.detail = detail


def process_store(
    cfg: dict,
    chrome: Path,
    user_data: Path,
    vpncmd: Path,
    store_info: dict,
    rows: list[ListingRow],
    *,
    use_vpn: bool,
    check_ip: bool,
) -> None:
    store_name = store_info["name"]
    profile = store_info["chrome_profile"]
    expected_ip = (store_info.get("ip") or "").strip()
    vpn_account = (store_info.get("vpn") or store_name).strip()
    port = int(cfg.get("debug_port") or 9222)
    driver: Optional[WebDriver] = None
    keep_chrome_open = False

    log.info("======== %s | Chrome %s | VPN %s | IP %s ========", store_name, profile, vpn_account, expected_ip or "(none)")

    if use_vpn:
        disconnect_all_vpns(vpncmd, cfg)
        time.sleep(2)
        if not connect_store_vpn(vpncmd, vpn_account) or not wait_for_vpn(vpncmd, vpn_account):
            mark(rows, "FAILED", f"VPN {vpn_account} did not connect")
            return

    try:
        close_chrome()
        profile_dir = automation_user_data(user_data, profile)
        start_chrome(chrome, profile_dir, profile, port)
        driver = connect_selenium(port)
    except Exception as exc:
        log.exception("Chrome launch failed for %s", store_name)
        mark(rows, "FAILED", f"Chrome launch failed: {exc}")
        close_chrome()
        return

    try:
        if check_ip:
            ok, shown = wait_for_expected_ip(driver, expected_ip) if expected_ip else (False, read_ip(driver))
            if not expected_ip:
                msg = f"No expected IP in stores.json (detected {shown or 'unknown'}) — skipped Reverb"
                log.warning("%s: %s", store_name, msg)
                mark(rows, "SKIPPED", msg)
                return
            if not ok:
                msg = f"IP mismatch: expected {expected_ip}, Chrome shows {shown or 'unknown'}"
                log.error("%s: %s", store_name, msg)
                screenshot(driver, f"{store_name}_ip_mismatch")
                mark(rows, "FAILED", msg)
                return

        for i, row in enumerate(rows, start=1):
            log.info("[%s %s/%s] %s | %s", store_name, i, len(rows), row.sku, row.title[:60])
            if not row.sku or not row.title:
                row.status = "FAILED"
                row.detail = "SKU and title are required"
                continue
            try:
                create_one(driver, row)
                row.status = "OK"
                row.detail = "Published" if row.publish else "Saved draft"
                log.info("OK %s", row.sku)
            except Exception as exc:
                text = str(exc)
                if "Captcha was not solved" in text or "Cloudflare" in text:
                    keep_chrome_open = True
                    row.status = "FAILED"
                    row.detail = text
                    log.error("FAIL %s: %s", row.sku, exc)
                    log.error("Stopped this store. The Chrome tabs were left open.")
                    break
                if "not signed in" in text:
                    log.warning("Leaving this Chrome window open. Sign in there. The tabs will not be closed.")
                    keep_chrome_open = True
                    try:
                        let_user_finish_login(chrome, profile_dir, profile, port)
                        create_one(driver, row)
                        row.status = "OK"
                        row.detail = "Published" if row.publish else "Saved draft"
                        log.info("OK %s", row.sku)
                        keep_chrome_open = False
                        continue
                    except Exception as exc2:
                        exc = exc2
                        text = str(exc)
                        keep_chrome_open = True
                        if "Captcha was not solved" in text or "Cloudflare" in text:
                            row.status = "FAILED"
                            row.detail = text
                            log.error("FAIL %s: %s", row.sku, exc)
                            log.error("Stopped this store. The Chrome tabs were left open.")
                            break
                row.status = "FAILED"
                row.detail = str(exc)
                log.error("FAIL %s: %s", row.sku, exc)
                if driver and not keep_chrome_open:
                    screenshot(driver, f"{store_name}_{row.sku}")
                if "not signed in" in str(exc) or "Captcha" in str(exc) or "Cloudflare" in str(exc):
                    keep_chrome_open = True
            time.sleep(ROW_DELAY)
    finally:
        if keep_chrome_open:
            log.error("Chrome left open so you can finish the check. The tabs were not closed.")
        else:
            try:
                if driver:
                    driver.quit()
            except Exception:
                pass
            close_chrome()
        if use_vpn and not keep_chrome_open:
            try:
                run_vpncmd(vpncmd, ["AccountDisconnect", vpn_account], timeout=20)
            except Exception:
                pass


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Create Reverb listings from Excel using Selenium")
    p.add_argument("excel", nargs="?", default=str(DEFAULT_EXCEL), help="Excel file path")
    p.add_argument("--store", help="Process only this store name")
    p.add_argument("--limit", type=int, default=0, help="Max rows per store (0 = all)")
    p.add_argument("--no-vpn", action="store_true")
    p.add_argument("--skip-ip-check", action="store_true")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    excel_path = Path(args.excel)
    if not excel_path.is_file():
        log.error("Excel not found: %s", excel_path)
        return 1

    cfg = _load_config()
    chrome, user_data, vpncmd = resolve_paths(cfg)
    log.info("Chrome: %s", chrome)
    log.info("User Data: %s", user_data)
    log.info("vpncmd: %s", vpncmd)

    rows = read_excel(excel_path)
    if args.store:
        only = args.store.strip().upper()
        rows = [r for r in rows if r.store == only]
    if not rows:
        log.error("No rows to process")
        return 1

    grouped: dict[str, list[ListingRow]] = {}
    for row in rows:
        grouped.setdefault(row.store, []).append(row)
    if args.limit and args.limit > 0:
        grouped = {k: v[: args.limit] for k, v in grouped.items()}
        rows = [r for group in grouped.values() for r in group]

    log.info("Loaded %s rows across %s store(s). Chrome will close while profiles switch.", len(rows), len(grouped))
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    results_path = OUTPUT_DIR / f"listing_results_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

    for store_name, store_rows in grouped.items():
        info = store_lookup(cfg, store_name)
        if not info:
            mark(store_rows, "FAILED", f"Unknown store '{store_name}'")
            continue
        process_store(
            cfg,
            chrome,
            user_data,
            vpncmd,
            info,
            store_rows,
            use_vpn=not args.no_vpn,
            check_ip=not args.skip_ip_check,
        )
        write_results(rows, results_path)

    write_results(rows, results_path)
    ok = sum(1 for r in rows if r.status == "OK")
    skipped = sum(1 for r in rows if r.status == "SKIPPED")
    failed = sum(1 for r in rows if r.status == "FAILED")
    log.info("Done. OK=%s  SKIPPED=%s  FAILED=%s", ok, skipped, failed)
    log.info("Results: %s", results_path)
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log.error("Stopped by user")
        raise SystemExit(130)
