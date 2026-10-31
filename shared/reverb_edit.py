"""Shared Reverb edit session for the SKU and price/stock bots."""
from __future__ import annotations

import re
import time
from datetime import datetime
from pathlib import Path
from typing import Callable
import sys
from openpyxl import Workbook
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver

_PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_PROJECT / "create_listing"))

import create_listing as cl

LISTINGS_URL = "https://reverb.com/my/selling/listings"
SELLING_URL = LISTINGS_URL
ROW_DELAY = 2.5
STATUS_FILTERS = ("Live", "Draft", "Ended", "Out of Stock")


def _keyword_box(driver: WebDriver):
    return driver.execute_script(
        """
        return document.querySelector('#rc-accordion__keywordSearch-accordion input')
          || document.querySelector('input[aria-label="Search your listings"]')
          || document.querySelector('input[placeholder="Search your listings"]')
          || document.querySelector('input.search-input-group__input');
        """
    )


def _choose_status(driver: WebDriver, status: str) -> None:
    driver.execute_script(
        """
        const want = arguments[0].toLowerCase();
        const idMap = {
          'live': 'status-radio-group--live',
          'draft': 'status-radio-group--draft',
          'ended': 'status-radio-group--ended',
          'out of stock': 'status-radio-group--sold_out',
        };
        const id = idMap[want];
        if (id) {
          const forLabel = document.querySelector('label[for="' + id + '"]');
          const input = document.getElementById(id);
          const target = forLabel || input;
          if (target) {
            target.scrollIntoView({block: 'center'});
            target.click();
            return;
          }
        }
        const labels = [...document.querySelectorAll('label')];
        const label = labels.find(el => {
          const text = (el.innerText || '').trim().toLowerCase();
          return text.startsWith(want) && text.length < 40;
        });
        if (!label) return;
        label.scrollIntoView({block: 'center'});
        label.click();
        """,
        status,
    )


def _keyword_search(driver: WebDriver, sku: str) -> None:
    box = None
    deadline = time.time() + 15
    while time.time() < deadline:
        box = _keyword_box(driver)
        if box is not None:
            break
        time.sleep(0.5)
    if box is None:
        raise RuntimeError("Could not find the listings Keyword Search")
    placeholder = (box.get_attribute("placeholder") or "").strip()
    if "select" in placeholder.lower() or "shop for" in placeholder.lower():
        raise RuntimeError("The listings Keyword Search was not the box that was selected")
    cl.log.info("Using Keyword Search box%s", f" ({placeholder})" if placeholder else "")
    cl.fill_element(box, sku)
    clicked = driver.execute_script(
        """
        const box = arguments[0];
        const field = box.getBoundingClientRect();
        let root = box.parentElement;
        for (let i = 0; i < 3 && root; i++) {
          const button = [...root.querySelectorAll('button, [role="button"]')].find(el => {
            const rect = el.getBoundingClientRect();
            const sameRow = Math.abs((rect.top + rect.height / 2) - (field.top + field.height / 2)) < 24;
            return sameRow && rect.width > 8 && rect.width < 80 && rect.left >= field.left;
          });
          if (button) {
            button.click();
            return true;
          }
          root = root.parentElement;
        }
        return false;
        """,
        box,
    )
    if not clicked:
        box.send_keys(Keys.ENTER)
    time.sleep(2)


def _sku_on_card(driver: WebDriver, sku: str) -> bool:
    return bool(re.search(rf"SKU:\s*{re.escape(sku)}\b", cl.page_text(driver), re.I))


def search_listing(driver: WebDriver, sku: str) -> None:
    driver.get(LISTINGS_URL)
    time.sleep(2)
    cl.wait_for_human_check(driver)
    if cl.on_auth_page(driver):
        raise RuntimeError("Reverb is not signed in")
    if re.search(r"page isn't available|sorry, this page", cl.page_text(driver), re.I):
        raise RuntimeError("Reverb listings page did not open")
    for status in STATUS_FILTERS:
        cl.log.info("Searching %s listings for %s", status, sku)
        _choose_status(driver, status)
        time.sleep(1.2)
        _keyword_search(driver, sku)
        cl.wait_for_human_check(driver)
        if _sku_on_card(driver, sku):
            cl.log.info("Found %s under %s", sku, status)
            return
    raise RuntimeError("No listings found for this SKU")


def _open_listing_card(driver: WebDriver, sku: str) -> None:
    opened = driver.execute_script(
        """
        const sku = arguments[0].toLowerCase();
        const line = [...document.querySelectorAll('div, span, p, li, a')].find(el => {
          const text = (el.innerText || '').trim().toLowerCase();
          return text.startsWith('sku:') && text.includes(sku) && text.length < 90;
        });
        if (!line) return '';
        const titleLink = (root) => [...root.querySelectorAll('a')].filter(el => {
          const text = (el.innerText || '').trim();
          return el.href && /^https?:/i.test(el.href) && text.length > 8
            && !/more listing|publish|export|create new/i.test(text);
        }).sort((a, b) => b.innerText.trim().length - a.innerText.trim().length)[0];
        let node = line;
        for (let i = 0; i < 16 && node; i++) {
          const link = titleLink(node);
          if (link) {
            link.scrollIntoView({block: 'center'});
            return link.href;
          }
          node = node.parentElement;
        }
        node = line;
        for (let i = 0; i < 16 && node; i++) {
          const button = [...node.querySelectorAll('button, [role="button"]')].find(el =>
            /more listing actions/i.test((el.innerText || '').trim())
          );
          if (button) {
            button.scrollIntoView({block: 'center'});
            button.click();
            return 'menu';
          }
          node = node.parentElement;
        }
        return '';
        """,
        sku,
    ) or ""
    if opened == "menu":
        time.sleep(0.8)
        if not cl.click_text(driver, r"Edit Listing|^Edit$", timeout=4):
            raise RuntimeError("Could not open the listing")
        return
    if not opened:
        raise RuntimeError("Could not open the listing")
    cl.log.info("Opening listing %s", opened[:120])
    driver.get(opened)


def open_edit_listing(driver: WebDriver, sku: str) -> None:
    if not _sku_on_card(driver, sku):
        raise RuntimeError("No listings found for this SKU")
    _open_listing_card(driver, sku)
    time.sleep(2)
    cl.wait_for_human_check(driver)
    deadline = time.time() + 20
    while time.time() < deadline:
        body = cl.page_text(driver)
        if re.search(r"Tell us about your gear|SKU & Inventory|Number in stock|Listing price", body, re.I):
            return
        if re.search(r"\bReview\b", body) and _click_section_edit(driver, "Product Info"):
            time.sleep(2)
            continue
        if re.search(r"\bEdit\b", body) and re.search(r"Publish Listing|This is a preview|Status:", body, re.I):
            if cl.click_button_label(driver, "Edit", timeout=2) or cl.click_text(driver, r"^Edit$", timeout=2):
                time.sleep(2)
                continue
        time.sleep(0.5)
    raise RuntimeError("Edit listing page did not open")


def _click_section_edit(driver: WebDriver, section: str) -> bool:
    return bool(
        driver.execute_script(
            """
            const want = arguments[0].toLowerCase();
            const heads = [...document.querySelectorAll('h1, h2, h3, h4, div, span')].filter(el =>
              (el.innerText || '').trim().toLowerCase() === want
            );
            for (const head of heads) {
              let node = head;
              for (let i = 0; i < 8 && node; i++) {
                const edit = [...node.querySelectorAll('button, a, [role="button"]')].find(el =>
                  /^edit$/i.test((el.innerText || '').trim())
                );
                if (edit) {
                  edit.scrollIntoView({block: 'center'});
                  edit.click();
                  return true;
                }
                node = node.parentElement;
              }
            }
            return false;
            """,
            section,
        )
    )


def scroll_to_sku_inventory(driver: WebDriver) -> None:
    driver.execute_script(
        """
        const nodes = [...document.querySelectorAll('h1,h2,h3,h4,legend,div,span,p')];
        const el = nodes.find(n => /SKU\\s*[&,]\\s*Inventory/i.test(n.innerText || '') && (n.innerText || '').length < 80);
        if (el) el.scrollIntoView({block: 'center'});
        """
    )


def fill_named(driver: WebDriver, label: str, words: tuple[str, ...], value: str) -> bool:
    if cl.fill_labeled(driver, label, value):
        return True
    return cl.fill_matching(driver, words, value)


def clear_upc_exempt(driver: WebDriver) -> None:
    driver.execute_script(
        """
        const boxes = [...document.querySelectorAll('input[type=checkbox]')];
        for (const box of boxes) {
          const label = (box.labels && box.labels[0] && box.labels[0].innerText) || box.getAttribute('aria-label') || '';
          const name = box.name || '';
          if (/exempt|does not apply|upc_does_not_apply/i.test(label + ' ' + name) && box.checked) {
            box.click();
            return;
          }
        }
        """
    )


def _input_by_label(driver: WebDriver, label: str):
    return driver.execute_script(
        """
        const want = arguments[0].toLowerCase();
        const bad = (el) => {
          const blob = ((el.placeholder || '') + ' ' + (el.id || '') + ' ' + (el.getAttribute('aria-label') || '')).toLowerCase();
          return /shop for used|music gear|select one|search your listings/.test(blob);
        };
        const nodes = [...document.querySelectorAll('label, span, p, div, legend, h2, h3')];
        const labelEl = nodes.find(el => {
          const text = (el.innerText || '').trim().toLowerCase().replace(/\\*/g, '').replace(/required/g, '').trim();
          return (text === want || text.startsWith(want)) && (el.innerText || '').trim().length < want.length + 30;
        });
        if (!labelEl) return null;
        labelEl.scrollIntoView({block: 'center'});
        let node = labelEl.parentElement;
        for (let i = 0; i < 6 && node; i++) {
          const input = [...node.querySelectorAll('input, textarea')].find(el =>
            (labelEl.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING) && !bad(el)
          );
          if (input) return input;
          node = node.parentElement;
        }
        return null;
        """,
        label,
    )


def set_quantity(driver: WebDriver, quantity: str) -> None:
    qty = int(float(quantity))
    if qty == 1:
        chosen = cl.choose_in_select(driver, "This is a unique item", kind="inventory", force=True)
        if chosen:
            return
    cl.choose_in_select(driver, "I have more than 1", kind="inventory", force=True)
    time.sleep(0.4)
    shown = str(max(qty, 0))
    field = _input_by_label(driver, "Number in stock")
    if field is None:
        scroll_to_sku_inventory(driver)
        time.sleep(0.4)
        field = _input_by_label(driver, "Number in stock")
    if field is None:
        raise RuntimeError("Could not find the inventory field")
    cl.fill_element(field, shown)


def _on_review(driver: WebDriver) -> bool:
    return bool(re.search(r"(^|\n)\s*Review\s*(\n|$)", cl.page_text(driver)))


def _click_continue(driver: WebDriver) -> bool:
    clicked = driver.execute_script(
        """
        const nodes = [...document.querySelectorAll('button, [role="button"]')];
        const el = nodes.find(node => {
          const text = (node.innerText || '').trim().toLowerCase();
          const box = node.getBoundingClientRect();
          return text === 'continue' && box.height > 10 && box.width > 40;
        });
        if (!el) return false;
        el.scrollIntoView({block: 'center'});
        el.click();
        return true;
        """
    )
    if clicked:
        cl.log.info("Clicked Continue")
        time.sleep(1.5)
    return bool(clicked)


def save_changes(driver: WebDriver) -> None:
    """Live to any other status uses Publish Changes. Any other status to Live uses Save Changes."""
    published = False
    for _ in range(8):
        cl.wait_for_human_check(driver)
        if cl.click_button_label(driver, "Publish Changes", timeout=1):
            cl.log.info("Clicked Publish Changes")
            time.sleep(2)
            published = True
            break
        if cl.click_button_label(driver, "Save Changes", timeout=1) or cl.click_button_label(driver, "Update", timeout=1):
            cl.log.info("Clicked Save Changes")
            time.sleep(2)
            published = True
            break
        if _click_continue(driver):
            continue
        break
    if not published:
        raise RuntimeError("Could not find Publish Changes or Save Changes")
    body = cl.page_text(driver)
    if re.search(r"couldn't save|could not save|please fix|there was a problem", body, re.I):
        raise RuntimeError("Save may have failed")


def _open_editor_step(driver: WebDriver, name: str) -> bool:
    clicked = driver.execute_script(
        """
        const want = arguments[0].toLowerCase();
        const nodes = [...document.querySelectorAll('a, button, [role="button"], li, span, div')];
        const el = nodes.find(node => {
          const text = (node.innerText || '').trim().toLowerCase();
          const box = node.getBoundingClientRect();
          return text === want && box.height > 10 && box.height < 72 && box.width > 40 && box.width < 280 && box.left < 460;
        });
        if (!el) return false;
        el.scrollIntoView({block: 'center'});
        el.click();
        return true;
        """,
        name,
    )
    if clicked:
        cl.log.info("Opened editor step %s", name)
        time.sleep(1.2)
    return bool(clicked)


def _click_edit_near(driver: WebDriver, label: str) -> bool:
    clicked = driver.execute_script(
        """
        const want = arguments[0].toLowerCase();
        const nodes = [...document.querySelectorAll('h1, h2, h3, h4, label, span, p, div, dt')];
        const starts = nodes.filter(el => (el.innerText || '').trim().toLowerCase() === want);
        for (const start of starts) {
          let node = start;
          for (let i = 0; i < 8 && node; i++) {
            const edits = [...node.querySelectorAll('button, a, [role="button"]')].filter(el =>
              /^edit$/i.test((el.innerText || '').trim())
            );
            if (edits.length && (node.innerText || '').length < 3500) {
              edits[0].scrollIntoView({block: 'center'});
              edits[0].click();
              return true;
            }
            node = node.parentElement;
          }
        }
        return false;
        """,
        label,
    )
    if clicked:
        cl.log.info("Clicked Edit next to %s", label)
        time.sleep(1.5)
    return bool(clicked)


def update_sku(driver: WebDriver, sku: str, new_sku: str, upc: str) -> None:
    search_listing(driver, sku)
    open_edit_listing(driver, sku)
    filled = False
    for _ in range(6):
        cl.wait_for_human_check(driver)
        field = _input_by_label(driver, "SKU")
        if field is not None:
            cl.fill_element(field, new_sku)
            field.send_keys(Keys.TAB)
            cl.log.info("Filled SKU")
            filled = True
            break
        if _click_edit_near(driver, "SKU"):
            continue
        break
    if not filled:
        raise RuntimeError("Could not find the SKU field")
    if upc:
        clear_upc_exempt(driver)
        upc_field = _input_by_label(driver, "UPC or EAN code") or _input_by_label(driver, "UPC")
        if upc_field is None:
            raise RuntimeError("Could not find the UPC field")
        cl.fill_element(upc_field, upc)
        cl.log.info("Filled UPC")
    if not _click_continue(driver):
        raise RuntimeError("Could not find Continue after the SKU update")
    for _ in range(5):
        if not _click_continue(driver):
            return
        cl.wait_for_human_check(driver)


def update_price_and_stock(driver: WebDriver, sku: str, price: str, quantity: str) -> None:
    search_listing(driver, sku)
    open_edit_listing(driver, sku)
    price_left = price
    quantity_left = quantity
    for _ in range(8):
        cl.wait_for_human_check(driver)
        if quantity_left:
            field = _input_by_label(driver, "Number in stock")
            if field is not None:
                set_quantity(driver, quantity_left)
                quantity_left = ""
            elif _click_edit_near(driver, "Number in stock") or _click_edit_near(driver, "Photos & Description"):
                continue
        if price_left:
            price_field = _input_by_label(driver, "Listing price")
            if price_field is not None:
                cl.fill_element(price_field, price_left)
                price_field.send_keys(Keys.TAB)
                cl.log.info("Filled Listing price")
                price_left = ""
                if not _click_continue(driver):
                    raise RuntimeError("Could not find Continue after the price update")
                continue
            elif _click_edit_near(driver, "Pricing") or _click_edit_near(driver, "Listing price"):
                continue
        if not price_left and not quantity_left:
            save_changes(driver)
            return
        if not cl.click_button_label(driver, "Continue", timeout=2):
            break
        time.sleep(1.5)
    missing = []
    if price_left:
        missing.append("Listing price")
    if quantity_left:
        missing.append("Number in stock")
    raise RuntimeError("Could not find " + " or ".join(missing))


def _listings_title_ready(title: str, url: str) -> bool:
    if not re.search(r"^https://reverb\.com/my/selling(?:/|$|\?)", url):
        return False
    if re.search(r"/(signin|login|session)", url, re.I):
        return False
    if re.search(r"__cf_chl|just a moment|are you human|verifying|checking your browser", f"{title} {url}", re.I):
        return False
    return bool(title.strip())


def let_user_open_selling(chrome: Path, user_data: Path, profile: str, port: int) -> None:
    cl.close_chrome()
    cl.launch_for_user(chrome, user_data, profile, SELLING_URL, port)
    cl.log.warning("Chrome is open. Click Are you human if it is showing. The bot continues when the Listings page is open.")
    deadline = time.time() + cl.SIGN_IN_TIMEOUT
    last = ""
    ready_since = 0.0
    ready_key = ""
    while time.time() < deadline:
        pages = [
            tab
            for tab in cl.chrome_tabs(port)
            if tab.get("type") in (None, "page") and str(tab.get("url") or "").startswith("https://reverb.com/")
        ]
        challenge = False
        for tab in pages:
            url = str(tab.get("url") or "")
            title = str(tab.get("title") or "")
            shown = f"{title} | {url.split('?', 1)[0]}".strip()
            if shown and shown != last:
                cl.log.info("Chrome page: %s", shown[:180])
                last = shown
            if not _listings_title_ready(title, url):
                if "/my/selling" in url:
                    challenge = True
                    ready_key = ""
                continue
            key = f"{title}|{url.split('?', 1)[0]}"
            if key != ready_key:
                ready_key = key
                ready_since = time.time()
                continue
            if time.time() - ready_since >= 3:
                cl.log.info("Listings page is open. Continuing.")
                return
        if challenge:
            ready_key = ""
        time.sleep(2)
    raise RuntimeError("Listings page did not open. Finish the Are you human check and run the script again.")


def needs_user(message: str) -> bool:
    return "not signed in" in message or "Cloudflare" in message


def run_store_rows(
    cfg: dict,
    chrome: Path,
    user_data: Path,
    vpncmd: Path,
    store_info: dict,
    rows: list,
    apply_row: Callable,
    *,
    use_vpn: bool,
    check_ip: bool,
) -> None:
    store_name = store_info["name"]
    profile = store_info["chrome_profile"]
    expected_ip = (store_info.get("ip") or "").strip()
    vpn_account = (store_info.get("vpn") or store_name).strip()
    port = int(cfg.get("debug_port") or 9222)
    driver = None
    keep_chrome_open = False
    profile_dir = None

    cl.log.info(
        "======== %s | Chrome %s | VPN %s | IP %s ========",
        store_name,
        profile,
        vpn_account,
        expected_ip or "(none)",
    )

    if use_vpn:
        cl.disconnect_all_vpns(vpncmd, cfg)
        time.sleep(2)
        if not cl.connect_store_vpn(vpncmd, vpn_account) or not cl.wait_for_vpn(vpncmd, vpn_account):
            for row in rows:
                if not row.status:
                    row.status = "FAILED"
                    row.detail = f"VPN {vpn_account} did not connect"
            return

    try:
        cl.close_chrome()
        profile_dir = cl.automation_user_data(user_data, profile)
        cl.start_chrome(chrome, profile_dir, profile, port)
        driver = cl.connect_selenium(port)
    except Exception as exc:
        cl.log.exception("Chrome launch failed for %s", store_name)
        for row in rows:
            if not row.status:
                row.status = "FAILED"
                row.detail = f"Chrome launch failed: {exc}"
        cl.close_chrome()
        return

    try:
        if check_ip:
            ok, shown = cl.wait_for_expected_ip(driver, expected_ip) if expected_ip else (False, cl.read_ip(driver))
            if not expected_ip:
                msg = f"No expected IP in stores.json (detected {shown or 'unknown'}) — skipped Reverb"
                cl.log.warning("%s: %s", store_name, msg)
                for row in rows:
                    row.status = "SKIPPED"
                    row.detail = msg
                return
            if not ok:
                msg = f"IP mismatch: expected {expected_ip}, Chrome shows {shown or 'unknown'}"
                cl.log.error("%s: %s", store_name, msg)
                cl.screenshot(driver, f"{store_name}_ip_mismatch")
                for row in rows:
                    row.status = "FAILED"
                    row.detail = msg
                return

        for index, row in enumerate(rows, start=1):
            cl.log.info("[%s %s/%s] %s", store_name, index, len(rows), row.sku)
            if row.status:
                continue
            try:
                apply_row(driver, row)
                row.status = "OK"
                row.detail = "Updated"
                cl.log.info("OK %s", row.sku)
            except Exception as exc:
                text = str(exc)
                if needs_user(text) and profile_dir is not None:
                    cl.log.warning("Automated window cannot pass Cloudflare. Handing Chrome to you.")
                    try:
                        if driver:
                            driver.quit()
                    except Exception:
                        pass
                    driver = None
                    for _ in range(2):
                        let_user_open_selling(chrome, profile_dir, profile, port)
                        driver = cl.connect_selenium(port)
                        if not cl.human_check_visible(driver):
                            break
                        cl.log.warning("Are you human is still open. Click the checkbox in Chrome.")
                        try:
                            driver.quit()
                        except Exception:
                            pass
                        driver = None
                    if driver is None or cl.human_check_visible(driver):
                        raise RuntimeError("Are you human is still open. Click the checkbox, then run the script again.")
                    try:
                        apply_row(driver, row)
                        row.status = "OK"
                        row.detail = "Updated"
                        cl.log.info("OK %s", row.sku)
                        continue
                    except Exception as exc2:
                        text = str(exc2)
                        exc = exc2
                row.status = "FAILED"
                row.detail = str(exc)
                cl.log.error("FAIL %s: %s", row.sku, exc)
                if driver:
                    cl.screenshot(driver, f"{store_name}_{row.sku}")
                if needs_user(text):
                    keep_chrome_open = True
            time.sleep(ROW_DELAY)
    finally:
        if keep_chrome_open:
            cl.log.error("Chrome left open. Sign in to Reverb there, then run the script again.")
        else:
            try:
                if driver:
                    driver.quit()
            except Exception:
                pass
            cl.close_chrome()
        if use_vpn and not keep_chrome_open:
            try:
                cl.run_vpncmd(vpncmd, ["AccountDisconnect", vpn_account], timeout=20)
            except Exception:
                pass


def write_sheet(path: Path, headers: list[str], rows: list[list]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "results"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


def result_path(prefix: str) -> Path:
    cl.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return cl.OUTPUT_DIR / f"{prefix}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
