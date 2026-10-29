#!/usr/bin/env python3
"""
Scrape AliExpress product links into Create.xlsx for create_listing.py.

Prices come from the AliExpress Drop Shipping API in AUD for Australia.
The browser path is only used when ALIEXPRESS_ACCESS_TOKEN is missing.

Urls.xlsx in this folder uses the same columns as the Amazon scraper.
This script scrapes
aliexpress.com links only. Amazon.com links stay in scrape_amazon.py.
Check Create.xlsx, then run:

  python create_listing/create_listing.py create_listing/Create.xlsx
"""
from __future__ import annotations

import hashlib
import hmac
import json
import re
import sys
import time
from pathlib import Path
from typing import Any, Optional

import requests

_PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_PROJECT / "create_listing"))
sys.path.insert(0, str(_PROJECT / "scrape_amazon"))

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from create_listing import log
from scrape_amazon import (
    CREATE_PATH,
    ROOT,
    append_create_row,
    cell_text,
    clean_brand,
    collect_photos,
    ensure_urls_file,
    header_key,
    listing_description,
    listing_quantity,
    map_reverb_category,
    read_done_urls,
    read_urls,
    start_browser,
    text_of,
)

URLS_PATH = Path(__file__).resolve().parent / "Urls.xlsx"
ALI_POSTCODE = "2762"
ALI_STATE = "New South Wales"
ALI_CITY = "Schofields"
ALI_PROFILE = ROOT / "chrome_aliexpress"
ALI_SHIP_MARKER = ALI_PROFILE / "ship_to_2762.txt"


class AliExpressCheck(RuntimeError):
    pass


ALIEXPRESS_RE = re.compile(r"https?://(?:[a-z0-9-]+\.)?aliexpress\.[a-z.]+/item/(\d+)", re.I)


def aliexpress_page_text(driver: WebDriver) -> str:
    try:
        text = driver.execute_script("return document.body ? document.body.innerText : ''")
    except Exception:
        return ""
    return text or ""


def aliexpress_traffic_blocked(driver: WebDriver) -> bool:
    return bool(re.search(r"unusual traffic|try again later", aliexpress_page_text(driver), re.I))


def aliexpress_blocked(driver: WebDriver) -> bool:
    if aliexpress_traffic_blocked(driver):
        return True
    try:
        blocked = driver.execute_script(
            """
            const text = document.body ? document.body.innerText : '';
            if (/slide to verify|please verify|are you a robot|i'm not a robot/i.test(text)) return true;
            for (const frame of document.querySelectorAll('iframe')) {
              const box = frame.getBoundingClientRect();
              const label = (frame.src || '') + ' ' + (frame.title || '');
              if (box.width > 40 && box.height > 40 && /recaptcha|captcha/i.test(label)) return true;
            }
            return false;
            """
        )
    except Exception:
        return False
    return bool(blocked)


def wait_if_aliexpress_blocked(driver: WebDriver) -> None:
    if aliexpress_traffic_blocked(driver):
        raise AliExpressCheck("AliExpress blocked this network. Close Chrome and run the scraper again later.")
    if not aliexpress_blocked(driver):
        return
    log.warning("AliExpress asked for a check. Tick I'm not a robot in the Chrome window. The scraper will wait.")
    deadline = time.time() + 600
    while time.time() < deadline:
        if aliexpress_traffic_blocked(driver):
            raise AliExpressCheck("AliExpress blocked this network. Close Chrome and run the scraper again later.")
        if not aliexpress_blocked(driver):
            log.info("AliExpress check cleared")
            time.sleep(2)
            return
        time.sleep(2)
    raise AliExpressCheck("AliExpress robot check was not finished")


AE_PDP_HOOK = r"""
(() => {
  window.__aePdp = null;
  for (let i = 1; i <= 6; i++) {
    const name = 'mtopjsonp' + i;
    let real;
    Object.defineProperty(window, name, {
      configurable: true,
      enumerable: true,
      get() {
        return function (data) {
          try {
            if (data && data.api && String(data.api).indexOf('pdp.pc.query') >= 0) window.__aePdp = data;
          } catch (e) {}
          if (typeof real === 'function') return real.apply(this, arguments);
        };
      },
      set(fn) { real = fn; }
    });
  }
})();
"""
SKIP_CRUMBS = {"home", "all categories", "aliexpress"}
SPEC_CRUMB_ORDER = ("category", "type", "product type", "style", "item type")


def install_aliexpress_hook(driver: WebDriver) -> None:
    if getattr(driver, "ae_hook_ready", False):
        return
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": AE_PDP_HOOK})
    driver.ae_hook_ready = True


def aliexpress_payload(driver: WebDriver) -> dict[str, Any]:
    try:
        payload = driver.execute_script(
            """
            const data = window.__aePdp;
            if (!data || !data.data || !data.data.result) return null;
            const result = data.data.result;
            const globalData = ((result.GLOBAL_DATA || {}).globalData) || {};
            const props = ((result.PRODUCT_PROP_PC || {}).showedProps) || [];
            const images = ((result.HEADER_IMAGE_PC || {}).imagePathList) || [];
            const price = result.PRICE || {};
            const sku = String(price.selectedSkuId || '');
            const map = price.skuIdStrPriceInfoMap || {};
            const chosen = map[sku] || Object.values(map)[0] || {};
            const formatted = String(chosen.salePriceString || '');
            let amount = formatted;
            if (!amount) {
              const parts = String(chosen.salePriceLocal || '').split('|').filter(part => part !== '');
              const digits = (value) => /^[0-9]+$/.test(value);
              if (parts.length >= 3 && digits(parts[1]) && digits(parts[2])) amount = parts[1] + '.' + parts[2];
              else if (parts.length >= 2) amount = parts[1];
              else amount = parts[0] || '';
            }
            return {
              pathList: ((globalData.breadcrumb || {}).pathList) || [],
              images,
              price: amount,
              currency: ((chosen.originalPrice || {}).currency) || price.region || '',
              props: props.map(p => ({attrName: p.attrName || '', attrValue: p.attrValue || ''})),
            };
            """
        )
    except Exception:
        payload = None
    return payload if isinstance(payload, dict) else {}


def wait_for_aliexpress_payload(driver: WebDriver) -> dict[str, Any]:
    deadline = time.time() + 20
    payload: dict[str, Any] = {}
    while time.time() < deadline:
        payload = aliexpress_payload(driver)
        if payload:
            return payload
        time.sleep(0.4)
    return payload


def clean_gallery(urls: list[Any]) -> list[str]:
    cleaned: list[str] = []
    for raw in urls:
        url = str(raw or "").replace("\\u002F", "/").replace("\\/", "/").split("?")[0].strip()
        url = re.sub(r"_\d+x\d+\.(jpg|jpeg|png|webp).*$", r".\1", url, flags=re.I)
        if not re.search(r"/kf/.+\.(?:jpg|jpeg|png|webp)$", url, re.I):
            continue
        if url not in cleaned:
            cleaned.append(url)
    return cleaned


def aliexpress_images(page: str, payload: Optional[dict[str, Any]] = None) -> list[str]:
    raw: list[Any] = list((payload or {}).get("images") or [])
    if not raw:
        page_norm = page.replace("\\u002F", "/").replace("\\/", "/")
        match = re.search(r'"imagePathList"\s*:\s*(\[[^\]]+\])', page_norm)
        if match:
            try:
                parsed = json.loads(match.group(1))
            except Exception:
                parsed = re.findall(r"https://[^\"\\]+", match.group(1))
            if isinstance(parsed, list):
                raw.extend(parsed)
        if not raw:
            raw.extend(
                re.findall(
                    r"https://[^\"\\]+\.(?:alicdn\.com|aliexpress-media\.com)/kf/[^\"\\]+",
                    page_norm,
                    re.I,
                )
            )
    return clean_gallery(raw)


def price_amount(raw: str) -> str:
    match = re.search(r"(\d[\d,]*(?:\.\d{1,2})?)", raw or "")
    if not match:
        return ""
    return match.group(1).replace(",", "")


def visible_sale_price(driver: WebDriver) -> str:
    try:
        text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""
    for line in text.splitlines():
        line = " ".join(line.split())
        if not line or len(line) > 40:
            continue
        if re.search(r"off on|\bsave\b|coupon|\beach\b|shipping|\bwas\b", line, re.I):
            continue
        match = re.search(
            r"(?:AU\$|A\$|Rs\.|PKR|US\s*\$|USD|\$)\s*(\d[\d,]*(?:\.\d{1,2})?)",
            line,
            re.I,
        )
        if match:
            return match.group(1).replace(",", "")
    return ""


def aliexpress_price(page: str, payload: Optional[dict[str, Any]] = None, driver: Optional[WebDriver] = None) -> str:
    amount = price_amount(str((payload or {}).get("price") or ""))
    if amount:
        return amount
    match = re.search(r'"salePriceString"\s*:\s*"([^"]+)"', page)
    if match:
        amount = price_amount(match.group(1))
        if amount:
            return amount
    if driver is not None:
        return visible_sale_price(driver)
    return ""


def prop_pairs(props: Optional[list[Any]]) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for prop in props or []:
        if isinstance(prop, dict):
            pairs.append((str(prop.get("attrName") or ""), str(prop.get("attrValue") or "")))
    return pairs


def aliexpress_brand(page: str, props: Optional[list[Any]] = None) -> str:
    brand = ""
    for name, value in prop_pairs(props):
        if name.strip().lower() == "brand name":
            brand = value
            break
    if not brand:
        match = re.search(r'"attrName"\s*:\s*"Brand Name"\s*,\s*"attrValue"\s*:\s*"([^"]+)"', page, re.I)
        if not match:
            match = re.search(r'"attrValue"\s*:\s*"([^"]+)"\s*,\s*"attrName"\s*:\s*"Brand Name"', page, re.I)
        brand = match.group(1) if match else ""
    brand = clean_brand(brand)
    if brand.lower() in {"none", "no brand", "unbranded"}:
        return ""
    return brand


def aliexpress_specs(page: str, props: Optional[list[Any]] = None) -> list[str]:
    pairs = prop_pairs(props)
    if not pairs:
        pairs = re.findall(
            r'"attrName"\s*:\s*"([^"]+)"\s*,\s*"attrValue"\s*:\s*"([^"]+)"',
            page,
        )
        if not pairs:
            pairs = re.findall(
                r'"attrValue"\s*:\s*"([^"]+)"\s*,\s*"attrName"\s*:\s*"([^"]+)"',
                page,
            )
            pairs = [(name, value) for value, name in pairs]
    lines: list[str] = []
    seen: set[str] = set()
    for name, value in pairs:
        name = re.sub(r"\s+", " ", name).strip()
        value = re.sub(r"\s+", " ", value).strip()
        key = name.lower()
        if not name or not value or key in seen or key in {"brand name", "choice"}:
            continue
        if value.lower() in {"none", "no", "n/a"}:
            continue
        seen.add(key)
        lines.append(f"{name}: {value}")
    return lines[:12]


def aliexpress_crumbs(payload: dict[str, Any]) -> list[str]:
    names: list[str] = []
    for item in payload.get("pathList") or []:
        if isinstance(item, dict):
            name = str(item.get("name") or "").strip()
        else:
            name = str(item or "").strip()
        if name and name.lower() not in SKIP_CRUMBS:
            names.append(name)
    if names:
        return names
    by_name: dict[str, str] = {}
    for name, value in prop_pairs(payload.get("props")):
        key = re.sub(r"\s+", " ", name).strip().lower()
        value = re.sub(r"\s+", " ", value).strip()
        if key not in SPEC_CRUMB_ORDER or not value or len(value) > 80:
            continue
        if value.lower() in {"none", "no", "yes", "other", "n/a"}:
            continue
        by_name[key] = value
    return [by_name[key] for key in SPEC_CRUMB_ORDER if key in by_name]


def aliexpress_stock(driver: WebDriver, page: str) -> int:
    text = ""
    try:
        text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        text = ""
    low = f"{text}\n{page}".lower()
    if re.search(r"out of stock|sold out|no longer available", low):
        return 0
    left = re.search(r"only\s+(\d+)\s+left|(\d+)\s+pieces available", low)
    if left:
        return int(left.group(1) or left.group(2))
    try:
        options = driver.execute_script(
            """
            const sel = document.querySelector('select, input[class*="quantity"]');
            if (sel && sel.tagName === 'SELECT') {
              return [...sel.options].map(o => parseInt(o.value, 10)).filter(n => !Number.isNaN(n));
            }
            const max = sel ? parseInt(sel.max || sel.getAttribute('max') || '', 10) : NaN;
            return Number.isNaN(max) ? [] : [max];
            """
        )
    except Exception:
        options = []
    if options:
        return max(int(value) for value in options)
    if re.search(r"in stock|available|add to cart|buy now", low):
        return 11
    return 0


def _visible_select_popup(driver: WebDriver):
    popups = driver.find_elements(By.CSS_SELECTOR, "[class*='select--popup']")
    visible = []
    for popup in popups:
        try:
            box = popup.rect
        except Exception:
            continue
        if box.get("height", 0) > 20 and box.get("width", 0) > 20 and popup.is_displayed():
            visible.append(popup)
    return visible[-1] if visible else None


def _choose_ship_value(driver: WebDriver, index: int, query: str) -> None:
    wrap = driver.find_element(By.CSS_SELECTOR, "[class*='es--contentWrap']")
    current = " ".join(wrap.find_elements(By.CSS_SELECTOR, "[class*='select--text']")[index].text.split())
    if current.lower() == query.lower():
        return
    wrap.find_elements(By.CSS_SELECTOR, "[class*='select--text']")[index].click()
    time.sleep(0.5)
    popup = _visible_select_popup(driver)
    if popup is None:
        raise RuntimeError(f"Ship-to list for {query} did not open")
    field = popup.find_element(By.CSS_SELECTOR, "input")
    field.click()
    field.send_keys(Keys.CONTROL, "a")
    field.send_keys(Keys.BACKSPACE)
    field.send_keys(query)
    deadline = time.time() + 8
    while time.time() < deadline:
        popup = _visible_select_popup(driver)
        if popup is None:
            time.sleep(0.3)
            continue
        match = None
        for item in popup.find_elements(By.CSS_SELECTOR, "[class*='select--item']"):
            label = " ".join(item.text.split())
            if label.lower() == query.lower():
                match = item
                break
        if match is not None:
            driver.execute_script("arguments[0].click();", match)
            time.sleep(1.2)
            return
        time.sleep(0.3)
    raise RuntimeError(f"Could not choose {query}")


def _apply_aliexpress_postcode(driver: WebDriver) -> None:
    menu = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[class*='ship-to--menuItem']"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu)
    ActionChains(driver).move_to_element(menu).click(menu).perform()
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "[class*='es--contentWrap']"))
    )
    time.sleep(0.4)
    _choose_ship_value(driver, 0, "Australia")
    time.sleep(0.8)
    _choose_ship_value(driver, 1, ALI_STATE)
    time.sleep(1.2)
    _choose_ship_value(driver, 2, ALI_CITY)
    root = driver.find_element(By.CSS_SELECTOR, "[class*='es--contentWrap']")
    saved = driver.execute_script(
        """
        const root = arguments[0];
        const button = [...root.querySelectorAll('button, [role="button"], div, span')].find(el =>
          (el.innerText || '').trim() === 'Save'
        );
        if (!button) return false;
        button.click();
        return true;
        """,
        root,
    )
    if not saved:
        raise RuntimeError("Ship-to Save button was not found")
    time.sleep(1.5)


def set_aliexpress_ship_to(driver: WebDriver) -> None:
    if getattr(driver, "ae_ship_ready", False):
        return
    install_aliexpress_hook(driver)
    driver.execute_cdp_cmd(
        "Network.setCookie",
        {
            "name": "aep_usuc_f",
            "value": "site=glo&c_tp=AUD&region=AU&b_locale=en_US",
            "domain": ".aliexpress.com",
            "path": "/",
            "url": "https://www.aliexpress.com/",
        },
    )
    driver.get("https://www.aliexpress.com/")
    time.sleep(2)
    wait_if_aliexpress_blocked(driver)
    if ALI_SHIP_MARKER.is_file():
        log.info("AliExpress ship-to already Australia %s", ALI_POSTCODE)
        driver.ae_ship_ready = True
        return
    for attempt in range(2):
        try:
            _apply_aliexpress_postcode(driver)
            ALI_SHIP_MARKER.write_text(ALI_POSTCODE, encoding="utf-8")
            log.info("AliExpress ship-to Australia %s (%s, %s)", ALI_POSTCODE, ALI_STATE, ALI_CITY)
            break
        except AliExpressCheck:
            raise
        except Exception as exc:
            if aliexpress_blocked(driver):
                wait_if_aliexpress_blocked(driver)
                continue
            log.warning(
                "AliExpress postcode %s was not set (%s). Prices stay in AUD for Australia.",
                ALI_POSTCODE,
                exc,
            )
            break
    driver.ae_ship_ready = True


class AliExpressAPIError(RuntimeError):
    pass


IOP_GATEWAY = "https://api-sg.aliexpress.com"
DS_PRODUCT_METHOD = "aliexpress.ds.product.get"


def env_values() -> dict[str, str]:
    path = ROOT / ".env"
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def access_token_from_env(raw: str) -> str:
    for part in reversed(raw.split()):
        if re.fullmatch(r"[A-Za-z0-9_-]{20,}", part):
            return part
    return raw.strip()


def ali_api_config() -> dict[str, str]:
    values = env_values()
    token = access_token_from_env(values.get("ALIEXPRESS_ACCESS_TOKEN", ""))
    app_key = values.get("ALIEXPRESS_APP_KEY", "").strip()
    app_secret = values.get("ALIEXPRESS_APP_SECRET", "").strip()
    if not (token and app_key and app_secret):
        return {}
    return {"key": app_key, "secret": app_secret, "token": token}


def _sorted_params(params: dict[str, Any]) -> list[tuple[str, str]]:
    return sorted(
        (key, str(value))
        for key, value in params.items()
        if key != "sign" and value is not None and str(value) != ""
    )


def _sign(base: str, secret: str) -> str:
    digest = hmac.new(secret.encode("utf-8"), base.encode("utf-8"), hashlib.sha256)
    return digest.hexdigest().upper()


def _iop_error(data: Any) -> str:
    if not isinstance(data, dict):
        return "AliExpress API returned an unreadable response"
    if isinstance(data.get("error_response"), dict):
        err = data["error_response"]
        code = err.get("code") or err.get("type") or ""
        msg = err.get("msg") or err.get("message") or "API error"
        return f"{code}: {msg}".strip(": ")
    code = data.get("code") or data.get("error_code") or ""
    if str(code) not in {"", "0", "200"}:
        msg = data.get("message") or data.get("msg") or data.get("error_msg") or "API error"
        return f"{code}: {msg}".strip(": ")
    return ""


def _post_iop(url: str, params: dict[str, str], *, as_query: bool) -> dict[str, Any]:
    try:
        if as_query:
            response = requests.post(url, params=params, timeout=30)
        else:
            response = requests.post(url, data=params, timeout=30)
    except requests.RequestException as exc:
        raise AliExpressAPIError(f"AliExpress API request failed: {exc}") from exc
    try:
        data = response.json()
    except ValueError as exc:
        raise AliExpressAPIError(f"AliExpress API returned HTTP {response.status_code}") from exc
    if response.status_code >= 400:
        raise AliExpressAPIError(_iop_error(data) or f"AliExpress API HTTP {response.status_code}")
    err = _iop_error(data)
    if err:
        raise AliExpressAPIError(err)
    return data


def _token_rejected(message: str) -> bool:
    return bool(re.search(r"access[ _]?token|illegalaccesstoken|invalid.?session|token.{0,20}expired", message, re.I))


def fetch_ds_product(cfg: dict[str, str], product_id: str) -> dict[str, Any]:
    business = {
        "product_id": product_id,
        "ship_to_country": "AU",
        "target_currency": "AUD",
        "target_language": "EN",
    }
    common = {
        "app_key": cfg["key"],
        "sign_method": "sha256",
        "timestamp": str(int(time.time() * 1000)),
        "method": DS_PRODUCT_METHOD,
        "session": cfg["token"],
        "simplify": "true",
    }
    common.update(business)
    sync_params = dict(common)
    sync_params["sign"] = _sign("".join(f"{key}{value}" for key, value in _sorted_params(sync_params)), cfg["secret"])
    errors: list[str] = []
    for label, call in (
        ("sync", lambda: _post_iop(f"{IOP_GATEWAY}/sync", sync_params, as_query=True)),
        ("sync-form", lambda: _post_iop(f"{IOP_GATEWAY}/sync", sync_params, as_query=False)),
    ):
        try:
            data = call()
            result = _ds_result(data)
            if result:
                return result
            errors.append(f"{label}: product was not in the response")
        except AliExpressAPIError as exc:
            errors.append(f"{label}: {exc}")
            if _token_rejected(str(exc)):
                raise AliExpressAPIError(
                    "AliExpress rejected the access token. Connect AliExpress again in SellerPilot Hub, then paste only the access token into .env."
                ) from exc
    rest = {
        "app_key": cfg["key"],
        "sign_method": "sha256",
        "timestamp": str(int(time.time() * 1000)),
        "access_token": cfg["token"],
        "format": "json",
        "v": "2.0",
        "simplify": "true",
    }
    rest.update(business)
    path = "/aliexpress/ds/product/get"
    rest["sign"] = _sign(path + "".join(f"{key}{value}" for key, value in _sorted_params(rest)), cfg["secret"])
    try:
        data = _post_iop(f"{IOP_GATEWAY}/rest{path}", rest, as_query=False)
    except AliExpressAPIError as exc:
        if _token_rejected(str(exc)):
            raise AliExpressAPIError(
                "AliExpress rejected the access token. Connect AliExpress again in SellerPilot Hub, then paste only the access token into .env."
            ) from exc
        errors.append(f"rest: {exc}")
        raise AliExpressAPIError("; ".join(errors)) from exc
    result = _ds_result(data)
    if not result:
        raise AliExpressAPIError("AliExpress did not return this product")
    return result


def _ds_result(data: dict[str, Any]) -> dict[str, Any]:
    for key in ("aliexpress_ds_product_get_response", "aliexpressDsProductGetResponse"):
        root = data.get(key)
        if isinstance(root, dict) and isinstance(root.get("result"), dict):
            return root["result"]
    result = data.get("result")
    return result if isinstance(result, dict) else {}


def _walk_numbers(node: Any, keys: set[str]) -> list[float]:
    found: list[float] = []

    def walk(value: Any) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                if key in keys:
                    amount = price_amount(str(child or ""))
                    if amount:
                        number = float(amount)
                        if number > 0:
                            found.append(number)
                elif isinstance(child, (dict, list)):
                    walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(node)
    return found


def ds_price(result: dict[str, Any]) -> str:
    skus = result.get("ae_item_sku_info_dtos") or result.get("aeItemSkuInfoDtos") or {}
    amounts = _walk_numbers(
        skus,
        {
            "offer_sale_price",
            "offerSalePrice",
            "sku_price",
            "skuPrice",
            "offer_bulk_sale_price",
            "offerBulkSalePrice",
        },
    )
    if not amounts:
        base = result.get("ae_item_base_info_dto") or result.get("aeItemBaseInfoDto") or {}
        amounts = _walk_numbers(base, {"product_min_price", "productMinPrice", "target_sale_price", "targetSalePrice"})
    if not amounts:
        return ""
    return f"{min(amounts):.2f}"


def ds_stock(result: dict[str, Any]) -> int:
    total = 0
    found = False

    def walk(value: Any) -> None:
        nonlocal total, found
        if isinstance(value, dict):
            for key in ("sku_available_stock", "skuAvailableStock", "ipm_sku_stock", "ipmSkuStock"):
                raw = value.get(key)
                if raw in (None, ""):
                    continue
                try:
                    total += max(0, int(float(str(raw))))
                    found = True
                except ValueError:
                    pass
                break
            for child in value.values():
                if isinstance(child, (dict, list)):
                    walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(result.get("ae_item_sku_info_dtos") or result.get("aeItemSkuInfoDtos") or {})
    if found:
        return total
    base = result.get("ae_item_base_info_dto") or result.get("aeItemBaseInfoDto") or {}
    status = str(base.get("product_status_type") or base.get("productStatusType") or "").lower()
    if status == "onselling":
        return 11
    return 0


def ds_props(result: dict[str, Any]) -> list[dict[str, str]]:
    bag = result.get("ae_item_properties") or result.get("aeItemProperties") or {}
    pairs: list[dict[str, str]] = []

    def walk(value: Any) -> None:
        if isinstance(value, dict):
            name = str(value.get("attr_name") or value.get("attrName") or "").strip()
            attr_value = str(value.get("attr_value") or value.get("attrValue") or "").strip()
            if name and attr_value:
                pairs.append({"attrName": name, "attrValue": attr_value})
            for child in value.values():
                if isinstance(child, (dict, list)):
                    walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(bag)
    return pairs


def ds_images(result: dict[str, Any]) -> list[str]:
    media = result.get("ae_multimedia_info_dto") or result.get("aeMultimediaInfoDto") or {}
    raw = ""
    if isinstance(media, dict):
        raw = str(media.get("image_urls") or media.get("imageUrls") or "")
    urls = []
    for part in re.split(r"[;\s]+", raw):
        part = part.strip()
        if part.startswith("//"):
            part = "https:" + part
        if part.startswith("http"):
            urls.append(part)
    return clean_gallery(urls)


def scrape_aliexpress_api(cfg: dict[str, str], item: dict[str, str], url: str) -> dict[str, str]:
    match = ALIEXPRESS_RE.search(url)
    if not match:
        raise AliExpressAPIError("AliExpress product id was not found in the URL")
    product_id = match.group(1)
    log.info("AliExpress API %s", url)
    result = fetch_ds_product(cfg, product_id)
    base = result.get("ae_item_base_info_dto") or result.get("aeItemBaseInfoDto") or {}
    title = str(base.get("subject") or "").strip()
    if not title:
        raise AliExpressAPIError("Product title was not found")
    props = ds_props(result)
    crumbs = aliexpress_crumbs({"props": props})
    category, subcategory = map_reverb_category(crumbs)
    if category:
        log.info("AliExpress categories %s -> %s / %s", " > ".join(crumbs), category, subcategory or "(none)")
    else:
        category = item.get("category", "")
        subcategory = item.get("subcategory", "")
        log.warning("No Reverb category for %s", " > ".join(crumbs) or "(no breadcrumb)")
    photos = collect_photos(ds_images(result), item.get("sku") or product_id)
    log.info("Kept %s photos", len(photos))
    price = ds_price(result)
    stock = ds_stock(result)
    quantity = listing_quantity(stock)
    log.info("AliExpress stock %s mapped to quantity %s", stock, quantity)
    log.info("AliExpress price %s AUD", price or "(none)")
    return {
        "store_name": item["store_name"],
        "title": title[:140],
        "description": listing_description(title, aliexpress_specs("", props)) or title,
        "remove_branding": aliexpress_brand("", props),
        "price": price,
        "quantity": quantity,
        "Category": category,
        "Subcategory": subcategory,
        "sku": item.get("sku") or product_id,
        "photos": "\n".join(photos),
        "status": item["status"] or "Draft",
        "upc": item.get("upc") or "",
        "free_shipping": item["free_shipping"] or "False",
        "amazon_url": url,
    }


def scrape_aliexpress(driver: WebDriver, item: dict[str, str], url: str) -> dict[str, str]:
    log.info("Scraping AliExpress %s", url)
    set_aliexpress_ship_to(driver)
    driver.get(url)
    time.sleep(3)
    wait_if_aliexpress_blocked(driver)
    WebDriverWait(driver, 25).until(lambda d: bool(text_of(d, "h1") or text_of(d, "[data-pl='product-title']")))
    payload = wait_for_aliexpress_payload(driver)
    page = driver.page_source
    title = text_of(driver, "h1") or text_of(driver, "[data-pl='product-title']")
    if not title:
        raise RuntimeError("Product title was not found")
    landed = driver.current_url or url
    item_id = ""
    match = ALIEXPRESS_RE.search(landed) or ALIEXPRESS_RE.search(url)
    if match:
        item_id = match.group(1)
    sku = item.get("sku") or item_id
    upc = item.get("upc") or ""
    stock = aliexpress_stock(driver, page)
    quantity = listing_quantity(stock)
    log.info("AliExpress stock %s mapped to quantity %s", stock, quantity)
    props = payload.get("props") or []
    crumbs = aliexpress_crumbs(payload)
    category, subcategory = map_reverb_category(crumbs)
    if category:
        log.info("AliExpress categories %s -> %s / %s", " > ".join(crumbs), category, subcategory or "(none)")
    else:
        category = item.get("category", "")
        subcategory = item.get("subcategory", "")
        log.warning("No Reverb category for %s", " > ".join(crumbs) or "(no breadcrumb)")
    photos = collect_photos(aliexpress_images(page, payload), sku or "photo")
    log.info("Kept %s photos", len(photos))
    price = aliexpress_price(page, payload, driver)
    log.info("AliExpress price %s %s", price or "(none)", payload.get("currency") or "")
    return {
        "store_name": item["store_name"],
        "title": title[:140],
        "description": listing_description(title, aliexpress_specs(page, props)) or title,
        "remove_branding": aliexpress_brand(page, props),
        "price": price,
        "quantity": quantity,
        "Category": category,
        "Subcategory": subcategory,
        "sku": sku,
        "photos": "\n".join(photos),
        "status": item["status"] or "Draft",
        "upc": upc,
        "free_shipping": item["free_shipping"] or "False",
        "amazon_url": url,
    }


def facts_from_description(text: str) -> list[str]:
    facts: list[str] = []
    heading = ""
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("-"):
            body = line.lstrip("-").strip()
            if heading:
                facts.append(f"{heading}: {body}")
                heading = ""
            elif body:
                facts.append(body)
        else:
            heading = line
    if heading:
        facts.append(heading)
    return facts


def description_needs_points(text: str) -> bool:
    content = [line.strip() for line in (text or "").splitlines() if line.strip()]
    points = [line for line in content if line.startswith("-")]
    if len(points) < 4:
        return True
    return any(not line.startswith("-") for line in content)


def rewrite_saved_descriptions(path) -> int:
    if not path.is_file():
        return 0
    from openpyxl import load_workbook

    wb = load_workbook(path)
    ws = wb.active
    headers = [header_key(cell.value) for cell in ws[1]]
    try:
        url_col = headers.index("amazonurl")
        title_col = headers.index("title")
        desc_col = headers.index("description")
    except ValueError:
        return 0
    changed = 0
    for excel_row in range(2, ws.max_row + 1):
        url = cell_text(ws.cell(excel_row, url_col + 1).value)
        if "aliexpress." not in url.lower():
            continue
        current = cell_text(ws.cell(excel_row, desc_col + 1).value)
        if not description_needs_points(current):
            continue
        title = cell_text(ws.cell(excel_row, title_col + 1).value)
        ws.cell(excel_row, desc_col + 1).value = listing_description(title, facts_from_description(current))
        changed += 1
    if changed:
        wb.save(path)
        log.info("Rewrote %s descriptions with at least 4 points", changed)
    return changed


def ali_needs_again(row: tuple[Any, ...], index: dict[str, int]) -> bool:
    def take(name: str) -> str:
        pos = index.get(name)
        if pos is None or pos >= len(row):
            return ""
        return cell_text(row[pos])

    if "aliexpress." not in take("amazonurl").lower():
        return False
    return not (take("category") and take("photos") and take("price"))


def main() -> None:
    ensure_urls_file(URLS_PATH)
    try:
        rewrite_saved_descriptions(CREATE_PATH)
    except PermissionError:
        log.warning("Close Create.xlsx so descriptions can be rewritten, then run this script again.")
    items = [item for item in read_urls(URLS_PATH) if ALIEXPRESS_RE.search(item["amazon_url"])]
    if not items:
        raise SystemExit("Urls.xlsx has no AliExpress links")
    done = read_done_urls(CREATE_PATH, ali_needs_again)
    pending = [item for item in items if item["amazon_url"] not in done]
    if not pending:
        log.info("%s already has every AliExpress link. Check it, then run create_listing/create_listing.py", CREATE_PATH.name)
        return
    cfg = ali_api_config()
    if cfg:
        failed: list[str] = []
        for item in pending:
            try:
                append_create_row(CREATE_PATH, scrape_aliexpress_api(cfg, item, item["amazon_url"]))
                log.info("Wrote %s", item["amazon_url"])
            except AliExpressAPIError as exc:
                if _token_rejected(str(exc)) or "rejected the access token" in str(exc):
                    raise SystemExit(str(exc)) from exc
                failed.append(item["amazon_url"])
                log.error("Failed %s: %s", item["amazon_url"], exc)
            except Exception as exc:
                failed.append(item["amazon_url"])
                log.error("Failed %s: %s", item["amazon_url"], exc)
        if failed:
            raise SystemExit(f"{len(failed)} link(s) failed. Fix them and run the scraper again.")
        log.info("Create.xlsx is complete. Check the file, then run: python create_listing/create_listing.py create_listing/Create.xlsx")
        return
    driver: Optional[WebDriver] = None
    failed: list[str] = []
    stop_message = ""
    try:
        driver = start_browser("AliExpress AU", ALI_PROFILE)
        for index, item in enumerate(pending):
            if index:
                time.sleep(5)
            try:
                append_create_row(CREATE_PATH, scrape_aliexpress(driver, item, item["amazon_url"]))
                log.info("Wrote %s", item["amazon_url"])
            except AliExpressCheck as exc:
                stop_message = str(exc)
                log.error("%s", exc)
                break
            except Exception as exc:
                if aliexpress_traffic_blocked(driver):
                    stop_message = "AliExpress blocked this network. Close Chrome and run the scraper again later."
                    log.error("%s", stop_message)
                    break
                if aliexpress_blocked(driver):
                    stop_message = "AliExpress is still showing the robot check. Tick I'm not a robot, then run the scraper again."
                    log.error("%s", stop_message)
                    break
                failed.append(item["amazon_url"])
                log.error("Failed %s: %s", item["amazon_url"], exc)
    finally:
        if driver is not None:
            driver.quit()
    if stop_message:
        raise SystemExit(stop_message)
    if failed:
        raise SystemExit(f"{len(failed)} link(s) failed. Fix them and run the scraper again.")
    log.info("Create.xlsx is complete. Check the file, then run: python create_listing/create_listing.py create_listing/Create.xlsx")


if __name__ == "__main__":
    main()
