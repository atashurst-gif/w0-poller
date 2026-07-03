"""
W0 Poller — Instant WhatsApp on new enquiries
Watches BST Form sheet and UKDT CT sheet every 60s.
Fires bst_nc0 or ukdt_w0 the moment a new row appears.
Runs standalone alongside wati_sequence and gmail_monitor.
"""

import os
import re
import json
import time
import base64
import logging
import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────

WATI_API_URL  = os.getenv("WATI_API_URL", "https://eu-api.wati.io/602557")
WATI_TOKEN    = os.getenv("WATI_TOKEN", "")
WATI_API_URL_DECLAN = os.getenv("WATI_API_URL_DECLAN", "")
WATI_TOKEN_DECLAN   = os.getenv("WATI_TOKEN_DECLAN", "")
POLL_INTERVAL = int(os.getenv("W0_POLL_INTERVAL", "60"))
SEEN_FILE     = os.getenv("W0_SEEN_FILE", "w0_seen.json")
HC_PING_URL   = os.getenv("HC_PING_URL", "https://hc-ping.com/1c584e6e-eb7c-464a-a546-71ae8a633ab8")

def ping(suffix=""):
    """Ping healthcheck. suffix='/fail' marks the check failed (turns red)."""
    if not HC_PING_URL:
        return
    try:
        requests.get(HC_PING_URL + suffix, timeout=10)
    except Exception:
        pass  # never let a ping failure break the poll loop

# BST sheet — 1Sp0Zo7j9a-73R4kYV-2MSQzXUb8BmlDD0hcCcTUcC1E
# Tabs: "BST Form Meta", "BST Website"
BST_SHEET_ID   = os.getenv("BST_SHEET_ID", "1Sp0Zo7j9a-73R4kYV-2MSQzXUb8BmlDD0hcCcTUcC1E")
BST_TEMPLATE   = "bst_nc0"

# UKDT sheet — 11lc2uiVgJrKE_tQE5BE-JdsMT9kdjCfXnyGOoxLw0CA
# Tabs: "UKDT CT", "UKDT CTWA 1%", "UKDT WEBSITE"
UKDT_SHEET_ID  = os.getenv("UKDT_SHEET_ID", "11lc2uiVgJrKE_tQE5BE-JdsMT9kdjCfXnyGOoxLw0CA")
UKDT_TEMPLATE  = "ukdt_w0"

# ── Out-of-hours booking gate ──────────────────
UK_TZ = ZoneInfo("Europe/London")
W0W_MAP         = {"ukdt_w0": "ukdt_w0w", "bst_nc0": "bst_w0w"}
LEAD_SOURCE_MAP = {"ukdt_w0": "ukdt",     "bst_nc0": "bst"}

def is_out_of_hours(now=None) -> bool:
    now = now or datetime.datetime.now(UK_TZ)
    wd, hr = now.weekday(), now.hour
    if wd <= 3:            return hr >= 18
    if wd == 4:            return hr >= 14
    return True

def booking_window_for(now=None) -> str:
    now = now or datetime.datetime.now(UK_TZ)
    wd = now.weekday()
    if wd in (4, 5): return "callbacks-monday"    # Fri/Sat -> Monday
    if wd == 6:      return "callbacks-suntue"     # Sun -> Mon+Tue
    return "callbacks-monday"                       # Mon-Thu eve -> Monday (until rolling event)

def set_lead_attributes(phone: str, lead_source: str, booking_window: str) -> bool:
    formatted = format_phone(phone)
    url = f"{WATI_API_URL}/api/v1/updateContactAttributes/{formatted}"
    headers = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json"}
    payload = {"customParams": [
        {"name": "lead_source",    "value": lead_source},
        {"name": "booking_window", "value": booking_window},
    ]}
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        if r.status_code in (200, 201):
            return True
        log.error(f"attr-set {r.status_code} for {formatted}: {r.text[:200]}")
        return False
    except Exception as e:
        log.error(f"attr-set failed for {formatted}: {e}")
        return False

# Tabs to watch per sheet — (sheet_id, tab_name, phone_col_index, name_col_index, skip_rows)
# phone_col_index / name_col_index = 0-based column index
# skip_rows = number of header rows to skip

WATCH_TABS = [
    # BST Form Meta: Created(0) Form(1) Creative(2) FirstName(3) Surname(4) Email(5) Number(6)
    {
        "sheet_id":    BST_SHEET_ID,
        "tab":         "BST Form Meta",
        "template":    BST_TEMPLATE,
        "phone_col":   6,
        "name_col":    3,
        "skip_rows":   1,
    },
    # BST Website: Form(0) FirstName(1) Surname(2) Email(3) Number(4) Source(5)
    {
        "sheet_id":    BST_SHEET_ID,
        "tab":         "BST Website",
        "template":    BST_TEMPLATE,
        "phone_col":   4,
        "name_col":    1,
        "skip_rows":   1,
    },
    # UKDT CT: Created(0) Form(1) AdName(2) FullName(3) Email(4) PhoneNumber(5)
    {
        "sheet_id":    UKDT_SHEET_ID,
        "tab":         "UKDT CT",
        "template":    UKDT_TEMPLATE,
        "phone_col":   5,
        "name_col":    3,
        "skip_rows":   1,
        "full_name":   True,   # name col contains full name — extract first word
    },
    # UKDT CTWA 1%: Date(0) Form(1) FullName(2) PhoneNumber(3)
    {
        "sheet_id":    UKDT_SHEET_ID,
        "tab":         "UKDT CTWA 1%",
        "template":    UKDT_TEMPLATE,
        "phone_col":   3,
        "name_col":    2,
        "skip_rows":   1,
        "full_name":   True,
    },
    # UKDT WEBSITE: Date(0) Form(1) FirstName(2) Surname(3) PhoneNumber(4) Email(5) Source(6)
    {
        "sheet_id":    UKDT_SHEET_ID,
        "tab":         "UKDT WEBSITE",
        "template":    UKDT_TEMPLATE,
        "phone_col":   4,
        "name_col":    2,
        "skip_rows":   1,
    },
    # UKDTCTD1 (Dec) — Declan's UKDT source: Created(0) Form(1) AdName(2) FullName(3) Email(4) Phone(5)
    # Whole tab routes to DECLAN's WATI with ukdt_ct_w0 (NOT Regen's).
    {
        "sheet_id":    UKDT_SHEET_ID,
        "tab":         "UKDTCTD1 (Dec)",
        "template":    "ukdt_ct_w0",
        "phone_col":   5,
        "name_col":    3,
        "skip_rows":   1,
        "full_name":   True,
        "wati":        "declan",
    },
]

# ─────────────────────────────────────────────
# Google Sheets auth
# ─────────────────────────────────────────────

def get_sheets_service():
    sa_b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_B64", "")
    if sa_b64:
        padded = sa_b64 + "=" * (-len(sa_b64) % 4)
        sa_dict = json.loads(base64.b64decode(padded).decode())
    else:
        with open("service_account.json") as f:
            sa_dict = json.load(f)
    creds = Credentials.from_service_account_info(
        sa_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    return build("sheets", "v4", credentials=creds)

# ─────────────────────────────────────────────
# Seen-row tracking (persisted to disk)
# ─────────────────────────────────────────────

def load_seen() -> dict:
    if Path(SEEN_FILE).exists():
        with open(SEEN_FILE) as f:
            return json.load(f)
    return {}

def save_seen(seen: dict):
    with open(SEEN_FILE, "w") as f:
        json.dump(seen, f, indent=2)

def seen_key(sheet_id: str, tab: str) -> str:
    return f"{sheet_id}::{tab}"

# ─────────────────────────────────────────────
# Phone normalisation
# ─────────────────────────────────────────────

def format_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", str(raw))
    # Handle scientific notation e.g. 4.4738E+11
    try:
        if "e" in digits.lower() or "." in str(raw).lower():
            digits = str(int(float(str(raw))))
            digits = re.sub(r"\D", "", digits)
    except Exception:
        pass
    if digits.startswith("07") and len(digits) == 11:
        return "44" + digits[1:]
    if digits.startswith("447") and len(digits) == 12:
        return digits
    if digits.startswith("44") and len(digits) >= 11:
        return digits
    if digits.startswith("7") and len(digits) == 10:
        return "44" + digits
    return digits

def is_valid_phone(phone: str) -> bool:
    digits = re.sub(r"\D", "", phone)
    return len(digits) >= 10

# ─────────────────────────────────────────────
# WATI send
# ─────────────────────────────────────────────

def send_w0(phone: str, first_name: str, template: str, api_url: str = None, token: str = None) -> bool:
    formatted = format_phone(phone)
    if not is_valid_phone(formatted):
        log.warning(f"Skipping invalid phone: {phone!r}")
        return False

    api_url = api_url or WATI_API_URL    # default: Regen's WATI
    token   = token   or WATI_TOKEN      # default: Regen's WATI
    url = f"{api_url}/api/v2/sendTemplateMessages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {
        "template_name":  template,
        "broadcast_name": f"w0_{template}_{formatted[-4:]}",
        "receivers": [
            {
                "whatsappNumber": formatted,
                "customParams": [{"name": "first_name", "value": first_name}],
            }
        ],
    }
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        if r.status_code in (200, 201):
            log.info(f"✓ W0 sent [{template}] → {formatted} ({first_name})")
            return True
        else:
            log.error(f"WATI {r.status_code} for {formatted}: {r.text[:200]}")
            if r.status_code in (401, 403):
                ping("/fail")  # auth dead — turn healthcheck red so we get alerted
            return False
    except Exception as e:
        log.error(f"WATI request failed for {formatted}: {e}")
        return False

# ─────────────────────────────────────────────
# Poll one tab
# ─────────────────────────────────────────────

def poll_tab(service, tab_cfg: dict, seen: dict) -> int:
    sheet_id  = tab_cfg["sheet_id"]
    tab       = tab_cfg["tab"]
    template  = tab_cfg["template"]
    phone_col = tab_cfg["phone_col"]
    name_col  = tab_cfg["name_col"]
    skip      = tab_cfg.get("skip_rows", 1)
    full_name = tab_cfg.get("full_name", False)

    key = seen_key(sheet_id, tab)

    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{tab}'!A:Z",
        ).execute()
    except Exception as e:
        log.warning(f"Could not read {tab}: {e}")
        return 0

    rows = result.get("values", [])
    data_rows = rows[skip:]  # skip header
    total = len(data_rows)

    # First run — record current count, fire nothing (don't spam existing data)
    if key not in seen:
        seen[key] = total
        log.info(f"{tab}: first run, seeding at row {total}")
        return 0

    prev_count = seen[key]
    if total <= prev_count:
        return 0

    new_rows = data_rows[prev_count:]
    fired = 0

    for row in new_rows:
        # Get phone
        if len(row) <= phone_col:
            continue
        raw_phone = str(row[phone_col]).strip()
        if not raw_phone or raw_phone.lower() in ("", "not found", "nan"):
            continue

        # Get name — first word only
        raw_name = str(row[name_col]).strip() if len(row) > name_col else ""
        if not raw_name or raw_name.lower() in ("not found", "nan", ""):
            first_name = "there"
        elif "@" in raw_name:
            first_name = "there"  # email landed in name field — bad form data
        elif " " in raw_name:
            first_name = raw_name.split()[0].title()
        elif full_name:
            m = re.match(r"[A-Z][a-z]+", raw_name)  # camelCase FirstSurname
            first_name = (m.group(0) if m else raw_name).title()
        else:
            first_name = raw_name.title()

        # ROUTING to Declan's (MDH) WATI — two cases:
        #  (a) whole-tab Declan sources via tab_cfg "wati"=="declan"
        #  (b) BST Form Meta rows with Creative=="Bailiff Companies"
        # Everything else is unchanged and uses Regen's WATI as before.
        creative = (str(row[2]).strip().lower() if len(row) > 2 and row[2] else "")
        route_declan = (tab_cfg.get("wati") == "declan") or \
                       (tab == "BST Form Meta" and creative == "bailiff companies")
        if route_declan:
            if not WATI_API_URL_DECLAN or not WATI_TOKEN_DECLAN:
                log.error(f"Declan-routed lead {raw_phone} ({tab}) but Declan WATI env not set — SKIPPING (not sending via Regen)")
                continue
            sent = send_w0(raw_phone, first_name, template,
                           api_url=WATI_API_URL_DECLAN, token=WATI_TOKEN_DECLAN)
        else:
            if is_out_of_hours() and template in W0W_MAP:
                w0w_template   = W0W_MAP[template]
                lead_source    = LEAD_SOURCE_MAP[template]
                booking_window = booking_window_for()
                set_lead_attributes(raw_phone, lead_source, booking_window)
                sent = send_w0(raw_phone, first_name, w0w_template)
            else:
                sent = send_w0(raw_phone, first_name, template)
        if sent:
            fired += 1
        time.sleep(0.5)  # gentle pacing

    seen[key] = total
    if fired:
        log.info(f"{tab}: {fired} W0 message(s) sent ({total - prev_count} new rows)")
    return fired

# ─────────────────────────────────────────────
# Main loop
# ─────────────────────────────────────────────

def main():
    log.info("W0 Poller starting...")
    log.info(f"BST template: {BST_TEMPLATE} | UKDT template: {UKDT_TEMPLATE}")
    log.info(f"Watching {len(WATCH_TABS)} tabs | Poll interval: {POLL_INTERVAL}s")

    seen = load_seen()

    while True:
        try:
            service = get_sheets_service()
            total_fired = 0
            for tab_cfg in WATCH_TABS:
                fired = poll_tab(service, tab_cfg, seen)
                total_fired += fired
            save_seen(seen)
            ping()  # healthy cycle — Sheets read OK, no auth failure
            if total_fired:
                log.info(f"Cycle complete — {total_fired} W0 message(s) sent total")
        except Exception as e:
            log.exception(f"Poll cycle error: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
