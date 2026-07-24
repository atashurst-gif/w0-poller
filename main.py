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
import sys
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
    handlers=[logging.StreamHandler(sys.stdout)],
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
MAX_ATTEMPTS  = int(os.getenv("W0_MAX_ATTEMPTS", "5"))  # retry a failed send up to N times before giving up
HC_PING_URL   = os.getenv("HC_PING_URL", "https://hc-ping.com/1c584e6e-eb7c-464a-a546-71ae8a633ab8")

def default_seen_file() -> str:
    volume_path = os.getenv("RAILWAY_VOLUME_MOUNT_PATH", "")
    if volume_path:
        return str(Path(volume_path) / "w0_seen.json")
    if Path("/data").exists():
        return "/data/w0_seen.json"
    return "w0_seen.json"

SEEN_FILE = os.getenv("W0_SEEN_FILE") or default_seen_file()

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
BST_TEMPLATE_DECLAN = "bst_w0"   # Declan tenant uses bst_w0, not bst_nc0

# UKDT sheet — 11lc2uiVgJrKE_tQE5BE-JdsMT9kdjCfXnyGOoxLw0CA
# Tabs: "UKDT CT", "UKDT CTWA 1%", "UKDT WEBSITE"
UKDT_SHEET_ID  = os.getenv("UKDT_SHEET_ID", "11lc2uiVgJrKE_tQE5BE-JdsMT9kdjCfXnyGOoxLw0CA")
UKDT_TEMPLATE  = "ukdt_w0"
AUTOMATION_SHEET_ID = os.getenv("AUTOMATION_SHEET_ID", "1bggrSflZQa3ZCng6TQXbf1vNrnQWDnOBZcjDoCcjdf4")
AUTOMATION_TAB = os.getenv("AUTOMATION_TAB", "Sheet1")
W0_TRACKING_SHEET_ID = os.getenv("W0_TRACKING_SHEET_ID", AUTOMATION_SHEET_ID)
W0_TRACKING_TAB = os.getenv("W0_TRACKING_TAB", "W0 Tracking")
BOOKING_PENDING_STATUS = "booking pending"
SEQUENCE_CAMPAIGN_MAP = {"ukdt": "UKDT CT", "bst": "BST"}
STOPPED_SEQUENCE_STATUSES = {
    "replied", "completed", "opted out", "converted", "do not contact",
    "dnc", "dnq", "callback", "interested", "agreed", "lead passed",
    "verified", "moc set", "moc approved", "cbna",
}

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

def booking_window_for(now=None, lead_source: str = "ukdt") -> str:
    # Routing by ENQUIRY day/time (only reached when is_out_of_hours() is True):
    #   Mon-Thu >=18:00  -> next day        (nextday event)
    #   Fri >=14:00, Sat -> coming Monday   (monday event)
    #   Sun              -> coming Mon+Tue  (suntue / sunday-bst event)
    # Order matters: Sunday first, then Fri/Sat, then weekday-evening. Friday evening
    # must fall into the Fri/Sat->Monday bucket, NOT next-day (next day = Sat, closed).
    now = now or datetime.datetime.now(UK_TZ)
    wd, hr = now.weekday(), now.hour
    brand = (lead_source or "").strip().lower()
    is_weekday_evening = (wd <= 3 and hr >= 18)          # Mon-Thu after 18:00
    is_fri_or_sat      = (wd == 4) or (wd == 5)          # Fri (gated >=14:00 upstream) or Sat
    # SIMPLIFIED: one calendar per brand. Availability schedule (Mon-Fri) handles
    # "next working day" logic, so we no longer pick different slugs by day.
    if brand == "bst":
        return "callbacks-monday-bst"
    return "callbacks-monday"

def set_lead_attributes(phone: str, lead_source: str, booking_window: str = "", booking_url: str = "") -> bool:
    formatted = format_phone(phone)
    url = f"{WATI_API_URL}/api/v1/updateContactAttributes/{formatted}"
    headers = {"Authorization": f"Bearer {WATI_TOKEN}", "Content-Type": "application/json"}
    params = [
        {"name": "lead_source",    "value": lead_source},
    ]
    if booking_window:
        params.append({"name": "booking_window", "value": booking_window})
    if booking_url:
        params.append({"name": "booking_url", "value": booking_url})
    payload = {"customParams": params}
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        if r.status_code in (200, 201):
            return True
        log.error(f"attr-set {r.status_code} for {formatted}: {r.text[:200]}")
        return False
    except Exception as e:
        log.error(f"attr-set failed for {formatted}: {e}")
        return False


def set_lead_attributes_retry(phone: str, lead_source: str, booking_window: str = "",
                              booking_url: str = "", attempts: int = 5, delay: float = 4.0) -> bool:
    """Set attributes with retry. The WATI contact is created by the template send, which
    may lag slightly, so the first attribute write can land before the contact fully exists.
    Retry a few times so lead_source reliably persists on new contacts."""
    import time as _t
    for i in range(attempts):
        if set_lead_attributes(phone, lead_source, booking_window, booking_url):
            # Verify it actually stuck (WATI can 200 on a non-existent contact)
            if _verify_lead_source(phone, lead_source):
                return True
        if i < attempts - 1:
            _t.sleep(delay)
    log.error(f"attr-set did NOT persist for {format_phone(phone)} after {attempts} attempts")
    return False


def _verify_lead_source(phone: str, expected: str) -> bool:
    formatted = format_phone(phone)
    url = f"{WATI_API_URL}/api/v1/getContacts?pageSize=1&pageNumber=1&name={formatted}"
    headers = {"Authorization": f"Bearer {WATI_TOKEN}"}
    try:
        r = requests.get(url, headers=headers, timeout=30)
        if r.status_code != 200:
            return False
        cl = r.json().get("contact_list", [])
        if not cl:
            return False
        for p in cl[0].get("customParams", []):
            if p.get("name") == "lead_source" and str(p.get("value", "")).strip().lower() == expected.strip().lower():
                return True
        return False
    except Exception:
        return False


def ensure_w0_tracking_sheet(service):
    meta = service.spreadsheets().get(spreadsheetId=W0_TRACKING_SHEET_ID).execute()
    tabs = [s['properties']['title'] for s in meta.get('sheets', [])]
    if W0_TRACKING_TAB not in tabs:
        log.info(f"Creating '{W0_TRACKING_TAB}' tab...")
        service.spreadsheets().batchUpdate(
            spreadsheetId=W0_TRACKING_SHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": W0_TRACKING_TAB}}}]}
        ).execute()
        service.spreadsheets().values().append(
            spreadsheetId=W0_TRACKING_SHEET_ID,
            range=f"'{W0_TRACKING_TAB}'!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [[
                "Sent At", "TL-REF", "First Name", "Phone", "Campaign", "Status"
            ]]}
        ).execute()
        log.info(f"'{W0_TRACKING_TAB}' created.")


def append_w0_tracking_row(service, phone: str, first_name: str, campaign: str, status: str, now=None) -> bool:
    now = now or datetime.datetime.now(UK_TZ)
    formatted = format_phone(phone)
    tl_ref = f"W0-{campaign.upper().replace(' ', '-')}-{now.strftime('%Y%m%d%H%M%S')}-{formatted[-4:]}"
    try:
        service.spreadsheets().values().append(
            spreadsheetId=W0_TRACKING_SHEET_ID,
            range=f"'{W0_TRACKING_TAB}'!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [[
                now.strftime("%d/%m/%Y %H:%M"),
                tl_ref,
                first_name or "there",
                formatted,
                campaign,
                status,
            ]]}
        ).execute()
        log.info(f"{formatted}: added to {W0_TRACKING_TAB} as {status} ({campaign})")
        return True
    except Exception as e:
        log.error(f"{formatted}: failed to add W0 tracking row: {e}")
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
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)

# ─────────────────────────────────────────────
# Seen-row tracking (persisted to disk)
# ─────────────────────────────────────────────

def load_seen() -> dict:
    if Path(SEEN_FILE).exists():
        with open(SEEN_FILE) as f:
            raw = json.load(f)
        # Migrate legacy format {key: int} -> {key: {"count": int, "failed": []}}
        migrated = {}
        for k, v in raw.items():
            if isinstance(v, int):
                migrated[k] = {"count": v, "failed": []}
            elif isinstance(v, dict):
                migrated[k] = {"count": v.get("count", 0), "failed": v.get("failed", [])}
            else:
                migrated[k] = {"count": 0, "failed": []}
        return migrated
    return {}

def save_seen(seen: dict):
    # Format: {"<sheet_id>::<tab>": {"count": int, "failed": [{"idx": int, "attempts": int, "phone": str}]}}
    path = Path(SEEN_FILE)
    if path.parent and str(path.parent) != ".":
        path.parent.mkdir(parents=True, exist_ok=True)
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


def is_stopped_sequence_status(status: str) -> bool:
    status = (status or "").lower()
    return any(s in status for s in STOPPED_SEQUENCE_STATUSES)


def append_booking_pending_lead(service, phone: str, first_name: str, lead_source: str, now=None) -> bool:
    campaign = SEQUENCE_CAMPAIGN_MAP.get((lead_source or "").lower())
    if not campaign:
        log.warning(f"No sequence campaign mapped for lead_source={lead_source!r}")
        return False
    return append_w0_tracking_row(service, phone, first_name, campaign, BOOKING_PENDING_STATUS, now=now)

# ─────────────────────────────────────────────
# WATI send
# ─────────────────────────────────────────────

def send_w0(phone: str, first_name: str, template: str, api_url: str = None, token: str = None) -> str:
    """Returns a status string: 'ok' | 'dead' | 'retry'.
       'ok'    = sent (200/201, number valid)
       'dead'  = permanent fail, do NOT retry (invalid number / bad request that won't fix itself)
       'retry' = transient fail, safe to retry next cycle (timeout, 429, 5xx, auth)"""
    formatted = format_phone(phone)
    if not is_valid_phone(formatted):
        log.warning(f"Skipping invalid phone: {phone!r}")
        return "dead"

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
            # WATI's isValidWhatsAppNumber flag is unreliable (returns false even when the
            # message delivers), so we trust the HTTP status, not the body. A 200 = accepted.
            log.info(f"✓ W0 sent [{template}] → {formatted} ({first_name})")
            return "ok"
        else:
            log.error(f"WATI {r.status_code} for {formatted}: {r.text[:200]}")
            if r.status_code in (401, 403):
                ping("/fail")  # auth dead — turn healthcheck red so we get alerted
                return "retry"  # token may recover / be refreshed — don't lose the lead
            if r.status_code == 400:
                return "retry"  # your incident: transient 400 on approved template — retry
            if r.status_code == 429 or r.status_code >= 500:
                return "retry"  # throttle / server error — retry
            return "dead"       # other 4xx (404 etc.) won't fix on retry
    except Exception as e:
        log.error(f"WATI request failed for {formatted}: {e}")
        return "retry"  # network blip / timeout — retry

# ─────────────────────────────────────────────
# Poll one tab
# ─────────────────────────────────────────────

def _extract_first_name(raw_name: str, full_name: bool) -> str:
    raw_name = str(raw_name).strip()
    if not raw_name or raw_name.lower() in ("not found", "nan", ""):
        return "there"
    if "@" in raw_name:
        return "there"  # email landed in name field — bad form data
    if " " in raw_name:
        return raw_name.split()[0].title()
    if full_name:
        m = re.match(r"[A-Z][a-z]+", raw_name)  # camelCase FirstSurname
        return (m.group(0) if m else raw_name).title()
    return raw_name.title()

def _send_for_row(row: list, tab_cfg: dict, service=None) -> str:
    """Resolve routing + out-of-hours for a single row and send.
       Returns send_w0 status ('ok'|'dead'|'retry') or 'skip' if the row is unsendable."""
    tab       = tab_cfg["tab"]
    template  = tab_cfg["template"]
    phone_col = tab_cfg["phone_col"]
    name_col  = tab_cfg["name_col"]
    full_name = tab_cfg.get("full_name", False)

    if len(row) <= phone_col:
        return "skip"
    raw_phone = str(row[phone_col]).strip()
    if not raw_phone or raw_phone.lower() in ("", "not found", "nan"):
        return "skip"

    first_name = _extract_first_name(row[name_col] if len(row) > name_col else "", full_name)

    # ROUTING to Declan's (MDH) WATI — two cases:
    #  (a) whole-tab Declan sources via tab_cfg "wati"=="declan"
    #  (b) BST Form Meta rows with Creative=="Bailiff Companies"
    creative = (str(row[2]).strip().lower() if len(row) > 2 and row[2] else "")
    form_val = (str(row[1]).strip().lower() if len(row) > 1 and row[1] else "")
    # Declan's BST leads are identified by FORM = "BAILIFF PHOENIX" (his campaign form),
    # NOT by creative. Regen's BST leads use form "BAILIFF FORM NEW" (LETTER ARKLE creatives)
    # and must stay on Regen's tenant. Previously routed on creative=="bailiff companies"
    # which leaked 9 of Declan's BAILIFF PHOENIX leads to Regen.
    route_declan = (tab_cfg.get("wati") == "declan") or \
                   (tab == "BST Form Meta" and form_val == "bailiff phoenix")
    if route_declan:
        if not WATI_API_URL_DECLAN or not WATI_TOKEN_DECLAN:
            log.error(f"Declan-routed lead {raw_phone} ({tab}) but Declan WATI env not set — SKIPPING (not sending via Regen)")
            return "skip"
        declan_template = BST_TEMPLATE_DECLAN if tab == "BST Form Meta" else template
        return send_w0(raw_phone, first_name, declan_template,
                       api_url=WATI_API_URL_DECLAN, token=WATI_TOKEN_DECLAN)
    if False and is_out_of_hours() and template in W0W_MAP:
        w0w_template   = W0W_MAP[template]
        lead_source    = LEAD_SOURCE_MAP[template]
        booking_window = booking_window_for(lead_source=lead_source)
        booking_url    = f"https://cal.com/debthelpbooking/{booking_window}"
        # Send FIRST so the WATI contact exists, THEN stamp attributes on it.
        # (Stamping before send hits a non-existent contact and silently no-ops,
        #  leaving real new leads with no lead_source -> misroute in the chatbot.)
        status = send_w0(raw_phone, first_name, w0w_template)
        if status == "ok":
            set_lead_attributes_retry(raw_phone, lead_source, booking_window, booking_url)
            if service:
                append_booking_pending_lead(service, raw_phone, first_name, lead_source)
        return status
    # In-hours (or non-W0W template): stamp brand identity on the WATI contact so
    # downstream flows/conditions always have lead_source. booking_window/booking_url
    # stay out-of-hours-only. Lowercase to match the WEEKEND FORM chatbot condition.
    brand_source = LEAD_SOURCE_MAP.get(template)
    status = send_w0(raw_phone, first_name, template)
    if status == "ok" and brand_source:
        set_lead_attributes_retry(raw_phone, brand_source)
    if status == "ok" and service:
        lead_source = LEAD_SOURCE_MAP.get(template, template.replace("_w0", "")).upper()
        append_w0_tracking_row(service, raw_phone, first_name, lead_source, "w0 sent")
    return status

def poll_tab(service, tab_cfg: dict, seen: dict) -> int:
    sheet_id  = tab_cfg["sheet_id"]
    tab       = tab_cfg["tab"]
    skip      = tab_cfg.get("skip_rows", 1)

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
        seen[key] = {"count": total, "failed": []}
        log.info(f"{tab}: first run, seeding at row {total}")
        return 0

    entry = seen[key]
    prev_count = entry["count"]
    fired = 0

    # ── Retry previously-failed rows first (re-read fresh by absolute index) ──
    still_failed = []
    phone_col = tab_cfg["phone_col"]
    for f in entry.get("failed", []):
        idx = f["idx"]
        want_phone = f.get("phone", "")
        # Verify the row at idx is still the same lead — sheet may have shifted (insert/delete above)
        if idx >= total or (want_phone and (len(data_rows[idx]) <= phone_col or str(data_rows[idx][phone_col]).strip() != want_phone)):
            # Re-scan for the phone across the sheet
            found = None
            if want_phone:
                for j, r in enumerate(data_rows):
                    if len(r) > phone_col and str(r[phone_col]).strip() == want_phone:
                        found = j
                        break
            if found is None:
                log.warning(f"{tab}: failed row (phone {want_phone or '?'}) no longer locatable — dropping")
                continue
            log.info(f"{tab}: failed row moved {idx}->{found} (sheet shifted) — retrying at new index")
            idx = found
        status = _send_for_row(data_rows[idx], tab_cfg, service)
        if status == "ok":
            fired += 1
            log.info(f"{tab}: retry OK for row idx {idx} (attempt {f['attempts']+1})")
        elif status == "skip" or status == "dead":
            log.warning(f"{tab}: dropping row idx {idx} — status={status} after {f['attempts']} attempt(s)")
        else:  # retry
            attempts = f["attempts"] + 1
            if attempts >= MAX_ATTEMPTS:
                log.error(f"{tab}: row idx {idx} hit MAX_ATTEMPTS ({MAX_ATTEMPTS}) — giving up, lead dropped")
            else:
                still_failed.append({"idx": idx, "attempts": attempts, "phone": f.get("phone", "")})
        time.sleep(0.5)

    # ── Process genuinely new rows ──
    if total > prev_count:
        for offset, row in enumerate(data_rows[prev_count:]):
            idx = prev_count + offset
            status = _send_for_row(row, tab_cfg, service)
            if status == "ok":
                fired += 1
            elif status == "retry":
                ph = str(row[tab_cfg["phone_col"]]).strip() if len(row) > tab_cfg["phone_col"] else ""
                still_failed.append({"idx": idx, "attempts": 1, "phone": ph})
                log.warning(f"{tab}: new row idx {idx} failed (transient) — queued for retry")
            # 'dead' / 'skip' — do nothing, not retried
            time.sleep(0.5)

    entry["count"]  = total
    entry["failed"] = still_failed
    if fired:
        log.info(f"{tab}: {fired} W0 message(s) sent")
    if still_failed:
        log.info(f"{tab}: {len(still_failed)} row(s) queued for retry next cycle")
    return fired

# ─────────────────────────────────────────────
# Main loop
# ─────────────────────────────────────────────

APPS2_SHEET_ID = os.getenv("APPS2_SHEET_ID", "16qnJ842lFAo-4FVhloipJMFdCs7Spc2hFgoWnVGwVUs")
APPS2_TAB      = os.getenv("APPS2_TAB", "Apps 2.0")
CALLBACK_SET_STATUS = "Callback Set"
CALLBACK_SYNC_DRY_RUN = os.getenv("CALLBACK_SYNC_DRY_RUN", "1") == "1"


def sync_callback_set(service) -> int:
    """Find leads who have booked (Apps 2.0 has an Appointment) and mark their
    Sheet1 row Status = 'Callback Set' so the NC sequence stops messaging them."""
    # 1. collect booked phone numbers from Apps 2.0
    try:
        apps = service.spreadsheets().values().get(
            spreadsheetId=APPS2_SHEET_ID, range=f"'{APPS2_TAB}'!A:L"
        ).execute().get("values", [])
    except Exception as e:
        log.error(f"callback-sync: cannot read Apps 2.0: {e}")
        return 0
    if not apps:
        return 0
    hdr = [h.strip().lower() for h in apps[0]]
    def ci(name):
        for i, h in enumerate(hdr):
            if name in h:
                return i
        return None
    a_num, a_appt = ci("number"), ci("appointment")
    if a_num is None or a_appt is None:
        log.error("callback-sync: Apps 2.0 missing Number/Appointment column")
        return 0
    booked = set()
    for r in apps[1:]:
        num = r[a_num].strip() if a_num < len(r) else ""
        appt = r[a_appt].strip() if a_appt < len(r) else ""
        if num and appt:
            ph = format_phone(num)
            if is_valid_phone(ph):
                booked.add(ph)
    if not booked:
        return 0

    # 2. scan Sheet1 and mark matching rows
    try:
        rows = service.spreadsheets().values().get(
            spreadsheetId=AUTOMATION_SHEET_ID, range="'Sheet1'!A:F"
        ).execute().get("values", [])
    except Exception as e:
        log.error(f"callback-sync: cannot read Sheet1: {e}")
        return 0

    updates = []
    for i, r in enumerate(rows[1:], start=2):
        phone_raw = r[3].strip() if len(r) > 3 else ""
        status    = r[5].strip() if len(r) > 5 else ""
        if not phone_raw:
            continue
        if is_stopped_sequence_status(status):
            continue
        if format_phone(phone_raw) in booked:
            updates.append((i, r[2].strip() if len(r) > 2 else ""))

    if not updates:
        return 0

    if CALLBACK_SYNC_DRY_RUN:
        for rownum, name in updates:
            log.info(f"callback-sync [DRY RUN] would set row {rownum} ({name}) -> {CALLBACK_SET_STATUS}")
        return 0

    done = 0
    for rownum, name in updates:
        try:
            service.spreadsheets().values().update(
                spreadsheetId=AUTOMATION_SHEET_ID,
                range=f"'Sheet1'!F{rownum}",
                valueInputOption="RAW",
                body={"values": [[CALLBACK_SET_STATUS]]},
            ).execute()
            log.info(f"callback-sync: row {rownum} ({name}) -> {CALLBACK_SET_STATUS}")
            done += 1
        except Exception as e:
            log.error(f"callback-sync: failed row {rownum}: {e}")
    return done


CBNA_DRY_RUN = os.getenv("CBNA_DRY_RUN", "1") == "1"

CBNA_SEQUENCE = [
    (1, "ukdt_cbna_1",  "bst_cbana_1", "name",       0),
    (2, "ukdt_cb_na_2", "bst_cbana_2", "name",       5),
    (3, "ukdt_cbana_3", "bst_cbna_3",  "name",       24 * 60),
    (4, "ukdt_cbna4",   "bst_cbna4",   None,         48 * 60),
    (5, "ukdt_cbna5",   "bst_cbna5",   "first_name", 72 * 60),
]

CBNA_STEP_COLS = {1: 12, 2: 14, 3: 16, 4: 18, 5: 20}
CBNA_RESP_COLS = {1: 13, 2: 15, 3: 17, 4: 19}
CBNA_ANSWERED_COL = 9
CBNA_NAME_COL     = 2
CBNA_PHONE_COL    = 3
CBNA_CAMPAIGN_COL = 4
CBNA_STAMP_FMT = "%d/%m/%Y %H:%M"


def _col_letter(idx):
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cbna_parse_stamp(cell):
    cell = (cell or "").strip()
    if not cell:
        return 0, None
    m = re.match(r"\s*(\d+)\s*-\s*(.+)", cell)
    if not m:
        return 0, None
    step = int(m.group(1))
    try:
        dt = datetime.datetime.strptime(m.group(2).strip(), CBNA_STAMP_FMT).replace(tzinfo=UK_TZ)
    except Exception:
        dt = None
    return step, dt


def _cbna_has_replied(phone, since_dt):
    if since_dt is None:
        return False
    try:
        url = WATI_API_URL + "/api/v1/getMessages/" + format_phone(phone)
        r = requests.get(url, headers={"Authorization": "Bearer " + WATI_TOKEN,
                                       "accept": "application/json"}, timeout=20)
        items = r.json().get("messages", {}).get("items", [])
    except Exception as e:
        log.warning("cbna: reply-check failed for %s: %s" % (phone, e))
        return False
    for i in items:
        if i.get("owner") is False:
            created = i.get("created") or ""
            try:
                dt = datetime.datetime.strptime(created[:19], "%Y-%m-%dT%H:%M:%S")
                dt = dt.replace(tzinfo=datetime.timezone.utc).astimezone(UK_TZ)
            except Exception:
                continue
            if dt > since_dt:
                return True
    return False


def _cbna_send(phone, template, param, first_name):
    formatted = format_phone(phone)
    if CBNA_DRY_RUN:
        log.info("cbna [DRY RUN] would send %s to %s (%s)" % (template, formatted, first_name))
        return True
    url = WATI_API_URL + "/api/v2/sendTemplateMessages"
    headers = {"Authorization": "Bearer " + WATI_TOKEN, "Content-Type": "application/json"}
    if param:
        receiver = {"whatsappNumber": formatted,
                    "customParams": [{"name": param, "value": first_name}]}
    else:
        receiver = {"whatsappNumber": formatted}
    payload = {"template_name": template,
               "broadcast_name": "cbna_" + template + "_" + formatted[-4:],
               "receivers": [receiver]}
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        if r.status_code in (200, 201):
            body = {}
            try:
                body = r.json()
            except Exception:
                pass
            if body.get("result") is True:
                log.info("cbna: sent %s to %s (%s)" % (template, formatted, first_name))
                return True
            log.error("cbna: REJECTED %s for %s: %s" % (template, formatted, str(body)[:200]))
            return False
        log.error("cbna: WATI %s for %s: %s" % (r.status_code, formatted, r.text[:200]))
        return False
    except Exception as e:
        log.error("cbna: request failed for %s: %s" % (formatted, e))
        return False


def sync_cbna(service):
    now = datetime.datetime.now(UK_TZ)
    rng = "'" + APPS2_TAB + "'!A:U"
    try:
        rows = service.spreadsheets().values().get(
            spreadsheetId=APPS2_SHEET_ID, range=rng).execute().get("values", [])
    except Exception as e:
        log.error("cbna: cannot read Apps 2.0: %s" % e)
        return 0
    if not rows:
        return 0
    sent_total = 0
    for i, r in enumerate(rows[1:], start=2):
        def g(idx):
            return r[idx].strip() if idx < len(r) else ""
        if g(CBNA_ANSWERED_COL).lower() != "no":
            continue
        _cut = os.getenv("CBNA_CUTOFF_DATE", "")
        if _cut:
            try:
                _cd = datetime.datetime.strptime(_cut, "%d/%m/%Y").replace(tzinfo=UK_TZ)
            except Exception:
                _cd = None
            if _cd:
                _rd = None
                for _f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
                    try:
                        _rd = datetime.datetime.strptime(g(0), _f).replace(tzinfo=UK_TZ)
                        break
                    except Exception:
                        pass
                if _rd is None or _rd < _cd:
                    continue
        phone = g(CBNA_PHONE_COL)
        if not phone or not is_valid_phone(format_phone(phone)):
            continue
        if any(g(c) for c in CBNA_RESP_COLS.values()):
            continue
        cur_step, last_dt = 0, None
        for step in (5, 4, 3, 2, 1):
            st, dt = _cbna_parse_stamp(g(CBNA_STEP_COLS[step]))
            if st:
                cur_step, last_dt = st, dt
                break
        if cur_step >= 5:
            continue
        if cur_step and _cbna_has_replied(phone, last_dt):
            resp_col = CBNA_RESP_COLS.get(cur_step)
            if resp_col is not None:
                stamp = now.strftime(CBNA_STAMP_FMT)
                if CBNA_DRY_RUN:
                    log.info("cbna [DRY RUN] row %s: reply detected, would stamp Response%s" % (i, cur_step))
                else:
                    service.spreadsheets().values().update(
                        spreadsheetId=APPS2_SHEET_ID,
                        range="'" + APPS2_TAB + "'!" + _col_letter(resp_col) + str(i),
                        valueInputOption="RAW",
                        body={"values": [[stamp]]}).execute()
                    log.info("cbna: row %s reply -> Response%s %s" % (i, cur_step, stamp))
            continue
        nxt = cur_step + 1
        step_no, ukdt_t, bst_t, param, delay_min = CBNA_SEQUENCE[nxt - 1]
        if nxt == 1:
            due = True
        else:
            if last_dt is None:
                continue
            due = now >= (last_dt + datetime.timedelta(minutes=delay_min))
        if not due:
            continue
        campaign = g(CBNA_CAMPAIGN_COL).upper()
        template = bst_t if "BST" in campaign else ukdt_t
        full = g(CBNA_NAME_COL)
        first_name = full.split()[0].title() if full else "there"
        if _cbna_send(phone, template, param, first_name):
            sent_total += 1
            cell = str(step_no) + " - " + now.strftime(CBNA_STAMP_FMT)
            if CBNA_DRY_RUN:
                log.info("cbna [DRY RUN] row %s: would write '%s' to CBNA%s" % (i, cell, step_no))
            else:
                service.spreadsheets().values().update(
                    spreadsheetId=APPS2_SHEET_ID,
                    range="'" + APPS2_TAB + "'!" + _col_letter(CBNA_STEP_COLS[step_no]) + str(i),
                    valueInputOption="RAW",
                    body={"values": [[cell]]}).execute()
            time.sleep(1)
    if sent_total:
        log.info("cbna: cycle complete - %s message(s) sent%s" % (sent_total, " [DRY RUN]" if CBNA_DRY_RUN else ""))
    return sent_total


def main():
    log.info("W0 Poller starting...")
    log.info(f"BST template: {BST_TEMPLATE} | UKDT template: {UKDT_TEMPLATE}")
    log.info(f"Watching {len(WATCH_TABS)} tabs | Poll interval: {POLL_INTERVAL}s")

    seen = load_seen()

    while True:
        try:
            service = get_sheets_service()
            ensure_w0_tracking_sheet(service)
            total_fired = 0
            for tab_cfg in WATCH_TABS:
                fired = poll_tab(service, tab_cfg, seen)
                total_fired += fired
            save_seen(seen)
            sync_callback_set(service)
            sync_cbna(service)
            ping()  # healthy cycle — Sheets read OK, no auth failure
            if total_fired:
                log.info(f"Cycle complete — {total_fired} W0 message(s) sent total")
        except Exception as e:
            log.exception(f"Poll cycle error: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
