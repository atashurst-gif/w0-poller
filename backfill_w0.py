"""
backfill_w0.py — ONE-OFF catch-up for missed W0 sends.
Reuses the live poller's format_phone() and send_w0().
Defaults to DRY RUN. LIVE=1 to actually send.
"""
import os, re, time, logging
import main as poller

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("backfill")
LIVE = os.getenv("LIVE", "") == "1"

BACKFILL = [
    {"sheet_id": poller.BST_SHEET_ID, "tab": "BST Form Meta", "template": poller.BST_TEMPLATE,
     "phone_col": 6, "name_col": 3, "skip_rows": 1, "full_name": False, "from_row": 880, "to_row": 892},
    {"sheet_id": poller.UKDT_SHEET_ID, "tab": "UKDT CT", "template": poller.UKDT_TEMPLATE,
     "phone_col": 5, "name_col": 3, "skip_rows": 1, "full_name": True, "from_row": 1475, "to_row": 1511},
]

def first_name_of(raw_name, full_name):
    raw_name = str(raw_name).strip()
    if not raw_name or raw_name.lower() in ("not found", "nan", ""):
        return "there"
    if "@" in raw_name:
        return "there"
    if " " in raw_name:
        raw_name = raw_name.split()[0]
    elif full_name:
        # camelCase fullname e.g. LeilaWilliams -> Leila (split before 2nd capital)
        m = re.match(r"[A-Z][a-z]+", raw_name)
        if m:
            raw_name = m.group(0)
    return raw_name.title()

def main():
    log.info("W0 backfill | " + ("LIVE — WILL SEND" if LIVE else "DRY RUN — sends nothing"))
    service = poller.get_sheets_service()
    cand = sent = skip = 0
    for cfg in BACKFILL:
        tab = cfg["tab"]
        rows = service.spreadsheets().values().get(
            spreadsheetId=cfg["sheet_id"], range=f"'{tab}'!A:Z").execute().get("values", [])
        data = rows[cfg["skip_rows"]:]
        window = data[cfg["from_row"]:cfg["to_row"]]
        log.info(f"\n=== {tab} | rows {cfg['from_row']}-{cfg['to_row']} | {len(window)} leads ===")
        for i, row in enumerate(window):
            rn = cfg["skip_rows"] + cfg["from_row"] + i + 1
            if len(row) <= cfg["phone_col"]:
                log.warning(f"  row {rn}: no phone col — SKIP"); skip += 1; continue
            raw_phone = str(row[cfg["phone_col"]]).strip()
            if not raw_phone or raw_phone.lower() in ("", "not found", "nan"):
                log.warning(f"  row {rn}: empty phone — SKIP"); skip += 1; continue
            raw_name = row[cfg["name_col"]] if len(row) > cfg["name_col"] else ""
            fn = first_name_of(raw_name, cfg["full_name"])
            fmt = poller.format_phone(raw_phone)
            if not poller.is_valid_phone(fmt):
                log.warning(f"  row {rn}: invalid phone {raw_phone!r} — SKIP"); skip += 1; continue
            cand += 1
            log.info(f"  row {rn}: {fn:<15} {fmt:<15} [{cfg['template']}]")
            if LIVE:
                if poller.send_w0(raw_phone, fn, cfg["template"]): sent += 1
                time.sleep(0.5)
    log.info("\n" + "="*50)
    log.info(f"Candidates: {cand} | Skipped: {skip}")
    log.info(f"SENT: {sent}" if LIVE else "DRY RUN — nothing sent. Re-run with LIVE=1 to send.")

if __name__ == "__main__":
    main()
