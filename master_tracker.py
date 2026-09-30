#!/usr/bin/env python3
"""MASTER TRACKER + WEEKLY SNAPSHOT for the Arkle workbook.

Collects every case from Claire's and Bianca's Case Trackers and the workbook's
own 'Leads Passed' tab (agent = Lead Gen column), writes them as plain VALUES
(no formulas, no charts) into two tabs, and only rewrites when the data changed.

Standalone (from ~/Desktop/w0-poller):
  python3 master_tracker.py          # preview: counts + first rows, writes nothing
  GO=1 python3 master_tracker.py     # build/refresh the two tabs now
Inside the poller: master_tracker.tick(get_sheets_service) every MT_INTERVAL_S.
"""
import os, sys, re, time, hashlib, datetime, logging, collections

log = logging.getLogger(__name__)

ARKLE      = os.getenv("CBS_APPS2_SHEET_ID", "16qnJ842lFAo-4FVhloipJMFdCs7Spc2hFgoWnVGwVUs")
CLAIRE     = os.getenv("MT_CLAIRE_SHEET_ID", "1sFBjBifnoLlDwN2kvAai-dDGrtLRtHx_pZjI55djM7o")
BIANCA     = os.getenv("MT_BIANCA_SHEET_ID", "1spQ-JmLpr0GvrF87-JDSfK5KGpSLRDdWX-vVclLUDwk")
MASTER_TAB = os.getenv("MT_MASTER_TAB", "MASTER TRACKER")
SNAP_TAB   = os.getenv("MT_SNAP_TAB", "WEEKLY SNAPSHOT")
ENABLED    = os.getenv("MASTER_TRACKER_ENABLED", "1") == "1"
DRY_RUN    = os.getenv("MASTER_TRACKER_DRY_RUN", "1") == "1"
INTERVAL_S = int(os.getenv("MT_INTERVAL_S", "900"))

# (agent label or None=take from a column, sheet id, range, agent column header)
SOURCES = [
    ("Claire", CLAIRE, "'Case Tracker'!A:N", None),
    ("Bianca", BIANCA, "A:N", None),                      # first tab of Bianca's sheet
    (None,     ARKLE,  "'Leads Passed'!A:J", "lead gen"),  # Josh / DEX / ELI
]

HEAD = ["Date Passed", "Week", "Agent", "TL Ref", "Client Name", "Lead Source", "Contact Number", "Advisor",
        "Stage", "Outcome", "Agreed Date", "SIP Complete", "Prop Back", "MOC Set", "Approval Date", "Approved",
        "Partner", "Notes", "From"]                       # 19 columns: A..S
SNAP_HEAD = ["Agent", "Leads Passed", "DMP Passed", "Agreed", "SIP Complete", "MOC Set", "Approved", "DNQ",
             "Lead > Agreed", "Agreed > SIP", "Lead > Approved"]
NEWEST_FIRST = os.getenv("MT_NEWEST_FIRST", "1") == "1"

_last_tick = 0.0
_last_hash = None


# ── helpers ────────────────────────────────────────────────────────────────────
def _clean(s):
    s = str(s or "").strip()
    s = re.sub(r"^[A-Za-z]+\s+", "", s)           # "Tuesday 09/06/2026"
    return s.split(" ")[0]


def pdate(s, order="dmy"):
    s = _clean(s)
    if not s:
        return None
    fmts = ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d") if order == "dmy" \
        else ("%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y", "%Y-%m-%d")
    for f in fmts:
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def date_order(values):
    """Work out whether a sheet writes day-month or month-day: any first part > 12
    means dmy, any second part > 12 means mdy, otherwise dmy (UK default)."""
    firsts = seconds = 0
    for v in values:
        m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-]\d{2,4}$", _clean(v))
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        firsts += a > 12; seconds += b > 12
    return "mdy" if seconds and not firsts else "dmy"


def fdate(d):
    return d.strftime("%d/%m/%Y") if d else ""


def monday(d):
    return d - datetime.timedelta(days=d.weekday())


def find_header(rows):
    """Row index whose cells include something like 'date' and 'ref'/'reference'."""
    for i, r in enumerate(rows[:10]):
        low = [str(c).strip().lower() for c in r]
        if any("date" in c for c in low) and any("ref" in c for c in low):
            return i
    return 0


def col(hdr, *names):
    low = [str(c).strip().lower() for c in hdr]
    for n in names:
        for i, c in enumerate(low):
            if c == n:
                return i
    for n in names:
        for i, c in enumerate(low):
            if n in c:
                return i
    return None


def cell(r, i):
    return str(r[i]).strip() if i is not None and i < len(r) and r[i] is not None else ""


def load_source(svc, agent, sid, rng, agent_col):
    rows = svc.spreadsheets().values().get(spreadsheetId=sid, range=rng).execute().get("values", [])
    if not rows:
        return []
    h = find_header(rows); hdr = rows[h]
    c_date = col(hdr, "date"); c_ref = col(hdr, "crm reference", "tl ref", "reference", "ref")
    c_name = col(hdr, "client name", "client", "name"); c_src = col(hdr, "lead source", "source", "campaign")
    c_num = col(hdr, "contact number", "number", "phone"); c_adv = col(hdr, "advisor", "sfm")
    c_stage = col(hdr, "status", "stage"); c_agreed = col(hdr, "agreed date", "agreed")
    c_sip = col(hdr, "sip complete", "sip"); c_moc = col(hdr, "moc set", "moc")
    c_appr = col(hdr, "approval date", "approval"); c_apprflag = col(hdr, "approved")
    c_out = col(hdr, "outcome"); c_prop = col(hdr, "prop back", "prop"); c_part = col(hdr, "partner")
    c_notes = col(hdr, "notes"); c_agent = col(hdr, agent_col) if agent_col else None
    label = agent or rng.split("!")[0].strip("'")
    order = date_order(cell(r, c_date) for r in rows[h + 1:])
    out = []
    for r in rows[h + 1:]:
        if not any(str(c).strip() for c in r):
            continue
        d = pdate(cell(r, c_date), order)
        ref = cell(r, c_ref); name = cell(r, c_name)
        if not d or not (ref or name):            # skips "WC 24/08" divider rows and date-only rows
            continue
        ag = agent or (cell(r, c_agent).strip().title() if c_agent is not None and cell(r, c_agent) else "Unassigned")
        out.append({
            "date": d, "agent": ag, "ref": ref, "name": name, "source": cell(r, c_src), "number": cell(r, c_num),
            "advisor": cell(r, c_adv), "stage": cell(r, c_stage), "outcome": cell(r, c_out), "agreed": cell(r, c_agreed),
            "sip": cell(r, c_sip), "prop": cell(r, c_prop), "moc": cell(r, c_moc), "appr": cell(r, c_appr),
            "apprflag": cell(r, c_apprflag), "partner": cell(r, c_part), "notes": cell(r, c_notes), "from": label,
        })
    return out


def flags(c):
    """Reached-at-least flags. Later stages imply the earlier ones."""
    t = (c["stage"] or "").lower()
    def yes(v):
        """A stage column counts only if it holds a date or an explicit yes -
        Claire's tracker keeps product codes (MW / DMP / N/A) in those cells."""
        v = (v or "").strip().lower()
        return bool(pdate(v)) or v in ("yes", "y", "done", "complete", "completed", "true", "\u2713")
    appr = yes(c["appr"]) or yes(c["apprflag"]) or "approv" in t
    moc = appr or yes(c["moc"]) or "moc" in t
    sip = moc or yes(c["sip"]) or "sip" in t
    agreed = sip or yes(c["agreed"]) or "agreed" in t
    dnq = "dnq" in t or "not enough" in t
    dmp = "dmp" in t or "dmp" in (c["outcome"] or "").lower() or "dmp" in (c["advisor"] or "").lower()
    if dmp:                                   # DMPs are counted in their own column, not the IVA stages
        agreed = sip = moc = appr = False
    return {"agreed": agreed, "sip": sip, "moc": moc, "appr": appr, "dnq": dnq, "dmp": dmp}


def _snap_row(label, n):
    dmp = n["dmp"]; p = n["passed"] - dmp           # Leads Passed = everything that is not a DMP
    a, sp, m, ap, d = n["agreed"], n["sip"], n["moc"], n["appr"], n["dnq"]
    return [label, p, dmp, a, sp, m, ap, d, round(a / p, 4) if p else 0, round(sp / a, 4) if a else 0, round(ap / p, 4) if p else 0]


def build(svc):
    cases = []
    for agent, sid, rng, agent_col in SOURCES:
        try:
            got = load_source(svc, agent, sid, rng, agent_col)
            log.info("master-tracker: %s -> %d cases" % (agent or rng, len(got)))
            cases.extend(got)
        except Exception as e:
            log.warning("master-tracker: could not read %s (%s): %s" % (agent or rng, sid[:8], e))
    cases.sort(key=lambda c: (c["date"] or datetime.date(1900, 1, 1), c["agent"], c["ref"]), reverse=NEWEST_FIRST)
    # master: a "W/C dd/mm/yyyy (n cases)" band above each week's cases
    master = [HEAD]
    by_week = collections.OrderedDict()
    for c in cases:
        by_week.setdefault(monday(c["date"]), []).append(c)
    for wk, group in by_week.items():
        sun = wk + datetime.timedelta(days=6)
        master.append(["W/C %s   (%s - %s)   %d case%s" % (fdate(wk), wk.strftime("%d %b"), sun.strftime("%d %b"), len(group), "" if len(group) == 1 else "s")] + [""] * (len(HEAD) - 1))
        for c in group:
            master.append([fdate(c["date"]), "w/c " + fdate(wk), c["agent"], c["ref"], c["name"], c["source"],
                           c["number"], c["advisor"], c["stage"], c["outcome"], c["agreed"], c["sip"], c["prop"], c["moc"],
                           c["appr"], c["apprflag"], c["partner"], c["notes"], c["from"]])
    # weekly snapshot: cohort by week passed; one table per week + an OVERALL table on top
    per = collections.defaultdict(lambda: collections.Counter())
    for c in cases:
        f = flags(c)
        for key in ((monday(c["date"]), c["agent"]), ("ALL", c["agent"])):
            per[key]["passed"] += 1
            for name in ("agreed", "sip", "moc", "appr", "dnq", "dmp"):
                per[key][name] += 1 if f[name] else 0
    def table(title, wk):
        agents = sorted({k[1] for k in per if k[0] == wk})
        rows = [[title] + [""] * (len(SNAP_HEAD) - 1), list(SNAP_HEAD)]
        tot = collections.Counter()
        for ag in agents:
            n = per[(wk, ag)]; tot.update(n)
            rows.append(_snap_row(ag, n))
        rows.append(_snap_row("TOTAL", tot))
        rows.append([""] * len(SNAP_HEAD))
        return rows
    snap = table("OVERALL  -  all weeks", "ALL")
    weeks = sorted({k[0] for k in per if k[0] != "ALL"}, reverse=True)
    for wk in weeks:
        sun = wk + datetime.timedelta(days=6)
        snap += table("W/C %s   (%s - %s)" % (fdate(wk), wk.strftime("%d %b"), sun.strftime("%d %b")), wk)
    return master, snap, len(cases)


# ── sheet writing ──────────────────────────────────────────────────────────────
def _gid(svc, title):
    meta = svc.spreadsheets().get(spreadsheetId=ARKLE, fields="sheets(properties(title,sheetId))").execute()
    for sh in meta["sheets"]:
        if sh["properties"]["title"] == title:
            return sh["properties"]["sheetId"]
    return None


def ensure_tabs(svc):
    for title, ncols in ((MASTER_TAB, len(HEAD)), (SNAP_TAB, len(SNAP_HEAD))):
        if _gid(svc, title) is None:
            svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": [
                {"addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": 2000, "columnCount": ncols + 4, "frozenRowCount": 1}}}}]}).execute()
            log.info("master-tracker: created tab %s" % title)


def _rgb(hexs):
    return {"red": int(hexs[1:3], 16) / 255, "green": int(hexs[3:5], 16) / 255, "blue": int(hexs[5:7], 16) / 255}


def _fmt(bg, fg="#000000", bold=False):
    return {"backgroundColor": _rgb(bg), "textFormat": {"foregroundColor": _rgb(fg), "bold": bold}}


def _rule(gid, rng, cond, fmt):
    return {"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [dict(sheetId=gid, **rng)], "booleanRule": {"condition": cond, "format": fmt}}}}


# colours matching the Status dropdown chips (guide from Aaron 30/09)
BLUE, GREYBLUE, GREEN, PURPLE, YELLOW, RED, DARKRED = "#BFE1F6", "#C6DBE1", "#D4EDBC", "#E6CFF2", "#FFE5A0", "#E23A5A", "#B10202"
STAGE_COLOURS = [   # lowest priority first (rules are inserted at index 0, so the last one here wins ties)
    ("in pods", BLUE, "#0842A0"), ("lead passed", BLUE, "#0842A0"), ("callback", BLUE, "#0842A0"), ("agreed", BLUE, "#0842A0"),
    ("sip", BLUE, "#0842A0"), ("moc set", BLUE, "#0842A0"), ("prop back", BLUE, "#0842A0"),
    ("prop out", GREYBLUE, "#215A6C"),
    ("refresh", PURPLE, "#5A3286"), ("sched", PURPLE, "#5A3286"), ("awaiting", PURPLE, "#5A3286"), ("on hold", PURPLE, "#5A3286"),
    ("dnq", DARKRED, "#FFFFFF"), ("money wellness", DARKRED, "#FFFFFF"), ("already in iva", DARKRED, "#FFFFFF"),
    ("dnc", RED, "#000000"),
    ("lost contact", YELLOW, "#473821"),
    ("agreed dmp", GREEN, "#11734B"), ("approved", GREEN, "#11734B"),
]


def ensure_formatting(svc):
    """One-off, idempotent: only adds rules to a tab that has NONE. Never re-adds."""
    meta = svc.spreadsheets().get(spreadsheetId=ARKLE, fields="sheets(properties(title,sheetId),conditionalFormats)").execute()
    info = {sh["properties"]["title"]: (sh["properties"]["sheetId"], len(sh.get("conditionalFormats", []))) for sh in meta["sheets"]}
    reqs = []
    gid, n = info.get(MASTER_TAB, (None, 1))
    if gid is not None and n == 0:
        stage = {"startRowIndex": 1, "startColumnIndex": 8, "endColumnIndex": 9}           # column I = Stage
        for word, bg, fg in STAGE_COLOURS:
            reqs.append(_rule(gid, stage, {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": word}]}, _fmt(bg, fg)))
        band = {"startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(HEAD)}
        reqs.append(_rule(gid, band, {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=LEFT($A2,4)="W/C "'}]}, _fmt("#434343", "#FFFFFF", True)))
        reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1}, "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "backgroundColor": _rgb("#EFEFEF")}}, "fields": "userEnteredFormat(textFormat.bold,backgroundColor)"}})
        reqs.append({"autoResizeDimensions": {"dimensions": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": len(HEAD)}}})
    gid, n = info.get(SNAP_TAB, (None, 1))
    if gid is not None and n == 0:
        full = {"startRowIndex": 0, "startColumnIndex": 0, "endColumnIndex": len(SNAP_HEAD)}
        reqs.append(_rule(gid, full, {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=$A1="Agent"'}]}, _fmt("#EFEFEF", "#000000", True)))
        reqs.append(_rule(gid, full, {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=$A1="TOTAL"'}]}, _fmt("#D9D9D9", "#000000", True)))
        reqs.append(_rule(gid, full, {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=OR(LEFT($A1,4)="W/C ",LEFT($A1,7)="OVERALL")'}]}, _fmt("#434343", "#FFFFFF", True)))
        reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 0, "startColumnIndex": 8, "endColumnIndex": 11},
                                    "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}}, "fields": "userEnteredFormat.numberFormat"}})
        reqs.append({"updateSheetProperties": {"properties": {"sheetId": gid, "gridProperties": {"frozenRowCount": 0}}, "fields": "gridProperties.frozenRowCount"}})
        reqs.append({"autoResizeDimensions": {"dimensions": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": len(SNAP_HEAD)}}})
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": reqs}).execute()
        log.info("master-tracker: formatting rules added (%d requests)" % len(reqs))


MAX_ROWS = 5000
ALLOWED_TABS = {"MASTER TRACKER", "WEEKLY SNAPSHOT"}


def write(svc, master, snap):
    # hard guards: only ever touch the two named tabs, never write something huge
    if MASTER_TAB not in ALLOWED_TABS or SNAP_TAB not in ALLOWED_TABS:
        raise RuntimeError("refusing to write to %s / %s" % (MASTER_TAB, SNAP_TAB))
    if len(master) > MAX_ROWS or len(snap) > MAX_ROWS:
        raise RuntimeError("refusing to write %d/%d rows (cap %d)" % (len(master), len(snap), MAX_ROWS))
    ensure_tabs(svc)
    ensure_formatting(svc)
    stamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    svc.spreadsheets().values().batchClear(spreadsheetId=ARKLE, body={"ranges": ["'%s'!A:T" % MASTER_TAB, "'%s'!A:L" % SNAP_TAB]}).execute()
    svc.spreadsheets().values().batchUpdate(spreadsheetId=ARKLE, body={"valueInputOption": "RAW", "data": [
        {"range": "'%s'!A1" % MASTER_TAB, "values": master},
        {"range": "'%s'!A1" % SNAP_TAB, "values": snap},
        {"range": "'%s'!V1:W1" % MASTER_TAB, "values": [["Updated " + stamp, _last_hash or ""]]},
    ]}).execute()


def run(svc, dry_run=True):
    global _last_hash
    master, snap, n = build(svc)
    h = hashlib.md5(repr(master + snap).encode()).hexdigest()
    if _last_hash is None:
        try:   # survive restarts: hash of the last write lives in MASTER TRACKER!W1
            v = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range="'%s'!W1" % MASTER_TAB).execute().get("values", [[""]])
            _last_hash = v[0][0] if v and v[0] else ""
        except Exception:
            _last_hash = ""
    if h == _last_hash:
        log.info("master-tracker: no change (%d cases)" % n)
        return False
    if dry_run:
        log.info("master-tracker [DRY RUN]: would write %d cases, %d snapshot rows" % (n, len(snap) - 1))
        return False
    _last_hash = h
    write(svc, master, snap)
    log.info("master-tracker: wrote %d cases + %d snapshot rows" % (n, len(snap) - 1))
    return True


def tick(get_svc):
    """Called every poller cycle; runs at most every INTERVAL_S."""
    global _last_tick
    if not ENABLED or time.time() - _last_tick < INTERVAL_S:
        return
    _last_tick = time.time()
    try:
        run(get_svc(), dry_run=DRY_RUN)
    except Exception as e:
        log.warning("master-tracker: failed: %s" % e)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    sys.path.insert(0, ".")
    import main
    svc = main.get_sheets_service()
    master, snap, n = build(svc)
    agents = collections.Counter(r[2] for r in master[1:] if r[2])
    print("cases:", n, "| by agent:", dict(agents), "| snapshot rows:", len(snap))
    print("first 3:", master[1:4]); print("last 3:", master[-3:])
    print("snapshot top:"); [print("   ", r) for r in snap[:9]]
    nodate = [r for r in master[1:] if not r[0]]  # section bands start with W/C so never blank
    if nodate:
        print("WARNING rows with unreadable date:", len(nodate), [(r[2], r[3], r[4]) for r in nodate[:5]])
    if os.getenv("GO") != "1":
        print("preview only - GO=1 writes the two tabs"); sys.exit()
    _last_hash = hashlib.md5(repr(master + snap).encode()).hexdigest()
    write(svc, master, snap)
    print("written:", MASTER_TAB, "and", SNAP_TAB)
