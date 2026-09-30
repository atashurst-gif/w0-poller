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
SNAP_HEAD = ["Week (Mon)", "Week", "Agent", "Leads Passed", "Agreed", "SIP Complete", "MOC Set", "Approved", "DNQ",
             "Lead > Agreed", "Agreed > SIP", "Lead > Approved"]

_last_tick = 0.0
_last_hash = None


# ── helpers ────────────────────────────────────────────────────────────────────
def pdate(s):
    s = str(s or "").strip()
    if not s:
        return None
    s = re.sub(r"^[A-Za-z]+\s+", "", s)           # "Tuesday 09/06/2026"
    s = s.split(" ")[0]
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y", "%m-%d-%Y", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


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
    out = []
    for r in rows[h + 1:]:
        if not any(str(c).strip() for c in r):
            continue
        d = pdate(cell(r, c_date))
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
    return {"agreed": agreed, "sip": sip, "moc": moc, "appr": appr, "dnq": dnq}


def build(svc):
    cases = []
    for agent, sid, rng, agent_col in SOURCES:
        try:
            got = load_source(svc, agent, sid, rng, agent_col)
            log.info("master-tracker: %s -> %d cases" % (agent or rng, len(got)))
            cases.extend(got)
        except Exception as e:
            log.warning("master-tracker: could not read %s (%s): %s" % (agent or rng, sid[:8], e))
    cases.sort(key=lambda c: (c["date"] or datetime.date(1900, 1, 1), c["agent"], c["ref"]))
    master = [HEAD]
    for c in cases:
        wk = monday(c["date"]) if c["date"] else None
        master.append([fdate(c["date"]), ("w/c " + fdate(wk)) if wk else "", c["agent"], c["ref"], c["name"], c["source"],
                       c["number"], c["advisor"], c["stage"], c["outcome"], c["agreed"], c["sip"], c["prop"], c["moc"],
                       c["appr"], c["apprflag"], c["partner"], c["notes"], c["from"]])
    # weekly snapshot: cohort by date passed, per agent + ALL, newest week first
    per = collections.defaultdict(lambda: collections.Counter())
    for c in cases:
        if not c["date"]:
            continue
        wk = monday(c["date"]); f = flags(c)
        for ag in (c["agent"], "ALL"):
            k = (wk, ag); per[k]["passed"] += 1
            for name in ("agreed", "sip", "moc", "appr", "dnq"):
                per[k][name] += 1 if f[name] else 0
    snap = [SNAP_HEAD]
    weeks = sorted({k[0] for k in per}, reverse=True)
    for wk in weeks:
        agents = sorted({k[1] for k in per if k[0] == wk and k[1] != "ALL"}) + ["ALL"]
        for ag in agents:
            n = per[(wk, ag)]
            p, a, s, m, ap, d = n["passed"], n["agreed"], n["sip"], n["moc"], n["appr"], n["dnq"]
            snap.append([fdate(wk), "w/c " + fdate(wk), ag, p, a, s, m, ap, d,
                         round(a / p, 4) if p else 0, round(s / a, 4) if a else 0, round(ap / p, 4) if p else 0])
    return master, snap, len(cases)


# ── sheet writing ──────────────────────────────────────────────────────────────
def _gid(svc, title):
    meta = svc.spreadsheets().get(spreadsheetId=ARKLE, fields="sheets(properties(title,sheetId))").execute()
    for sh in meta["sheets"]:
        if sh["properties"]["title"] == title:
            return sh["properties"]["sheetId"]
    return None


def ensure_tabs(svc):
    reqs = []
    for title, ncols, pct_cols in ((MASTER_TAB, len(HEAD), None), (SNAP_TAB, len(SNAP_HEAD), (9, 12))):
        gid = _gid(svc, title)
        if gid is None:
            res = svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": [
                {"addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": 2000, "columnCount": ncols + 4, "frozenRowCount": 1}}}}]}).execute()
            gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]
            reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1},
                                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}}, "fields": "userEnteredFormat.textFormat.bold"}})
            if pct_cols:
                reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 1, "startColumnIndex": pct_cols[0], "endColumnIndex": pct_cols[1]},
                                            "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}}, "fields": "userEnteredFormat.numberFormat"}})
            log.info("master-tracker: created tab %s" % title)
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": reqs}).execute()


MAX_ROWS = 5000
ALLOWED_TABS = {"MASTER TRACKER", "WEEKLY SNAPSHOT"}


def write(svc, master, snap):
    # hard guards: only ever touch the two named tabs, never write something huge
    if MASTER_TAB not in ALLOWED_TABS or SNAP_TAB not in ALLOWED_TABS:
        raise RuntimeError("refusing to write to %s / %s" % (MASTER_TAB, SNAP_TAB))
    if len(master) > MAX_ROWS or len(snap) > MAX_ROWS:
        raise RuntimeError("refusing to write %d/%d rows (cap %d)" % (len(master), len(snap), MAX_ROWS))
    ensure_tabs(svc)
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
    agents = collections.Counter(r[2] for r in master[1:])
    print("cases:", n, "| by agent:", dict(agents), "| snapshot rows:", len(snap) - 1)
    print("first 3:", master[1:4]); print("last 3:", master[-3:])
    print("snapshot top:", snap[1:6])
    nodate = [r for r in master[1:] if not r[0]]
    if nodate:
        print("WARNING rows with unreadable date:", len(nodate), [(r[2], r[3], r[4]) for r in nodate[:5]])
    if os.getenv("GO") != "1":
        print("preview only - GO=1 writes the two tabs"); sys.exit()
    _last_hash = hashlib.md5(repr(master + snap).encode()).hexdigest()
    write(svc, master, snap)
    print("written:", MASTER_TAB, "and", SNAP_TAB)
