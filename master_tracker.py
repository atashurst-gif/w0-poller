#!/usr/bin/env python3
"""MASTER TRACKER + WEEKLY SNAPSHOT for the Arkle workbook.

Collects every case from Claire's and Bianca's Case Trackers and the workbook's
own 'Leads Passed' tab (agent = Lead Gen column) and keeps three tabs current:
  MASTER TRACKER   every case, newest week first, weekly bands, coloured Stage
  MT DATA          one row per case as plain values (hidden) - the snapshot's source
  WEEKLY SNAPSHOT  one table driven by a week dropdown (This week by default) +
                   an all-time table underneath; light COUNTIFS over MT DATA
Only rewrites when the data changed. No charts. Colour rules are versioned and
applied once per version.

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
MASTER_TAB = "MASTER TRACKER"
SNAP_TAB   = "WEEKLY SNAPSHOT"
DATA_TAB   = "MT DATA"
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
        "Stage", "Type", "Outcome", "Agreed Date", "SIP Complete", "Prop Back", "MOC Set", "Approval Date", "Approved",
        "Partner", "Notes", "From"]                       # 20 columns: A..T   (Stage stays in column I)
DATA_HEAD = ["Week", "Date", "Agent", "Type", "Agreed", "SIP", "MOC", "Approved", "DNQ", "TL Ref", "Client"]
SNAP_HEAD = ["Agent", "IVA Passed", "DMP Passed", "IVA Agreed", "DMP Agreed", "SIP Complete", "MOC Set", "Approved",
             "DNQ", "IVA > Agreed", "Agreed > SIP", "IVA > Approved"]      # 12 columns: A..L
NEWEST_FIRST = os.getenv("MT_NEWEST_FIRST", "1") == "1"
FORMAT_VERSION = "fmt3"                                   # bump to re-apply colour rules once

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
    """Reached-at-least flags. Later stages imply the earlier ones. DMP cases keep
    'agreed' (Agreed DMP) but never count in the IVA stages."""
    t = (c["stage"] or "").lower()
    def yes(v):
        """A stage column counts only if it holds a date or an explicit yes -
        Claire's tracker keeps product codes (MW / DMP / N/A) in those cells."""
        v = (v or "").strip().lower()
        return bool(pdate(v)) or v in ("yes", "y", "done", "complete", "completed", "true", "✓")
    dmp = "dmp" in t or "dmp" in (c["outcome"] or "").lower() or "dmp" in (c["advisor"] or "").lower()
    appr = yes(c["appr"]) or yes(c["apprflag"]) or "approv" in t
    moc = appr or yes(c["moc"]) or "moc" in t
    sip = moc or yes(c["sip"]) or "sip" in t
    agreed = sip or yes(c["agreed"]) or "agreed" in t
    dnq = "dnq" in t or "not enough" in t
    if dmp:
        sip = moc = appr = False
    return {"agreed": agreed, "sip": sip, "moc": moc, "appr": appr, "dnq": dnq, "dmp": dmp}


# ── build ──────────────────────────────────────────────────────────────────────
def build(svc):
    cases = []
    for agent, sid, rng, agent_col in SOURCES:
        try:
            got = load_source(svc, agent, sid, rng, agent_col)
            log.info("master-tracker: %s -> %d cases" % (agent or rng, len(got)))
            cases.extend(got)
        except Exception as e:
            log.warning("master-tracker: could not read %s (%s): %s" % (agent or rng, sid[:8], e))
    cases.sort(key=lambda c: (c["date"], c["agent"], c["ref"]), reverse=NEWEST_FIRST)
    for c in cases:
        c["f"] = flags(c); c["type"] = "DMP" if c["f"]["dmp"] else "IVA"

    # MASTER TRACKER: a band above each week's cases
    master = [HEAD]
    by_week = collections.OrderedDict()
    for c in cases:
        by_week.setdefault(monday(c["date"]), []).append(c)
    for wk, group in by_week.items():
        sun = wk + datetime.timedelta(days=6)
        master.append(["W/C %s   (%s - %s)   %d case%s" % (fdate(wk), wk.strftime("%d %b"), sun.strftime("%d %b"),
                                                          len(group), "" if len(group) == 1 else "s")] + [""] * (len(HEAD) - 1))
        for c in group:
            master.append([fdate(c["date"]), "w/c " + fdate(wk), c["agent"], c["ref"], c["name"], c["source"],
                           c["number"], c["advisor"], c["stage"], c["type"], c["outcome"], c["agreed"], c["sip"], c["prop"],
                           c["moc"], c["appr"], c["apprflag"], c["partner"], c["notes"], c["from"]])

    # MT DATA: one row per case, ISO dates so the sheet stores real dates
    data = [DATA_HEAD]
    for c in cases:
        f = c["f"]
        data.append([monday(c["date"]).isoformat(), c["date"].isoformat(), c["agent"], c["type"],
                     int(f["agreed"]), int(f["sip"]), int(f["moc"]), int(f["appr"]), int(f["dnq"]), c["ref"], c["name"]])
    weeks = sorted(by_week.keys(), reverse=True)
    weeklist = [["Week options", "Monday"], ["This week", ""], ["Last week", ""], ["All time", ""]] + \
               [["W/C " + fdate(wk), wk.isoformat()] for wk in weeks]
    agents = sorted({c["agent"] for c in cases})
    return master, data, weeklist, agents, len(cases)


def snapshot_rows(agents, current_choice):
    """WEEKLY SNAPSHOT contents: dropdown row, week table, all-time table. Formulas
    are plain COUNTIFS over MT DATA; $N$1/$N$2 hold the selected week's dates."""
    D = "'%s'" % DATA_TAB
    def crit(week, agent_cell, extra=""):
        base = "%s!$C:$C,%s" % (D, agent_cell)
        if week:
            base += ',%s!$B:$B,">="&$N$1,%s!$B:$B,"<="&$N$2' % (D, D)
        return base + extra
    def table_rows(start_row, week):
        rows = [list(SNAP_HEAD)]
        first = start_row + 1
        for i, ag in enumerate(agents):
            r = first + i; a = "$A%d" % r
            iva = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"IVA"' % D)
            dmp = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"DMP"' % D)
            iva_ag = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"IVA",%s!$E:$E,1' % (D, D))
            dmp_ag = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"DMP",%s!$E:$E,1' % (D, D))
            sip = '=COUNTIFS(%s)' % crit(week, a, ',%s!$F:$F,1' % D)
            moc = '=COUNTIFS(%s)' % crit(week, a, ',%s!$G:$G,1' % D)
            apr = '=COUNTIFS(%s)' % crit(week, a, ',%s!$H:$H,1' % D)
            dnq = '=COUNTIFS(%s)' % crit(week, a, ',%s!$I:$I,1' % D)
            rows.append([ag, iva, dmp, iva_ag, dmp_ag, sip, moc, apr, dnq,
                         '=IF(B%d=0,"",D%d/B%d)' % (r, r, r), '=IF(D%d=0,"",F%d/D%d)' % (r, r, r), '=IF(B%d=0,"",H%d/B%d)' % (r, r, r)])
        last = first + len(agents) - 1; t = last + 1
        tot = ["TOTAL"] + ["=SUM(%s%d:%s%d)" % (col, first, col, last) for col in "BCDEFGHI"]
        tot += ['=IF(B%d=0,"",D%d/B%d)' % (t, t, t), '=IF(D%d=0,"",F%d/D%d)' % (t, t, t), '=IF(B%d=0,"",H%d/B%d)' % (t, t, t)]
        rows.append(tot)
        return rows
    start = ('=IF($B$1="This week",TODAY()-WEEKDAY(TODAY(),2)+1,IF($B$1="Last week",TODAY()-WEEKDAY(TODAY(),2)-6,'
             'IF($B$1="All time",DATE(2024,1,1),IFERROR(VLOOKUP($B$1,%s!$M:$N,2,FALSE),TODAY()-WEEKDAY(TODAY(),2)+1))))' % D)
    end = '=IF($B$1="All time",TODAY()+7,$N$1+6)'
    rows = [["WEEK", current_choice or "This week", "", '="Showing "&TEXT($N$1,"ddd dd/mm")&" to "&TEXT($N$2,"ddd dd/mm")'],
            [""] * len(SNAP_HEAD)]
    rows += table_rows(3, week=True)                      # header on row 3, agents from row 4
    n = len(agents)
    rows += [[""] * len(SNAP_HEAD), ["ALL WEEKS  -  every case since the trackers began"] + [""] * (len(SNAP_HEAD) - 1)]
    rows += table_rows(len(rows) + 1, week=False)
    # 1-based (header row, TOTAL row) of each table: week table header 3, agents 4.., TOTAL n+4;
    # blank n+5, ALL WEEKS title n+6, header n+7, agents.., TOTAL 2n+8
    return rows, {"N1": start, "N2": end, "week_rows": (3, n + 4), "all_rows": (n + 7, 2 * n + 8)}


# ── sheet writing ──────────────────────────────────────────────────────────────
MAX_ROWS = 5000
ALLOWED_TABS = {MASTER_TAB, SNAP_TAB, DATA_TAB}


def _meta(svc):
    return svc.spreadsheets().get(spreadsheetId=ARKLE, fields="sheets(properties(title,sheetId,hidden),conditionalFormats)").execute()


def _gids(svc):
    return {sh["properties"]["title"]: sh["properties"]["sheetId"] for sh in _meta(svc)["sheets"]}


def ensure_tabs(svc):
    gids = _gids(svc)
    for title, ncols in ((MASTER_TAB, len(HEAD)), (SNAP_TAB, 16), (DATA_TAB, 16)):
        if title not in gids:
            svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": [
                {"addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": 2000, "columnCount": ncols + 4}}}}]}).execute()
            log.info("master-tracker: created tab %s" % title)


def _rgb(hexs):
    return {"red": int(hexs[1:3], 16) / 255, "green": int(hexs[3:5], 16) / 255, "blue": int(hexs[5:7], 16) / 255}


def _fmt(bg, fg="#000000", bold=False):
    return {"backgroundColor": _rgb(bg), "textFormat": {"foregroundColor": _rgb(fg), "bold": bold}}


def _rule(gid, rng, cond, fmt):
    return {"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [dict(sheetId=gid, **rng)], "booleanRule": {"condition": cond, "format": fmt}}}}


def _cell(gid, r0, r1, c0, c1, fmt, fields):
    return {"repeatCell": {"range": {"sheetId": gid, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1},
                           "cell": {"userEnteredFormat": fmt}, "fields": fields}}


NAVY, NAVY_LIGHT, ROW_ALT, YELLOW_CELL, TEAL = "#1F4E79", "#D9E2F3", "#F3F7FB", "#FFF2CC", "#0B6E4F"
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
    """Versioned one-off. Runs only when MASTER TRACKER!X1 != FORMAT_VERSION:
    clears the tabs' colour rules and re-adds them, sets the dropdown, widths,
    hides MT DATA. Never runs on an ordinary refresh."""
    try:
        v = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range="'%s'!X1" % MASTER_TAB).execute().get("values", [[""]])
        if v and v[0] and v[0][0] == FORMAT_VERSION:
            return
    except Exception:
        pass
    meta = _meta(svc)
    info = {sh["properties"]["title"]: (sh["properties"]["sheetId"], len(sh.get("conditionalFormats", []))) for sh in meta["sheets"]}
    reqs = []
    for title in (MASTER_TAB, SNAP_TAB):
        gid, n = info[title]
        for _ in range(n):
            reqs.append({"deleteConditionalFormatRule": {"sheetId": gid, "index": 0}})
    mg, _ = info[MASTER_TAB]
    stage = {"startRowIndex": 1, "startColumnIndex": 8, "endColumnIndex": 9}                 # column I = Stage
    for word, bg, fg in STAGE_COLOURS:
        reqs.append(_rule(mg, stage, {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": word}]}, _fmt(bg, fg)))
    # DMP rows get a soft tint everywhere except the Stage cell, so the stage colour still shows
    reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": mg, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8},
                   {"sheetId": mg, "startRowIndex": 1, "startColumnIndex": 9, "endColumnIndex": len(HEAD)}],
        "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=$J2="DMP"'}]},
                        "format": {"backgroundColor": _rgb("#FBF3E6")}}}}})
    band = {"startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(HEAD)}
    reqs.append(_rule(mg, band, {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=LEFT($A2,4)="W/C "'}]}, _fmt(NAVY, "#FFFFFF", True)))
    reqs.append(_cell(mg, 0, 1, 0, len(HEAD), _fmt(NAVY_LIGHT, "#1F4E79", True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": mg, "gridProperties": {"frozenRowCount": 1}}, "fields": "gridProperties.frozenRowCount"}})
    reqs.append({"autoResizeDimensions": {"dimensions": {"sheetId": mg, "dimension": "COLUMNS", "startIndex": 0, "endIndex": len(HEAD)}}})
    sg, _ = info[SNAP_TAB]
    reqs.append({"setDataValidation": {"range": {"sheetId": sg, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 1, "endColumnIndex": 2},
                                       "rule": {"condition": {"type": "ONE_OF_RANGE", "values": [{"userEnteredValue": "='%s'!$M$2:$M$300" % DATA_TAB}]},
                                                "showCustomUi": True, "strict": False}}})
    widths = [110] + [82] * 8 + [95, 95, 105]
    for i, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {"range": {"sheetId": sg, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
                                                   "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {"range": {"sheetId": sg, "dimension": "COLUMNS", "startIndex": 12, "endIndex": 14},
                                               "properties": {"pixelSize": 60}, "fields": "pixelSize"}})
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": sg, "gridProperties": {"frozenRowCount": 0, "hideGridlines": True}}, "fields": "gridProperties(frozenRowCount,hideGridlines)"}})
    dg, _ = info[DATA_TAB]
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": dg, "hidden": True}, "fields": "hidden"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": reqs}).execute()
    svc.spreadsheets().values().update(spreadsheetId=ARKLE, range="'%s'!X1" % MASTER_TAB, valueInputOption="RAW", body={"values": [[FORMAT_VERSION]]}).execute()
    log.info("master-tracker: formatting %s applied (%d requests)" % (FORMAT_VERSION, len(reqs)))


def snapshot_format_requests(gid, layout, nrows):
    """Cell colours for the snapshot's fixed layout - a handful of repeatCell
    requests, re-applied only when the tab is rewritten."""
    F = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)"
    reqs = [_cell(gid, 0, nrows + 2, 0, 14, {"backgroundColor": _rgb("#FFFFFF"), "textFormat": {"bold": False, "foregroundColor": _rgb("#000000")}}, "userEnteredFormat(backgroundColor,textFormat)")]
    reqs.append(_cell(gid, 0, 1, 0, 1, _fmt("#FFFFFF", NAVY, True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append(_cell(gid, 0, 1, 1, 2, _fmt(YELLOW_CELL, "#000000", True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append(_cell(gid, 0, 1, 3, 4, {"textFormat": {"italic": True, "foregroundColor": _rgb("#555555")}}, "userEnteredFormat.textFormat"))
    for (h0, t) in (layout["week_rows"], layout["all_rows"]):
        h = h0 - 1                                                     # 0-based header row
        reqs.append(_cell(gid, h, h + 1, 0, len(SNAP_HEAD), dict(_fmt(NAVY, "#FFFFFF", True), horizontalAlignment="CENTER"), F))
        for r in range(h + 1, t - 1):
            if (r - h) % 2 == 0:
                reqs.append(_cell(gid, r, r + 1, 0, len(SNAP_HEAD), {"backgroundColor": _rgb(ROW_ALT)}, "userEnteredFormat.backgroundColor"))
        reqs.append(_cell(gid, t - 1, t, 0, len(SNAP_HEAD), _fmt(NAVY_LIGHT, NAVY, True), "userEnteredFormat(backgroundColor,textFormat)"))
        reqs.append(_cell(gid, h + 1, t, 1, 9, {"horizontalAlignment": "CENTER"}, "userEnteredFormat.horizontalAlignment"))
        reqs.append(_cell(gid, h + 1, t, 9, 12, {"horizontalAlignment": "CENTER", "numberFormat": {"type": "PERCENT", "pattern": "0%"}}, "userEnteredFormat(horizontalAlignment,numberFormat)"))
    a0 = layout["all_rows"][0] - 2                                     # the ALL WEEKS title row (0-based)
    reqs.append(_cell(gid, a0, a0 + 1, 0, len(SNAP_HEAD), _fmt(TEAL, "#FFFFFF", True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append(_cell(gid, 0, 2, 12, 14, {"textFormat": {"foregroundColor": _rgb("#999999"), "fontSize": 8}, "numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}}, "userEnteredFormat(textFormat,numberFormat)"))
    return reqs


def write(svc, master, data, weeklist, agents):
    if len(master) > MAX_ROWS or len(data) > MAX_ROWS:
        raise RuntimeError("refusing to write %d/%d rows (cap %d)" % (len(master), len(data), MAX_ROWS))
    ensure_tabs(svc)
    ensure_formatting(svc)
    gids = _gids(svc)
    # keep whatever week the user has selected in the dropdown
    try:
        cur = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range="'%s'!B1" % SNAP_TAB).execute().get("values", [[""]])
        choice = (cur[0][0] if cur and cur[0] else "") or "This week"
    except Exception:
        choice = "This week"
    snap, layout = snapshot_rows(agents, choice)
    stamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    svc.spreadsheets().values().batchClear(spreadsheetId=ARKLE, body={"ranges": [
        "'%s'!A:T" % MASTER_TAB, "'%s'!A1:L200" % SNAP_TAB, "'%s'!A:N" % DATA_TAB]}).execute()
    svc.spreadsheets().values().batchUpdate(spreadsheetId=ARKLE, body={"valueInputOption": "RAW", "data": [
        {"range": "'%s'!A1" % MASTER_TAB, "values": master},
        {"range": "'%s'!V1:W1" % MASTER_TAB, "values": [["Updated " + stamp, _last_hash or ""]]},
    ]}).execute()
    svc.spreadsheets().values().batchUpdate(spreadsheetId=ARKLE, body={"valueInputOption": "USER_ENTERED", "data": [
        {"range": "'%s'!A1" % DATA_TAB, "values": data},
        {"range": "'%s'!M1" % DATA_TAB, "values": weeklist},
        {"range": "'%s'!A1" % SNAP_TAB, "values": snap},
        {"range": "'%s'!M1:N2" % SNAP_TAB, "values": [["start", layout["N1"]], ["end", layout["N2"]]]},
    ]}).execute()
    svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": snapshot_format_requests(gids[SNAP_TAB], layout, len(snap))}).execute()


def run(svc, dry_run=True):
    global _last_hash
    master, data, weeklist, agents, n = build(svc)
    h = hashlib.md5(repr([master, data, weeklist, agents, FORMAT_VERSION]).encode()).hexdigest()
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
        log.info("master-tracker [DRY RUN]: would write %d cases" % n)
        return False
    _last_hash = h
    write(svc, master, data, weeklist, agents)
    log.info("master-tracker: wrote %d cases (%d agents, %d weeks)" % (n, len(agents), len(weeklist) - 4))
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
    master, data, weeklist, agents, n = build(svc)
    print("cases:", n, "| agents:", agents, "| weeks:", len(weeklist) - 4)
    print("newest band:", master[1][0]); print("first case:", master[2])
    print("types:", collections.Counter(r[3] for r in data[1:]))
    if os.getenv("GO") != "1":
        print("preview only - GO=1 writes the tabs"); sys.exit()
    _last_hash = hashlib.md5(repr([master, data, weeklist, agents, FORMAT_VERSION]).encode()).hexdigest()
    write(svc, master, data, weeklist, agents)
    print("written:", MASTER_TAB, ",", SNAP_TAB, "and", DATA_TAB)
