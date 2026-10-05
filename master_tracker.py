#!/usr/bin/env python3
"""WEEKLY SNAPSHOT for the Arkle workbook, built from ONE shared tracker tab.

Everyone (Josh, Claire, Bianca, ...) logs cases on the same tab, 'MASTER TRACKER'
(Josh's old 'Leads Passed' tab). This module only READS that tab - it never
writes to it, formats it or clears it - and keeps two tabs current:
  MT DATA          one row per case as plain values (hidden) - the snapshot's source
  WEEKLY SNAPSHOT  per-person table driven by a week dropdown + an all-time table;
                   people are split by the tracker's 'Lead Gen' column
Arkle's 'Approved' tab stamps approvals onto tracker cases; approvals that are not
on the tracker are counted too (row 'Pre-tracker' unless the Approved tab names a
lead gen who is on the tracker) - MT_LEGACY=0 leaves those out.
No other spreadsheet is read (the separate Claire / Bianca tracker files are
disconnected).

Safety
  * all-or-nothing: the snapshot is rewritten only when the tracker AND the
    Approved tab were both read in full. A timeout or any other read error leaves
    the snapshot exactly as it was and the refresh is retried RETRY_S later.
  * while the tab called 'MASTER TRACKER' is missing or is still the old generated
    list, nothing is written at all ("waiting").
  * every write goes through _ours(): only WEEKLY SNAPSHOT and MT DATA can be
    written, whatever happens.
  * the old MASTER_TRACKER_ENABLED / MASTER_TRACKER_DRY_RUN variables are ignored
    on purpose (set MASTER_TRACKER_DRY_RUN=1 on Railway so an old build can never
    write again). This build uses MT_ENABLED (default 1) and MT_DRY_RUN (default 0).

Standalone (from ~/Desktop/w0-poller):
  python3 master_tracker.py          # preview: counts per person, writes nothing
  GO=1 python3 master_tracker.py     # refresh WEEKLY SNAPSHOT + MT DATA now
Inside the poller: master_tracker.tick(get_sheets_service) every MT_INTERVAL_S.
"""
import os, sys, re, time, hashlib, datetime, logging, collections

try:
    from zoneinfo import ZoneInfo
    UK = ZoneInfo("Europe/London")
except Exception:                                   # pragma: no cover
    UK = None

log = logging.getLogger(__name__)

ARKLE       = os.getenv("CBS_APPS2_SHEET_ID", "16qnJ842lFAo-4FVhloipJMFdCs7Spc2hFgoWnVGwVUs")
TRACKER_TAB = os.getenv("MT_TRACKER_TAB", "MASTER TRACKER")    # the shared tab people type into: READ ONLY
SNAP_TAB    = "WEEKLY SNAPSHOT"
DATA_TAB    = "MT DATA"
APPROVED_RNG = "'Approved'!A:M"                                 # Arkle's MOC Approved list
ENABLED     = os.getenv("MT_ENABLED", "1") == "1"
DRY_RUN     = os.getenv("MT_DRY_RUN", "0") == "1"
INTERVAL_S  = int(os.getenv("MT_INTERVAL_S", "900"))
RETRY_S     = int(os.getenv("MT_RETRY_S", "120"))               # after a failed / waiting cycle
LEGACY      = os.getenv("MT_LEGACY", "1") == "1"                # count Arkle approvals that aren't on the tracker
NEWEST_FIRST = os.getenv("MT_NEWEST_FIRST", "1") == "1"
CODE_VERSION = "src3"                                           # stamped in MT DATA!Q4 - one_tab.py checks it
FORMAT_VERSION = "fmt5"                                         # bump to re-apply the snapshot layout once
STATE_RNG   = "'%s'!P1:Q4" % DATA_TAB                           # updated / hash / format / code
OURS        = (SNAP_TAB, DATA_TAB)                              # the only tabs this module may write

DATA_HEAD = ["Week", "Date", "Agent", "Type", "Agreed", "SIP", "MOC", "Approved", "DNQ", "TL Ref", "Client", "Legacy"]
SNAP_HEAD = ["Agent", "IVA Passed", "DMP Passed", "IVA Agreed", "DMP Agreed", "SIP Complete", "MOC Set", "Approved",
             "DNQ", "IVA > Agreed", "Agreed > SIP", "IVA > Approved"]      # 12 columns: A..L
MAX_ROWS = 5000

_last_tick = 0.0
_last_hash = None
_stamped = False


class NotReady(Exception):
    """The shared tracker tab isn't there (yet) - nothing may be written."""


# ── helpers ────────────────────────────────────────────────────────────────────
def _now():
    return datetime.datetime.now(UK) if UK else datetime.datetime.now()


def _clean(s):
    s = str(s or "").strip()
    s = re.sub(r"^[A-Za-z]+\s+", "", s)           # "Tuesday 09/06/2026"
    return s.split(" ")[0]


def pdate(s):
    """Typed text -> date, day first (UK)."""
    s = _clean(s)
    if not s:
        return None
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


EPOCH = datetime.date(1899, 12, 30)                 # Google Sheets day 0


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def from_serial(n):
    try:
        d = EPOCH + datetime.timedelta(days=int(n))
    except (OverflowError, ValueError):
        return None
    return d if 2015 <= d.year <= 2040 else None


def to_date(v):
    """A tracker date cell. Read unformatted, so a real date arrives as its serial
    number and the way the cell is displayed never matters; typed text is read
    day-first."""
    if is_num(v):
        return from_serial(v)
    if isinstance(v, bool) or v is None:
        return None
    return pdate(v)


def text(v):
    """Any cell as text: whole numbers without '.0', ticked boxes as Yes."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Yes" if v else ""
    if is_num(v):
        return str(int(v)) if float(v) == int(v) else str(v)
    return str(v).strip()


def date_text(v):
    """A stage-date cell (Agreed Date, SIP Complete, MOC Set...): a real date as
    dd/mm/yyyy, anything else as typed."""
    if is_num(v):
        d = from_serial(v)
        return fdate(d) if d else text(v)
    return text(v)


def fdate(d):
    return d.strftime("%d/%m/%Y") if d else ""


def norm_ref(r):
    return re.sub(r"[^A-Z0-9]", "", str(r or "").upper())      # "TL- 1154497" -> "TL1154497"


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


def raw(r, i):
    return r[i] if i is not None and i < len(r) else None


def cell(r, i):
    return text(raw(r, i))


def _ours(rng):
    """Every write passes through here: only our two tabs can ever be written."""
    tab = rng.split("!")[0].strip("'")
    if tab not in OURS:
        raise RuntimeError("refusing to write to '%s' - this module only writes %s" % (tab, " and ".join(OURS)))
    return rng


# ── reading ────────────────────────────────────────────────────────────────────
def _props(svc):
    meta = svc.spreadsheets().get(spreadsheetId=ARKLE,
                                  fields="sheets(properties(title,sheetId,hidden,gridProperties(rowCount,columnCount)))").execute()
    return [sh["properties"] for sh in meta.get("sheets", [])]


def is_generated(rows):
    """True for the list this module used to build (Date Passed / Week / Agent ... From)."""
    hdr = [str(c).strip().lower() for c in (rows[0] if rows else [])]
    return "week" in hdr and "agent" in hdr and "from" in hdr


def read_tracker(svc):
    """All rows of the shared tracker tab (unformatted). NotReady while the tab is
    missing or is still the old generated list."""
    titles = {p["title"].strip().lower(): p["title"] for p in _props(svc)}
    real = titles.get(TRACKER_TAB.strip().lower())
    if not real:
        raise NotReady("no '%s' tab in the workbook yet" % TRACKER_TAB)
    rows = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range="'%s'" % real.replace("'", "''"),
                                           valueRenderOption="UNFORMATTED_VALUE",
                                           dateTimeRenderOption="SERIAL_NUMBER").execute().get("values", [])
    if is_generated(rows):
        raise NotReady("'%s' is still the old generated list (one_tab.py not run yet)" % real)
    return rows, real


def parse_tracker(rows, title):
    """Shared-tab rows -> cases. Week divider rows ('WC 05/10') and anything else
    without a date are skipped. The person comes from the Lead Gen column."""
    if not rows:
        return []
    h = find_header(rows); hdr = rows[h]
    c_date = col(hdr, "date"); c_ref = col(hdr, "crm reference", "tl ref", "reference", "ref")
    c_name = col(hdr, "client name", "client", "name"); c_agent = col(hdr, "lead gen", "agent")
    c_adv = col(hdr, "advisor", "sfm"); c_stage = col(hdr, "status", "stage")
    c_agreed = col(hdr, "agreed date", "agreed"); c_sip = col(hdr, "sip complete", "sip")
    c_moc = col(hdr, "moc set", "moc"); c_appr = col(hdr, "approval date", "approval")
    c_apprflag = col(hdr, "approved"); c_out = col(hdr, "outcome")
    missing = [n for n, c in (("Date", c_date), ("TL Ref", c_ref), ("Client name", c_name), ("Lead Gen", c_agent), ("Stage", c_stage)) if c is None]
    if missing:
        raise RuntimeError("'%s' header row has no %s column - not refreshing" % (title, " / ".join(missing)))
    out = []
    for r in rows[h + 1:]:
        d = to_date(raw(r, c_date))
        ref = cell(r, c_ref); name = cell(r, c_name)
        if not d or not (ref or name):            # week dividers, blank rows, date-only rows
            continue
        lg = " ".join(cell(r, c_agent).split())
        out.append({
            "date": d, "agent": lg.title() if lg else "Unassigned", "ref": ref, "name": name,
            "advisor": cell(r, c_adv), "stage": cell(r, c_stage), "outcome": cell(r, c_out),
            "agreed": date_text(raw(r, c_agreed)), "sip": date_text(raw(r, c_sip)), "moc": date_text(raw(r, c_moc)),
            "appr": date_text(raw(r, c_appr)), "apprflag": date_text(raw(r, c_apprflag)), "legacy": False,
        })
    return out


def load_approved(svc):
    """Arkle's Approved tab -> {norm_ref: {...}}. Dates there are often "28th April"
    with no year: rows are chronological, so the year is carried forward and bumped
    when the month goes backwards. Raises on a read error (the caller then leaves
    the snapshot alone)."""
    rows = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range=APPROVED_RNG).execute().get("values", [])
    if not rows:
        return {}
    h = find_header(rows); hdr = rows[h]
    c_ref = col(hdr, "reference", "ref"); c_name = col(hdr, "client name", "client")
    c_lg = col(hdr, "lg", "lead gen"); c_date = col(hdr, "moc approval date", "approval date", "approved", "date")
    out = {}; year = 2025; prev = None
    for r in rows[h + 1:]:
        ref = norm_ref(cell(r, c_ref))
        if not ref.startswith("TL"):
            continue
        rawd = cell(r, c_date); d = pdate(rawd)
        if d is None and rawd:
            m = re.match(r"^\s*(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)", rawd)
            if m:
                for f in ("%d %B %Y", "%d %b %Y"):
                    try:
                        d = datetime.datetime.strptime("%s %s %d" % (m.group(1), m.group(2), year), f).date()
                        break
                    except ValueError:
                        pass
                if d and prev and d < prev - datetime.timedelta(days=60):
                    year += 1; d = d.replace(year=year)
                while d and d > datetime.date.today() + datetime.timedelta(days=30):   # never in the future
                    year -= 1; d = d.replace(year=year)
        if d:
            prev = d; year = d.year
        if ref not in out:
            out[ref] = {"date": d, "name": cell(r, c_name), "agent": cell(r, c_lg)}
    return out


def flags(c):
    """Reached-at-least flags. Later stages imply the earlier ones. DMP cases keep
    'agreed' (Agreed DMP) but never count in the IVA stages."""
    t = (c["stage"] or "").lower()
    def yes(v):
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
    """Reads everything, returns (data, weeklist, agents, info). Raises if any read
    fails - the caller then writes nothing."""
    rows, title = read_tracker(svc)
    cases = parse_tracker(rows, title)
    on_tracker = len(cases)
    approved = load_approved(svc)
    seen = set()
    for c in cases:
        a = approved.get(norm_ref(c["ref"]))
        if a:
            seen.add(norm_ref(c["ref"]))
            c["apprflag"] = c["apprflag"] or "yes"
            if not c["appr"] and a["date"]:
                c["appr"] = fdate(a["date"])
    if LEGACY:
        people = {c["agent"].lower(): c["agent"] for c in cases}
        for ref, a in approved.items():
            if ref in seen or not a["date"]:
                continue
            first = (a["agent"] or "").strip().split(" ")[0].lower()
            cases.append({"date": a["date"], "agent": people.get(first, "Pre-tracker"), "ref": ref[:2] + "-" + ref[2:],
                          "name": a["name"], "advisor": "", "stage": "MOC Approved", "outcome": "", "agreed": "",
                          "sip": "", "moc": "", "appr": fdate(a["date"]), "apprflag": "yes", "legacy": True})
    cases.sort(key=lambda c: (c["date"], c["agent"], c["ref"]), reverse=NEWEST_FIRST)
    for c in cases:
        c["f"] = flags(c); c["type"] = "DMP" if c["f"]["dmp"] else "IVA"
        if c["legacy"]:                            # an approval that week, not a new pass
            c["f"].update({"agreed": False, "sip": False, "moc": False, "appr": True, "dnq": False})
    data = [DATA_HEAD]
    for c in cases:                                # ISO dates so the sheet stores real dates
        f = c["f"]
        data.append([monday(c["date"]).isoformat(), c["date"].isoformat(), c["agent"], c["type"],
                     int(f["agreed"]), int(f["sip"]), int(f["moc"]), int(f["appr"]), int(f["dnq"]), c["ref"], c["name"], int(c["legacy"])])
    weeks = sorted({monday(c["date"]) for c in cases}, reverse=True)
    weeklist = [["Week options", "Monday"], ["This week", ""], ["Last week", ""], ["All time", ""]] + \
               [["W/C " + fdate(wk), wk.isoformat()] for wk in weeks[:30]]      # dropdown: last 30 weeks
    agents = sorted({c["agent"] for c in cases})
    info = {"tab": title, "on_tracker": on_tracker, "legacy": len(cases) - on_tracker, "approved_refs": len(approved),
            "people": dict(collections.Counter(c["agent"] for c in cases if not c["legacy"]))}
    return data, weeklist, agents, info


def counts(data, start, end):
    """The snapshot's numbers for one date window, worked out here (same rules as
    the sheet's COUNTIFS) - used by the preview."""
    out = collections.OrderedDict()
    for r in data[1:]:
        d = datetime.date.fromisoformat(r[1])
        if not (start <= d <= end):
            continue
        a = out.setdefault(r[2], [0] * 8)
        a[0] += r[3] == "IVA" and not r[11]; a[1] += r[3] == "DMP" and not r[11]
        a[2] += r[3] == "IVA" and r[4]; a[3] += r[3] == "DMP" and r[4]
        a[4] += r[5]; a[5] += r[6]; a[6] += r[7]; a[7] += r[8]
    return collections.OrderedDict(sorted(out.items()))


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
            iva = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"IVA",%s!$L:$L,0' % (D, D))
            dmp = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"DMP",%s!$L:$L,0' % (D, D))
            iva_ag = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"IVA",%s!$E:$E,1' % (D, D))
            dmp_ag = '=COUNTIFS(%s)' % crit(week, a, ',%s!$D:$D,"DMP",%s!$E:$E,1' % (D, D))
            sip = '=COUNTIFS(%s)' % crit(week, a, ',%s!$F:$F,1' % D)
            moc = '=COUNTIFS(%s)' % crit(week, a, ',%s!$G:$G,1' % D)
            apr = '=COUNTIFS(%s)' % crit(week, a, ',%s!$H:$H,1' % D)
            dnq = '=COUNTIFS(%s)' % crit(week, a, ',%s!$I:$I,1' % D)
            rows.append([ag, iva, dmp, iva_ag, dmp_ag, sip, moc, apr, dnq,
                         '=IF(B%d=0,"",D%d/B%d)' % (r, r, r), '=IF(D%d=0,"",F%d/D%d)' % (r, r, r), '=IF(B%d=0,"",H%d/B%d)' % (r, r, r)])
        last = first + len(agents) - 1; t = last + 1
        tot = ["TOTAL"] + ["=SUM(%s%d:%s%d)" % (cl, first, cl, last) for cl in "BCDEFGHI"]
        tot += ['=IF(B%d=0,"",D%d/B%d)' % (t, t, t), '=IF(D%d=0,"",F%d/D%d)' % (t, t, t), '=IF(B%d=0,"",H%d/B%d)' % (t, t, t)]
        rows.append(tot)
        return rows
    start = ('=IF($B$1="This week",TODAY()-WEEKDAY(TODAY(),2)+1,IF($B$1="Last week",TODAY()-WEEKDAY(TODAY(),2)-6,'
             'IF($B$1="All time",DATE(2024,1,1),IFERROR(VLOOKUP($B$1,%s!$M:$N,2,FALSE),TODAY()-WEEKDAY(TODAY(),2)+1))))' % D)
    end = '=IF($B$1="All time",TODAY()+7,$N$1+6)'
    rows = [["WEEK", current_choice or "This week", "", '="Showing "&TEXT($N$1,"ddd dd/mm")&" to "&TEXT($N$2,"ddd dd/mm")'],
            [""] * len(SNAP_HEAD)]
    rows += table_rows(3, week=True)                      # header on row 3, people from row 4
    n = len(agents)
    rows += [[""] * len(SNAP_HEAD), ["ALL WEEKS  -  every case on the MASTER TRACKER"] + [""] * (len(SNAP_HEAD) - 1)]
    rows += table_rows(len(rows) + 1, week=False)
    # 1-based (header row, TOTAL row) of each table: week table header 3, people 4.., TOTAL n+4;
    # blank n+5, ALL WEEKS title n+6, header n+7, people.., TOTAL 2n+8
    return rows, {"N1": start, "N2": end, "week_rows": (3, n + 4), "all_rows": (n + 7, 2 * n + 8)}


# ── sheet writing (WEEKLY SNAPSHOT and MT DATA only) ───────────────────────────
def _rgb(hexs):
    return {"red": int(hexs[1:3], 16) / 255, "green": int(hexs[3:5], 16) / 255, "blue": int(hexs[5:7], 16) / 255}


def _fmt(bg, fg="#000000", bold=False):
    return {"backgroundColor": _rgb(bg), "textFormat": {"foregroundColor": _rgb(fg), "bold": bold}}


def _cell(gid, r0, r1, c0, c1, fmt, fields):
    return {"repeatCell": {"range": {"sheetId": gid, "startRowIndex": r0, "endRowIndex": r1, "startColumnIndex": c0, "endColumnIndex": c1},
                           "cell": {"userEnteredFormat": fmt}, "fields": fields}}


NAVY, NAVY_LIGHT, ROW_ALT, YELLOW_CELL, TEAL = "#1F4E79", "#D9E2F3", "#F3F7FB", "#FFF2CC", "#0B6E4F"


def ensure_tabs(svc, need_rows=0):
    """Our two tabs exist and are big enough (state lives in MT DATA!P:Q)."""
    props = {p["title"]: p for p in _props(svc)}
    reqs = []
    for title in OURS:
        p = props.get(title)
        if p is None:
            reqs.append({"addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": 2000, "columnCount": 20}}}})
            continue
        g = p.get("gridProperties", {})
        if g.get("columnCount", 0) < 18:
            reqs.append({"appendDimension": {"sheetId": p["sheetId"], "dimension": "COLUMNS", "length": 20 - g.get("columnCount", 0)}})
        if title == DATA_TAB and g.get("rowCount", 0) < need_rows + 50:
            reqs.append({"appendDimension": {"sheetId": p["sheetId"], "dimension": "ROWS", "length": need_rows + 500 - g.get("rowCount", 0)}})
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": reqs}).execute()
        log.info("master-tracker: prepared tabs (%d change%s)" % (len(reqs), "" if len(reqs) == 1 else "s"))


def read_state(svc):
    """{'updated','hash','format','code'} from MT DATA!P1:Q4 ({} if not there yet)."""
    try:
        vals = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range=STATE_RNG).execute().get("values", [])
    except Exception:
        return {}
    return {str(r[0]).strip(): str(r[1]).strip() for r in vals if len(r) > 1}


def write_state(svc, **kv):
    st = read_state(svc); st.update(kv)
    svc.spreadsheets().values().update(spreadsheetId=ARKLE, range=_ours(STATE_RNG), valueInputOption="RAW", body={"values": [
        ["updated", st.get("updated", "")], ["hash", st.get("hash", "")], ["format", st.get("format", "")], ["code", st.get("code", "")]]}).execute()


def ensure_formatting(svc):
    """Versioned one-off for the snapshot tab: dropdown, widths, gridlines; hides
    MT DATA. Never runs on an ordinary refresh and never touches the tracker tab."""
    if read_state(svc).get("format") == FORMAT_VERSION:
        return
    meta = svc.spreadsheets().get(spreadsheetId=ARKLE, fields="sheets(properties(title,sheetId),conditionalFormats)").execute()
    info = {sh["properties"]["title"]: (sh["properties"]["sheetId"], len(sh.get("conditionalFormats", []))) for sh in meta["sheets"]}
    sg, n = info[SNAP_TAB]; dg, _ = info[DATA_TAB]
    reqs = [{"deleteConditionalFormatRule": {"sheetId": sg, "index": 0}} for _ in range(n)]
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
    reqs.append({"updateSheetProperties": {"properties": {"sheetId": dg, "hidden": True}, "fields": "hidden"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": reqs}).execute()
    write_state(svc, format=FORMAT_VERSION)
    log.info("master-tracker: snapshot layout %s applied (%d requests)" % (FORMAT_VERSION, len(reqs)))


def snapshot_format_requests(gid, layout, nrows):
    """Cell colours for the snapshot's fixed layout - a handful of repeatCell
    requests, re-applied only when the tab is rewritten."""
    F = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)"
    reqs = [_cell(gid, 0, nrows + 2, 0, 14, {"backgroundColor": _rgb("#FFFFFF"), "textFormat": {"bold": False, "foregroundColor": _rgb("#000000")}}, "userEnteredFormat(backgroundColor,textFormat)")]
    reqs.append(_cell(gid, 0, 1, 0, 1, _fmt("#FFFFFF", NAVY, True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append(_cell(gid, 0, 1, 1, 2, _fmt(YELLOW_CELL, "#000000", True), "userEnteredFormat(backgroundColor,textFormat)"))
    reqs.append(_cell(gid, 0, 1, 3, 4, {"textFormat": {"italic": True, "foregroundColor": _rgb("#555555")}}, "userEnteredFormat.textFormat"))
    reqs.append(_cell(gid, 0, 1, 9, 12, {"textFormat": {"italic": True, "foregroundColor": _rgb("#999999")}}, "userEnteredFormat.textFormat"))   # "Updated ..."
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


def write(svc, data, weeklist, agents, h):
    """Rewrites MT DATA and WEEKLY SNAPSHOT. The hash is stored last, so a write
    that dies half way is simply done again on the next cycle."""
    if len(data) > MAX_ROWS:
        raise RuntimeError("refusing to write %d rows (cap %d)" % (len(data), MAX_ROWS))
    ensure_tabs(svc, need_rows=len(data))
    ensure_formatting(svc)
    gids = {p["title"]: p["sheetId"] for p in _props(svc)}
    try:   # keep whatever week is selected in the dropdown
        cur = svc.spreadsheets().values().get(spreadsheetId=ARKLE, range="'%s'!B1" % SNAP_TAB).execute().get("values", [[""]])
        choice = (cur[0][0] if cur and cur[0] else "") or "This week"
    except Exception:
        choice = "This week"
    snap, layout = snapshot_rows(agents, choice)
    stamp = _now().strftime("%d/%m %H:%M")
    svc.spreadsheets().values().batchClear(spreadsheetId=ARKLE, body={"ranges": [
        _ours("'%s'!A1:L200" % SNAP_TAB), _ours("'%s'!A:N" % DATA_TAB)]}).execute()
    svc.spreadsheets().values().batchUpdate(spreadsheetId=ARKLE, body={"valueInputOption": "USER_ENTERED", "data": [
        {"range": _ours("'%s'!A1" % DATA_TAB), "values": data},
        {"range": _ours("'%s'!M1" % DATA_TAB), "values": weeklist},
        {"range": _ours("'%s'!A1" % SNAP_TAB), "values": snap},
        {"range": _ours("'%s'!J1" % SNAP_TAB), "values": [["Updated " + stamp]]},
        {"range": _ours("'%s'!M1:N2" % SNAP_TAB), "values": [["start", layout["N1"]], ["end", layout["N2"]]]},
    ]}).execute()
    svc.spreadsheets().batchUpdate(spreadsheetId=ARKLE, body={"requests": snapshot_format_requests(gids[SNAP_TAB], layout, len(snap))}).execute()
    write_state(svc, updated=_now().strftime("%d/%m/%Y %H:%M"), hash=h, code=CODE_VERSION)   # last: marks the write complete


def _hash(data, weeklist, agents):
    return hashlib.md5(repr([data, weeklist, agents, FORMAT_VERSION]).encode()).hexdigest()


def run(svc, dry_run=True, poller=False):
    global _last_hash, _stamped
    if _last_hash is None:
        _last_hash = read_state(svc).get("hash", "")
    if poller and not dry_run and not _stamped:
        # tells one_tab.py that the live poller is this build (it must be, before the tab is renamed)
        if read_state(svc).get("code") != CODE_VERSION:
            ensure_tabs(svc)
            write_state(svc, code=CODE_VERSION)
            log.info("master-tracker: build %s live" % CODE_VERSION)
        _stamped = True
    data, weeklist, agents, info = build(svc)            # raises -> nothing is written
    n = len(data) - 1
    desc = "%d on %s + %d Arkle-approved not on it" % (info["on_tracker"], info["tab"], info["legacy"])
    h = _hash(data, weeklist, agents)
    if h == _last_hash:
        log.info("master-tracker: no change (%s)" % desc)
        return False
    if dry_run:
        log.info("master-tracker [DRY RUN]: would write %d rows (%s)" % (n, desc))
        return False
    write(svc, data, weeklist, agents, h)
    _last_hash = h
    log.info("master-tracker: snapshot written - %s | %s" % (desc, ", ".join("%s %d" % kv for kv in sorted(info["people"].items()))))
    return True


def tick(get_svc):
    """Called every poller cycle; runs at most every INTERVAL_S (RETRY_S after a
    cycle that could not read or is still waiting for the tab)."""
    global _last_tick
    if not ENABLED or time.time() - _last_tick < INTERVAL_S:
        return
    _last_tick = time.time()
    try:
        run(get_svc(), dry_run=DRY_RUN, poller=True)
    except NotReady as e:
        _last_tick = time.time() - INTERVAL_S + RETRY_S
        log.info("master-tracker: waiting - %s. Nothing written." % e)
    except Exception as e:
        _last_tick = time.time() - INTERVAL_S + RETRY_S
        log.warning("master-tracker: not refreshed this cycle, snapshot left as it was (retry in %ds): %s" % (RETRY_S, e))


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    sys.path.insert(0, ".")
    import main
    svc = main.get_sheets_service()
    try:
        data, weeklist, agents, info = build(svc)
    except NotReady as e:
        print("waiting - %s. Nothing written." % e); sys.exit()
    print("source tab: %s | cases on it: %d | Arkle-approved not on it: %d | Approved tab refs: %d" % (
        info["tab"], info["on_tracker"], info["legacy"], info["approved_refs"]))
    print("people (cases on the tab):", ", ".join("%s %d" % kv for kv in sorted(info["people"].items())))
    today = _now().date(); wk = monday(today)
    cols = ["IVA", "DMP", "IVAagr", "DMPagr", "SIP", "MOC", "Appr", "DNQ"]
    for label, a, b in (("THIS WEEK", wk, wk + datetime.timedelta(days=6)), ("LAST WEEK", wk - datetime.timedelta(days=7), wk - datetime.timedelta(days=1)),
                        ("ALL TIME", datetime.date(2024, 1, 1), today + datetime.timedelta(days=7))):
        print("%s  (%s - %s)" % (label, a.strftime("%d/%m"), b.strftime("%d/%m")))
        print("  %-12s" % "" + "".join("%7s" % c for c in cols))
        for ag, v in counts(data, a, b).items():
            print("  %-12s" % ag + "".join("%7d" % x for x in v))
    if os.getenv("GO") != "1":
        print("preview only - GO=1 writes %s and %s (never %s)" % (SNAP_TAB, DATA_TAB, TRACKER_TAB)); sys.exit()
    write(svc, data, weeklist, agents, _hash(data, weeklist, agents))
    print("written:", SNAP_TAB, "and", DATA_TAB)
