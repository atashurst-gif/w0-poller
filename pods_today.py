#!/usr/bin/env python3
"""PODs TODAY for the Arkle workbook - the same idea as CB'S TODAY for callbacks.

Reads the 'PODs' tab (one row per POD booked) and fills the two time grids on
the 'PODs TODAY' tab:
  TODAY     PODs whose date (PODs column G) is today
  TOMORROW  PODs dated the next working day (on a Friday that is Monday)

Every cycle, any POD not yet on its grid is put on the row for its time. If that
row is already taken, or is the lunch block, a new row with the same time is
inserted next to it, so nothing anyone has typed is ever overwritten.

When the date beside TODAY or TOMORROW is no longer right (the first cycle of a
new day), that grid is reset: extra rows are removed, entries and highlight
colours are cleared, the date is updated, and the grid is filled again from PODs.

Writes only to 'PODs TODAY'. Never writes to 'PODs'.

Columns copied (PODs -> PODs TODAY):
  TL Reference B->C, Name C->D, Number D->E, Creditor Name I->F,
  creditor Number J->G, Answered K->H, Notes L->I;  Time H -> the time row (B)

Standalone (from ~/Desktop/w0-poller):
  python3 pods_today.py          # preview: what it would do, writes nothing
  GO=1 python3 pods_today.py     # do it now
"""
import os
import sys
import time
import datetime
import logging
from zoneinfo import ZoneInfo

log = logging.getLogger(__name__)

SHEET_ID   = os.getenv("PODS_SHEET_ID", "16qnJ842lFAo-4FVhloipJMFdCs7Spc2hFgoWnVGwVUs")
SRC_TAB    = "PODs"
DST_TAB    = "PODs TODAY"
ENABLED    = os.getenv("PODS_TODAY_ENABLED", "1") == "1"
DRY_RUN    = os.getenv("PODS_TODAY_DRY_RUN", "0") == "1"
INTERVAL_S = int(os.getenv("PODS_TODAY_INTERVAL_S", "60"))
UK         = ZoneInfo("Europe/London")
STANDARD   = {"%02d:%02d" % (h, m) for h in range(9, 17) for m in (0, 30)}   # 09:00 .. 16:30
WHITE      = {"red": 1, "green": 1, "blue": 1}
SECTIONS   = ("TODAY", "TOMORROW")
_last_tick = 0.0
_last_warn = 0.0


# ── small helpers ─────────────────────────────────────────────────────────────
def _now():
    return datetime.datetime.now(UK)


def norm_time(v):
    """'9:00', '09:00', '9.30', '9:00 AM', '09:00:00', 0.375 -> 'HH:MM'; None if not a time."""
    s = str(v if v is not None else "").strip().upper().replace(".", ":")
    if not s:
        return None
    ampm = ""
    if s.endswith("AM") or s.endswith("PM"):
        ampm, s = s[-2:], s[:-2].strip()
    if ":" not in s:
        if s.isdigit() and ampm:
            s += ":00"
        else:
            try:
                f = float(s)
            except ValueError:
                return None
            if not 0 <= f < 1:
                return None
            mins = int(round(f * 1440))
            return "%02d:%02d" % (mins // 60, mins % 60)
    parts = s.split(":")
    try:
        hh, mm = int(parts[0]), int(parts[1][:2])
    except (ValueError, IndexError):
        return None
    if ampm == "PM" and hh != 12:
        hh += 12
    if ampm == "AM" and hh == 12:
        hh = 0
    if not (0 <= hh < 24 and 0 <= mm < 60):
        return None
    return "%02d:%02d" % (hh, mm)


def norm_date(v):
    """'16/10/2026', '16/10/26', '16-10-2026', '2026-10-16' or a sheet serial -> date; None otherwise."""
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    s = str(v if v is not None else "").strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def serial_date(d):
    return (d - datetime.date(1899, 12, 30)).days


def serial_time(t):
    h, m = map(int, t.split(":"))
    return (h * 60 + m) / 1440.0


def _d(d):
    return d.strftime("%d/%m/%Y") if d else "no date"


def next_working_day(d):
    d = d + datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d += datetime.timedelta(days=1)
    return d


def pod_key(number, tl, name, t):
    digits = "".join(c for c in str(number) if c.isdigit())[-9:]
    who = digits or str(tl).strip().upper() or str(name).strip().lower()
    return (who, t)


# ── reading ───────────────────────────────────────────────────────────────────
def read(svc):
    res = svc.spreadsheets().values().batchGet(
        spreadsheetId=SHEET_ID,
        ranges=["'%s'!A:L" % SRC_TAB, "'%s'!A1:I500" % DST_TAB]).execute()
    vr = res.get("valueRanges", [])
    return vr[0].get("values", []), vr[1].get("values", [])


def pods_by_date(src):
    """{date: [pod, ...]} from the PODs tab. Rows without a usable date/time or person are skipped."""
    out, skipped = {}, 0
    for r in src[1:]:
        r = (list(r) + [""] * 12)[:12]
        if not any(str(x).strip() for x in r):
            continue
        d, t = norm_date(r[6]), norm_time(r[7])
        if not d or not t or not (str(r[3]).strip() or str(r[2]).strip()):
            skipped += 1
            continue
        vals = [str(x).strip() for x in (r[1], r[2], r[3], r[8], r[9], r[10], r[11])]
        out.setdefault(d, []).append({"t": t, "vals": vals, "key": pod_key(r[3], r[1], r[2], t)})
    for d in out:
        out[d].sort(key=lambda p: p["t"])
    return out, skipped


def sections(grid):
    """Find the TODAY / TOMORROW grids: label row, the date beside it, and the time rows under the header."""
    secs = {}
    for i, row in enumerate(grid):
        label = str(row[1]).strip().upper() if len(row) > 1 else ""
        if label not in SECTIONS or label in secs:
            continue
        header = grid[i + 1] if i + 1 < len(grid) else []
        if len(header) < 3 or "TL" not in str(header[2]).upper():
            continue
        rows, j = [], i + 2
        while j < len(grid):
            rr = (list(grid[j]) + [""] * 9)[:9]
            t = norm_time(rr[1])
            if not t:
                break
            rows.append({"r": j + 1, "t": t, "cells": [str(x).strip() for x in rr[2:9]]})
            j += 1
        secs[label] = {"label_row": i + 1, "date": norm_date(row[3] if len(row) > 3 else ""), "rows": rows}
    return secs


def meta(svc):
    """sheetId of PODs TODAY and which rows are merged across the entry columns (the lunch block)."""
    m = svc.spreadsheets().get(spreadsheetId=SHEET_ID,
                               fields="sheets(properties(sheetId,title),merges)").execute()
    for sh in m.get("sheets", []):
        if sh["properties"]["title"] == DST_TAB:
            merged = {}
            for mg in sh.get("merges", []) or []:
                if mg.get("startColumnIndex", 0) <= 8 and mg.get("endColumnIndex", 0) >= 3:
                    first, last = mg["startRowIndex"] + 1, mg["endRowIndex"]
                    for r in range(first, last + 1):
                        merged[r] = (first, last)
            return {"gid": sh["properties"]["sheetId"], "merged": merged}
    raise RuntimeError("tab %r not found" % DST_TAB)


# ── writing (every write is to PODs TODAY only) ──────────────────────────────
class Writer:
    def __init__(self, svc, gid, dry):
        self.svc, self.gid, self.dry = svc, gid, dry

    def _rng(self, a1):
        return "'%s'!%s" % (DST_TAB, a1)

    def values(self, a1, rows):
        if self.dry:
            return
        self.svc.spreadsheets().values().batchUpdate(spreadsheetId=SHEET_ID, body={
            "valueInputOption": "RAW", "data": [{"range": self._rng(a1), "values": rows}]}).execute()

    def clear(self, a1_list):
        if self.dry or not a1_list:
            return
        self.svc.spreadsheets().values().batchClear(spreadsheetId=SHEET_ID, body={
            "ranges": [self._rng(a) for a in a1_list]}).execute()

    def structure(self, reqs):
        if self.dry or not reqs:
            return
        for q in reqs:   # only ever this tab's rows
            for v in q.values():
                rng = v.get("range", {})
                assert rng.get("sheetId") == self.gid, "pods-today: refusing to touch another tab"
        self.svc.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={"requests": reqs}).execute()

    def white(self, r):
        return {"repeatCell": {"range": {"sheetId": self.gid, "startRowIndex": r - 1, "endRowIndex": r,
                                         "startColumnIndex": 2, "endColumnIndex": 9},
                               "cell": {"userEnteredFormat": {"backgroundColor": WHITE}},
                               "fields": "userEnteredFormat.backgroundColor"}}


def _extras(sec, merged):
    """Rows added on top of the 09:00-16:30 template: off-template times, or a second row for a time."""
    groups = {}
    for row in sec["rows"]:
        groups.setdefault(row["t"], []).append(row)
    out = []
    for t, rows in groups.items():
        if t not in STANDARD:
            out += [x["r"] for x in rows if x["r"] not in merged]
            continue
        keep = next((x for x in rows if x["r"] in merged), rows[0])
        out += [x["r"] for x in rows if x is not keep and x["r"] not in merged]
    return sorted(out)


def reset(w, sec, merged, name, target):
    """Bring a grid back to the empty template for a new date. Returns a description of what was done."""
    extras = _extras(sec, merged)
    if extras:
        w.structure([{"deleteDimension": {"range": {"sheetId": w.gid, "dimension": "ROWS",
                                                     "startIndex": r - 1, "endIndex": r}}}
                     for r in sorted(extras, reverse=True)])
        return "%s: removed %d extra row(s) from %s" % (name, len(extras), _d(sec["date"]))
    plain = [row for row in sec["rows"] if row["r"] not in merged]
    w.clear(["C%d:I%d" % (row["r"], row["r"]) for row in plain if any(row["cells"])])
    w.structure([w.white(row["r"]) for row in plain])
    w.values("D%d" % sec["label_row"], [[serial_date(target)]])
    return "%s: cleared %d entr(ies) from %s, date set to %s" % (
        name, sum(1 for row in plain if any(row["cells"])), _d(sec["date"]), target.strftime("%d/%m/%Y"))


def place(w, sec, merged, pod):
    """Put one POD on the grid. Returns (description, structural_change)."""
    for row in sec["rows"]:
        if row["t"] == pod["t"] and row["r"] not in merged and not any(row["cells"]):
            w.values("C%d:I%d" % (row["r"], row["r"]), [pod["vals"]])
            row["cells"] = list(pod["vals"])
            return "row %d" % row["r"], False
    after = None
    for row in sec["rows"]:
        if row["t"] <= pod["t"]:
            after = row
    if after is None:                                    # earlier than the first time on the grid
        at, inherit = sec["rows"][0]["r"], False
    elif after["r"] in merged:                           # never insert inside the lunch block
        first, last = merged[after["r"]]
        last_t = next((x["t"] for x in sec["rows"] if x["r"] == last), pod["t"])
        at, inherit = (first, True) if pod["t"] < last_t else (last + 1, True)
    else:
        at, inherit = after["r"] + 1, True
    w.structure([{"insertDimension": {"range": {"sheetId": w.gid, "dimension": "ROWS",
                                                "startIndex": at - 1, "endIndex": at},
                                      "inheritFromBefore": inherit}},
                 w.white(at)])
    w.values("B%d:I%d" % (at, at), [[serial_time(pod["t"])] + pod["vals"]])
    return "new row %d" % at, True


def run(svc, dry_run=False, now=None):
    """One pass. Returns the list of things done (or that would be done in a dry run)."""
    now = now or _now()
    today = now.date()
    targets = {"TODAY": today, "TOMORROW": next_working_day(today)}
    done = []
    for _ in range(80):                                  # re-read after every structural change
        src, grid = read(svc)
        secs = sections(grid)
        if any(n not in secs or not secs[n]["rows"] for n in SECTIONS):
            raise RuntimeError("could not find the TODAY and TOMORROW grids on %s - nothing written" % DST_TAB)
        pods, _ = pods_by_date(src)
        info, w, again = None, None, False
        for name in SECTIONS:
            sec, target = secs[name], targets[name]
            if sec["date"] != target:
                info = info or meta(svc)
                w = w or Writer(svc, info["gid"], dry_run)
                done.append("pods-today: " + reset(w, sec, info["merged"], name, target))
                again = True
                break
            present = {pod_key(r["cells"][2], r["cells"][0], r["cells"][1], r["t"])
                       for r in sec["rows"] if any(r["cells"])}
            seen, missing = set(), []
            for p in pods.get(target, []):
                if p["key"] not in present and p["key"] not in seen:
                    seen.add(p["key"]); missing.append(p)
            for p in missing:
                info = info or meta(svc)
                w = w or Writer(svc, info["gid"], dry_run)
                where, structural = place(w, sec, info["merged"], p)
                done.append("pods-today: %s %s | %s | %s -> %s (%s)" % (
                    name, p["t"], p["vals"][1][:24], p["vals"][2], where, target.strftime("%d/%m/%Y")))
                if structural:
                    again = True
                    break
            if again:
                break
        if not again or dry_run:
            return done
    log.warning("pods-today: stopped after 80 passes")
    return done


def tick(get_svc):
    """Called every poller cycle; runs at most every INTERVAL_S."""
    global _last_tick, _last_warn
    if not ENABLED or time.time() - _last_tick < INTERVAL_S:
        return
    _last_tick = time.time()
    try:
        for line in run(get_svc(), dry_run=DRY_RUN):
            log.info(("[DRY RUN] " if DRY_RUN else "") + line)
    except Exception as e:
        if time.time() - _last_warn > 3600:
            _last_warn = time.time()
            log.warning("pods-today: not updated this cycle: %s" % e)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    sys.path.insert(0, ".")
    import main
    go = os.getenv("GO") == "1"
    lines = run(main.get_sheets_service(), dry_run=not go)
    print("\n".join(lines) if lines else "pods-today: nothing to do - grids are up to date")
    if not go and lines:
        print("(preview only - run with GO=1 to write)")
