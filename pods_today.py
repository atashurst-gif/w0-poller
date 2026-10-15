#!/usr/bin/env python3
"""PODs TODAY for the Arkle workbook - the same idea as CB'S TODAY for callbacks.

1. The day grids
   Reads the 'PODs' tab (one row per POD booked) and fills the two time grids on
   the 'PODs TODAY' tab:
     TODAY     PODs dated today
     TOMORROW  PODs dated the next working day (on a Friday that is Monday)
   Every cycle, any POD not yet on its grid is put on the row for its time. If that
   row is already taken, or is the lunch block, a new row with the same time is
   inserted next to it, so nothing anyone has typed is ever overwritten.
   When the date beside TODAY or TOMORROW is no longer right (the first cycle of a
   new day), that grid is reset: extra rows removed, entries and highlight colours
   cleared, date updated, then filled again from PODs.

2. Answered + missed POD message (TODAY grid only)
   Answered (Yes/No) lives only on the TODAY grid (column H). When a POD there is
   marked No, the client gets the WATI template pod_missed_call1 (Regen's WATI)
   once. The PODs row is stamped in its "Missed Call 1" column
   ("1 - dd/mm/yyyy hh:mm", or "Failed - ..." if WATI refuses it) so it is never
   sent twice. Nothing is sent until WATI shows the template as APPROVED.

PODs columns are found by their headings (TL Reference, Name, the first Number,
the Date next to Time, Time, Creditor Name, the Number after it, Notes), so adding
or removing a column there does not break anything. The only write to PODs is
the "Missed Call 1" column (added at the end if it is not there).

Grid columns: B time, C TL Reference, D Name, E Number, F Creditor Name,
G creditor Number, H Answered, I Notes.

Standalone (from ~/Desktop/w0-poller):
  python3 pods_today.py          # preview: what it would do, writes and sends nothing
  GO=1 python3 pods_today.py     # do it now
"""
import os
import re
import sys
import time
import datetime
import logging
from zoneinfo import ZoneInfo

import requests

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

MISSED_ENABLED  = os.getenv("POD_MISSED_ENABLED", "1") == "1"
MISSED_DRY_RUN  = os.getenv("POD_MISSED_DRY_RUN", "0") == "1"
MISSED_TEMPLATE = os.getenv("POD_MISSED_TEMPLATE", "pod_missed_call1")
MISSED_DAYS     = int(os.getenv("POD_MISSED_DAYS", "7"))
MISSED_HEADER   = "Missed Call 1"
STAMP_FMT       = "%d/%m/%Y %H:%M"
WATI_API_URL    = os.getenv("WATI_API_URL", "https://eu-api.wati.io/602557")
WATI_TOKEN      = os.getenv("WATI_TOKEN", "")

# Same Yes/No chips as the Answered column on PODs and Apps 2.0
YES_NO_RULE = {"condition": {"type": "ONE_OF_LIST", "values": [
    {"userEnteredValue": "Yes", "dropdownStyle": {
        "backgroundColorStyle": {"rgbColor": {"red": 0.83137256, "green": 0.92941177, "blue": 0.7372549}},
        "foregroundColorStyle": {"rgbColor": {"red": 0.06666667, "green": 0.4509804, "blue": 0.29411766}}}},
    {"userEnteredValue": "No", "dropdownStyle": {
        "backgroundColorStyle": {"rgbColor": {"red": 1, "green": 0.8117647, "blue": 0.7882353}},
        "foregroundColorStyle": {"rgbColor": {"red": 0.69411767, "green": 0.007843138, "blue": 0.007843138}}}}]},
    "showCustomUi": True, "strict": True}

_last_tick = 0.0
_warned = {}
_dropdowns_done = False
_template_cache = {"at": 0.0, "info": None}


def _warn_hourly(key, msg):
    if time.time() - _warned.get(key, 0) > 3600:
        _warned[key] = time.time()
        log.warning(msg)


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


def norm_answer(v):
    s = str(v or "").strip().lower()
    return "Yes" if s == "yes" else "No" if s == "no" else ""


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


def wa_number(raw):
    """Same rules as the poller's format_phone: UK mobile -> 447xxxxxxxxx."""
    s = str(raw or "").strip()
    digits = re.sub(r"\D", "", s)
    if "e" in s.lower() or s.endswith(".0"):
        try:
            digits = re.sub(r"\D", "", str(int(float(s))))
        except ValueError:
            pass
    if digits.startswith("07") and len(digits) == 11:
        return "44" + digits[1:]
    if digits.startswith("440") and len(digits) == 13:
        return "44" + digits[3:]
    if digits.startswith("7") and len(digits) == 10:
        return "44" + digits
    return digits


def _col(idx):
    letters, idx = "", idx + 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ── reading ───────────────────────────────────────────────────────────────────
def read(svc):
    res = svc.spreadsheets().values().batchGet(
        spreadsheetId=SHEET_ID,
        ranges=["'%s'!A:Z" % SRC_TAB, "'%s'!A1:I500" % DST_TAB]).execute()
    vr = res.get("valueRanges", [])
    return vr[0].get("values", []), vr[1].get("values", [])


def src_columns(header):
    """Find the PODs columns by their headings. Raises if the layout is not recognised."""
    h = [str(x).strip().lower() for x in header]
    def first(name, start=0):
        for i in range(start, len(h)):
            if h[i] == name:
                return i
        return None
    c = {"tl": first("tl reference"), "name": first("name"), "phone": first("number"),
         "time": first("time"), "creditor": first("creditor name"), "notes": first("notes"),
         "missed": first(MISSED_HEADER.lower())}
    if c["time"] is not None:
        c["date"] = next((i for i in range(c["time"] - 1, -1, -1) if h[i] == "date"), None)
    if c["creditor"] is not None:
        c["cred_phone"] = first("number", c["creditor"] + 1)
    need = ("tl", "name", "phone", "time", "date", "creditor")
    missing = [k for k in need if c.get(k) is None]
    if missing:
        raise RuntimeError("PODs headings not recognised (missing %s) - nothing written" % ", ".join(missing))
    if c["missed"] is None:
        c["missed"] = max(i for i, x in enumerate(h) if x) + 1
        c["missed_new"] = True
    return c


def pods_by_date(src):
    """({date: [pod, ...]}, columns) from the PODs tab. Rows without a usable date/time or person are skipped."""
    cols = src_columns(src[0] if src else [])
    g = lambda r, k: str(r[cols[k]]).strip() if cols.get(k) is not None and cols[k] < len(r) else ""
    out = {}
    for i, r in enumerate(src[1:], start=2):
        if not any(str(x).strip() for x in r):
            continue
        d, t = norm_date(g(r, "date")), norm_time(g(r, "time"))
        if not d or not t or not (g(r, "phone") or g(r, "name")):
            continue
        vals = [g(r, "tl"), g(r, "name"), g(r, "phone"), g(r, "creditor"), g(r, "cred_phone"), "", g(r, "notes")]
        out.setdefault(d, []).append({
            "t": t, "d": d, "row": i, "vals": vals, "key": pod_key(g(r, "phone"), g(r, "tl"), g(r, "name"), t),
            "name": g(r, "name"), "phone": g(r, "phone"), "creditor": g(r, "creditor"),
            "missed": "" if cols.get("missed_new") else g(r, "missed")})
    for d in out:
        out[d].sort(key=lambda p: p["t"])
    return out, cols


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
    """sheetIds, the PODs column count, and which grid rows are merged (the lunch block)."""
    m = svc.spreadsheets().get(
        spreadsheetId=SHEET_ID,
        fields="sheets(properties(sheetId,title,gridProperties(columnCount)),merges)").execute()
    out = {}
    for sh in m.get("sheets", []):
        p = sh["properties"]
        if p["title"] == DST_TAB:
            merged = {}
            for mg in sh.get("merges", []) or []:
                if mg.get("startColumnIndex", 0) <= 8 and mg.get("endColumnIndex", 0) >= 3:
                    first, last = mg["startRowIndex"] + 1, mg["endRowIndex"]
                    for r in range(first, last + 1):
                        merged[r] = (first, last)
            out.update(gid=p["sheetId"], merged=merged)
        elif p["title"] == SRC_TAB:
            out.update(src_gid=p["sheetId"], src_cols=p.get("gridProperties", {}).get("columnCount", 0))
    if "gid" not in out:
        raise RuntimeError("tab %r not found" % DST_TAB)
    return out


# ── writing ───────────────────────────────────────────────────────────────────
class Writer:
    """Grid writes go to PODs TODAY only."""
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

    def dropdown(self, r):
        """Yes/No chips on the grid's Answered cell (H)."""
        return {"setDataValidation": {"range": {"sheetId": self.gid, "startRowIndex": r - 1, "endRowIndex": r,
                                                "startColumnIndex": 7, "endColumnIndex": 8},
                                      "rule": YES_NO_RULE}}


def src_values(svc, a1, rows, dry, cols):
    """The only write to the PODs tab: the Missed Call 1 column."""
    assert re.match(r"^%s\d+$" % _col(cols["missed"]), a1), "pods-today: refusing to write %s on PODs" % a1
    if dry:
        return
    svc.spreadsheets().values().batchUpdate(spreadsheetId=SHEET_ID, body={
        "valueInputOption": "RAW", "data": [{"range": "'%s'!%s" % (SRC_TAB, a1), "values": rows}]}).execute()


# ── grids ─────────────────────────────────────────────────────────────────────
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
    w.structure([w.white(row["r"]) for row in plain] +
                ([w.dropdown(row["r"]) for row in plain] if name == "TODAY" else []))
    w.values("D%d" % sec["label_row"], [[serial_date(target)]])
    return "%s: cleared %d entr(ies) from %s, date set to %s" % (
        name, sum(1 for row in plain if any(row["cells"])), _d(sec["date"]), target.strftime("%d/%m/%Y"))


def place(w, sec, merged, pod, name="TODAY"):
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
                 w.white(at)] + ([w.dropdown(at)] if name == "TODAY" else []))
    w.values("B%d:I%d" % (at, at), [[serial_time(pod["t"])] + pod["vals"]])
    return "new row %d" % at, True


# ── missed POD message ────────────────────────────────────────────────────────
def template_info():
    """(status, [param names]) for MISSED_TEMPLATE from WATI, cached 10 minutes; (None, []) if not found."""
    if time.time() - _template_cache["at"] < 600 and _template_cache["info"] is not None:
        return _template_cache["info"]
    info, page = (None, []), 1
    while page <= 10:
        r = requests.get("%s/api/v1/getMessageTemplates?pageSize=100&pageNumber=%d" % (WATI_API_URL, page),
                         headers={"Authorization": "Bearer " + WATI_TOKEN}, timeout=30)
        r.raise_for_status()
        j = r.json()
        items = j.get("messageTemplates") or j.get("templates") or []
        for t in items:
            if str(t.get("elementName") or t.get("name") or "") == MISSED_TEMPLATE:
                params = [str(c.get("paramName") or c.get("name")) for c in (t.get("customParams") or [])]
                if not params:
                    params = re.findall(r"\{\{\s*(\w+)\s*\}\}", str(t.get("bodyOriginal") or t.get("body") or ""))
                info = (str(t.get("status") or "").upper(), list(dict.fromkeys(params)))
                break
        if info[0] is not None or len(items) < 100:
            break
        page += 1
    _template_cache.update(at=time.time(), info=info)
    return info


def _param_value(name, pod):
    n = name.lower()
    first = pod["name"].split()[0].title() if pod["name"] else "there"
    if n in ("name", "first_name", "firstname", "first", "client", "client_name"):
        return first
    if n in ("full_name", "fullname"):
        return pod["name"] or first
    if n in ("time", "pod_time", "call_time", "appointment_time"):
        return pod["t"]
    if n in ("date", "pod_date", "appointment_date"):
        return pod["d"].strftime("%d/%m/%Y")
    if n in ("creditor", "creditor_name"):
        return pod["creditor"] or "your creditor"
    return None


def send_missed(pod, params):
    """True sent, False refused by WATI, None could not reach WATI (try again next cycle)."""
    number = wa_number(pod["phone"])
    custom = []
    for p in params:
        v = _param_value(p, pod)
        if v is None:
            raise RuntimeError("template %s has a parameter {{%s}} this code does not know how to fill" % (MISSED_TEMPLATE, p))
        custom.append({"name": p, "value": v})
    payload = {"template_name": MISSED_TEMPLATE,
               "broadcast_name": "pod_missed_" + number[-4:],
               "receivers": [{"whatsappNumber": number, "customParams": custom}]}
    try:
        r = requests.post(WATI_API_URL + "/api/v2/sendTemplateMessages", json=payload, timeout=30,
                          headers={"Authorization": "Bearer " + WATI_TOKEN, "Content-Type": "application/json"})
    except Exception as e:
        log.warning("pod-missed: could not reach WATI for %s: %s" % (number, e))
        return None
    try:
        body = r.json()
    except Exception:
        body = {}
    if r.status_code in (200, 201) and body.get("result") is True:
        return True
    log.error("pod-missed: WATI refused %s for %s: %s %s" % (MISSED_TEMPLATE, number, r.status_code, str(body or r.text)[:200]))
    return False


def chase_missed(svc, info, secs, pods, cols, now, dry):
    """Send the missed-POD message for every POD marked No on the TODAY grid that has not had it."""
    if not MISSED_ENABLED:
        return []
    today = now.date()
    by_key = {}
    for p in pods.get(today, []):
        by_key.setdefault(p["key"], []).append(p)
    due, unmatched = [], []
    for row in secs["TODAY"]["rows"]:
        if not any(row["cells"]) or norm_answer(row["cells"][5]) != "No":
            continue
        k = pod_key(row["cells"][2], row["cells"][0], row["cells"][1], row["t"])
        ps = by_key.get(k)
        if not ps:
            unmatched.append("%s %s" % (row["t"], row["cells"][1][:24]))
            continue
        if not any(p["missed"] for p in ps):
            due.append(ps[0])
    if unmatched:
        _warn_hourly("unmatched", "pod-missed: marked No on the TODAY grid but not on PODs for today, so not sent: %s" % "; ".join(unmatched))
    if not due:
        return []
    if not WATI_TOKEN:
        _warn_hourly("token", "pod-missed: WATI_TOKEN not set - nothing sent")
        return []
    status, params = template_info()
    if status != "APPROVED":
        _warn_hourly("approval", "pod-missed: %d POD(s) marked No are waiting - %s is %s in WATI, nothing sent until it is APPROVED"
                     % (len(due), MISSED_TEMPLATE, status or "not found"))
        return []
    quiet = dry or MISSED_DRY_RUN
    if cols.get("missed_new") and not quiet:        # add the Missed Call 1 column (and the grid column if needed)
        if info.get("src_cols", 0) <= cols["missed"]:
            svc.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID, body={"requests": [
                {"appendDimension": {"sheetId": info["src_gid"], "dimension": "COLUMNS",
                                     "length": cols["missed"] + 1 - info["src_cols"]}}]}).execute()
            info["src_cols"] = cols["missed"] + 1
        src_values(svc, "%s1" % _col(cols["missed"]), [[MISSED_HEADER]], False, cols)
        cols["missed_new"] = False
    mc = _col(cols["missed"])
    done = []
    for p in due:
        if len(wa_number(p["phone"])) < 11:
            src_values(svc, "%s%d" % (mc, p["row"]), [["Failed - no valid number"]], quiet, cols)
            done.append("pod-missed: PODs row %d %s has no valid number - not sent" % (p["row"], p["name"][:24]))
            continue
        if quiet:
            done.append("pod-missed: would send %s to %s (%s, POD %s %s)" % (MISSED_TEMPLATE, wa_number(p["phone"]), p["name"][:24], _d(p["d"]), p["t"]))
            continue
        ok = send_missed(p, params)
        if ok is None:
            continue
        stamp = ("1 - " if ok else "Failed - ") + now.strftime(STAMP_FMT)
        src_values(svc, "%s%d" % (mc, p["row"]), [[stamp]], False, cols)
        p["missed"] = stamp
        done.append("pod-missed: %s %s to %s (%s, POD %s %s) -> PODs %s%d"
                    % ("sent" if ok else "WATI refused", MISSED_TEMPLATE, wa_number(p["phone"]), p["name"][:24], _d(p["d"]), p["t"], mc, p["row"]))
        time.sleep(1)
    return done


# ── one pass ──────────────────────────────────────────────────────────────────
def run(svc, dry_run=False, now=None):
    """One pass. Returns the list of things done (or that would be done in a dry run)."""
    global _dropdowns_done
    now = now or _now()
    today = now.date()
    targets = {"TODAY": today, "TOMORROW": next_working_day(today)}
    done = []
    for _ in range(80):                                  # re-read after every structural change
        src, grid = read(svc)
        secs = sections(grid)
        if any(n not in secs or not secs[n]["rows"] for n in SECTIONS):
            raise RuntimeError("could not find the TODAY and TOMORROW grids on %s - nothing written" % DST_TAB)
        pods, cols = pods_by_date(src)
        info, w, again = None, None, False
        if not _dropdowns_done:                           # Yes/No chips on every grid row, once per start
            info = meta(svc)
            w = Writer(svc, info["gid"], dry_run)
            w.structure([w.dropdown(r["r"]) for r in secs["TODAY"]["rows"] if r["r"] not in info["merged"]])
            _dropdowns_done = not dry_run
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
                where, structural = place(w, sec, info["merged"], p, name)
                done.append("pods-today: %s %s | %s | %s -> %s (%s)" % (
                    name, p["t"], p["vals"][1][:24], p["vals"][2], where, target.strftime("%d/%m/%Y")))
                if structural:
                    again = True
                    break
            if again:
                break
        if again and not dry_run:
            continue
        # grids are current: send any missed-POD messages for Nos on the TODAY grid
        try:
            info = info or meta(svc)
            done += chase_missed(svc, info, secs, pods, cols, now, dry_run)
        except Exception as e:
            _warn_hourly("chase", "pod-missed: not sent this cycle: %s" % e)
        return done
    log.warning("pods-today: stopped after 80 passes")
    return done


def tick(get_svc):
    """Called every poller cycle; runs at most every INTERVAL_S."""
    global _last_tick
    if not ENABLED or time.time() - _last_tick < INTERVAL_S:
        return
    _last_tick = time.time()
    try:
        for line in run(get_svc(), dry_run=DRY_RUN):
            log.info(("[DRY RUN] " if DRY_RUN else "") + line)
    except Exception as e:
        _warn_hourly("tick", "pods-today: not updated this cycle: %s" % e)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    sys.path.insert(0, ".")
    import main
    go = os.getenv("GO") == "1"
    lines = run(main.get_sheets_service(), dry_run=not go)
    print("\n".join(lines) if lines else "pods-today: nothing to do - grids are up to date")
    if not go and lines:
        print("(preview only - run with GO=1 to write)")
