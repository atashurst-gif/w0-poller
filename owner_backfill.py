#!/usr/bin/env python3
"""One-off: give every older contact on Regen's WATI an `owner` attribute, so the inbox
filter (owner = Josh / Leah / Noor / ...) can include chats from before the owner tags
started (19 Oct 2026).

  * only contacts with NO owner value are touched; Josh / Leah / Noor / TPA are never changed
  * they get owner = OWNER_BACKFILL_VALUE (default "Regen")
  * runs in its own thread (WATI API only, no Google client), paced, so first messages are
    never held up; progress and completion are kept in owner_backfill.json on the volume,
    so a redeploy carries on instead of starting again; when finished it never runs again

Env: OWNER_BACKFILL_ENABLED (1), OWNER_BACKFILL_VALUE (Regen), OWNER_BACKFILL_DRY_RUN (0)
"""
import os
import json
import time
import logging
import threading

import requests

log = logging.getLogger(__name__)

ENABLED = os.getenv("OWNER_BACKFILL_ENABLED", "1") == "1"
VALUE   = os.getenv("OWNER_BACKFILL_VALUE", "Regen")
DRY_RUN = os.getenv("OWNER_BACKFILL_DRY_RUN", "0") == "1"
ATTR    = os.getenv("OWNER_ATTRIBUTE", "owner")
API     = os.getenv("WATI_API_URL", "https://eu-api.wati.io/602557")
TOKEN   = os.getenv("WATI_TOKEN", "")
PAUSE_S = float(os.getenv("OWNER_BACKFILL_PAUSE_S", "0.4"))


def _state_path(seen_file):
    return os.path.join(os.path.dirname(seen_file) or ".", "owner_backfill.json")


def _load(path):
    try:
        return json.load(open(path))
    except Exception:
        return {}


def _save(path, st):
    try:
        json.dump(st, open(path, "w"))
    except Exception as e:
        log.warning("owner-backfill: could not save progress: %s" % e)


def _page(n):
    for attempt in range(4):
        try:
            r = requests.get("%s/api/v1/getContacts" % API, headers={"Authorization": "Bearer " + TOKEN},
                             params={"pageSize": 100, "pageNumber": n}, timeout=40)
            if r.status_code == 200:
                return r.json().get("contact_list", []) or []
            log.warning("owner-backfill: page %d http %s" % (n, r.status_code))
        except Exception as e:
            log.warning("owner-backfill: page %d: %s" % (n, e))
        time.sleep(10 * (attempt + 1))
    return None


def _owner(c):
    for p in c.get("customParams") or []:
        if str(p.get("name", "")).strip().lower() == ATTR.lower():
            return str(p.get("value") or "").strip()
    return ""


def _set(wa):
    try:
        r = requests.post("%s/api/v1/updateContactAttributes/%s" % (API, wa),
                          headers={"Authorization": "Bearer " + TOKEN, "Content-Type": "application/json"},
                          json={"customParams": [{"name": ATTR, "value": VALUE}]}, timeout=20)
        return r.status_code in (200, 201) and '"result":true' in r.text.replace(" ", "")
    except Exception:
        return False


def _run(path):
    st = _load(path)
    if st.get("done"):
        return
    passes = st.get("passes", 0)
    while passes < 3:                       # contacts can move pages as they update, so sweep until a pass finds none
        page, changed, seen = st.get("page", 1), st.get("changed_this_pass", 0), st.get("seen_this_pass", 0)
        log.info("owner-backfill: pass %d starting at page %d%s" % (passes + 1, page, " [DRY RUN]" if DRY_RUN else ""))
        while True:
            cl = _page(page)
            if cl is None:
                log.warning("owner-backfill: WATI not answering - will carry on after the next restart")
                return
            if not cl:
                break
            for c in cl:
                seen += 1
                wa = "".join(ch for ch in str(c.get("wAid") or c.get("phone") or "") if ch.isdigit())
                if not wa or _owner(c):
                    continue
                if DRY_RUN or _set(wa):
                    changed += 1
                time.sleep(PAUSE_S)
            page += 1
            st.update(page=page, changed_this_pass=changed, seen_this_pass=seen)
            _save(path, st)
            if page % 10 == 0:
                log.info("owner-backfill: page %d, %d contacts checked, %d set to %s this pass" % (page, seen, changed, VALUE))
            if len(cl) < 100:
                break
        passes += 1
        st.update(passes=passes, page=1, changed_this_pass=0, seen_this_pass=0,
                  total_changed=st.get("total_changed", 0) + changed)
        _save(path, st)
        log.info("owner-backfill: pass %d done - %d contacts checked, %d set to owner=%s" % (passes, seen, changed, VALUE))
        if changed == 0 or DRY_RUN:
            break
    st["done"] = True
    _save(path, st)
    log.info("owner-backfill: finished - %d older contacts now have owner=%s" % (st.get("total_changed", 0), VALUE))


def start(seen_file):
    """Start the backfill once, in the background. Safe to call on every boot."""
    if not ENABLED or not TOKEN:
        return
    path = _state_path(seen_file)
    if _load(path).get("done"):
        return
    threading.Thread(target=_run, args=(path,), name="owner-backfill", daemon=True).start()
