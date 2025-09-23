from fuzzywuzzy import process, fuzz
from helpers.linkProcessing import normalize
SAFE_KILL_DENYLIST = {"explorer.exe","wininit.exe","winlogon.exe","services.exe","lsass.exe", "ApplicationFrameHost.exe"}
import psutil

def best_match(query, index):
    candidates = list(index.keys())
    if not candidates:
        return None, None
    key, score = process.extractOne(normalize(query), candidates, scorer=fuzz.token_set_ratio)
    if score < 70:
        return None, None
    # Return the display name and the first launch spec (lnk/uwp)
    entry = index[key][0]
    return entry.get("display", key), entry

def _pids_matching(entry):
    names = {n.lower() for n in entry.get("expected_proc_names", set())}
    paths = {p.lower() for p in entry.get("expected_proc_paths", set())}
    pids = []
    for p in psutil.process_iter(["pid","name","exe"]):
        try:
            nm = (p.info["name"] or "").lower()
            exep = (p.info["exe"] or "").lower()
        except psutil.Error:
            continue
        if nm in SAFE_KILL_DENYLIST:
            continue
        if exep and exep in paths:
            pids.append(p.info["pid"]); continue
        if nm in names:
            pids.append(p.info["pid"]); continue
    return pids
