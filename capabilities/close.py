import psutil
from helpers.searching import _pids_matching, SAFE_KILL_DENYLIST

def close_entry(entry, force=False, wait_secs=3):
    """Find PIDs using precomputed exe names/paths and terminate them."""
    pids = _pids_matching(entry)
    if not pids:
        return 0, 0  # (terminated, killed)

    procs = []
    for pid in pids:
        try:
            procs.append(psutil.Process(pid))
        except psutil.Error:
            pass

    for p in procs:
        try:
            if (p.name() or "").lower() in SAFE_KILL_DENYLIST:
                continue
            p.terminate()
        except psutil.Error:
            pass

    gone, alive = psutil.wait_procs(procs, timeout=wait_secs)
    killed = 0
    if force and alive:
        for p in alive:
            try:
                p.kill()
                killed += 1
            except psutil.Error:
                pass
    return len(gone), killed
