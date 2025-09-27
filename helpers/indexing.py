from helpers.linkProcessing import _resolve_shortcut, START_MENU_DIRS, _add_aliases, _enumerate_store_apps, UWP_PROC_HINTS, SKIP_TARGET_EXE, JUNK_WORDS, _dedupe_push
from pathlib import Path
from helpers.logging import log_aliases
import win32com.client

def build_name_index():
    """
    Returns: dict normalized_name -> [entry,...]
    Win32 entry:
      {"type":"lnk","display","lnk","target","args",
       "exe_path","exe_name","expected_proc_names","expected_proc_paths"}
    UWP entry:
      {"type":"uwp","display","aumid","expected_proc_names","expected_proc_paths": set()}
    """
    index: dict[str, list[dict]] = {}

    # 1) Win32 apps from Start Menu shortcuts
    for root in START_MENU_DIRS:
        for lnk in Path(root).rglob("*.lnk"):
            try:
                info = _resolve_shortcut(lnk)
                base = lnk.stem.lower()
                t = (info["target"] or "").strip()
                if any(w in base for w in JUNK_WORDS):
                    continue
                if not t.lower().endswith(".exe"):
                    continue  # skip docs/web/uninstallers
                exe_name = Path(t).name.lower()
                if exe_name in SKIP_TARGET_EXE:
                    continue  # skip control panel, open-folder, etc.

                exe_path = str(Path(t).resolve())
                entry = {
                    "type": "lnk",
                    "display": info["display"],       # <-- stem preferred
                    "lnk": info["lnk"],
                    "target": info["target"],
                    "args": info["args"],
                    "exe_path": exe_path,
                    "exe_name": exe_name,
                    "expected_proc_names": {exe_name},
                    "expected_proc_paths": {exe_path.lower()},
                }

                # aliases: display + desc + exe-derived (incl. Edge/Office hand-tuned)
                for alias in _add_aliases(
                        info["display"], exe_name=exe_name, alt_desc=info.get("alt_desc")):
                    _dedupe_push(index, alias, entry)
            except Exception as e:
                print(f"[index] failed on {lnk}: {e}")
                pass


    # 2) UWP / Store apps via AppsFolder
    for app in _enumerate_store_apps():
        display = app["name"].strip()
        aumid = app["aumid"]
        expected = {"applicationframehost.exe"}
        aumid_l = aumid.lower()
        for key, names in UWP_PROC_HINTS.items():
            if key in aumid_l:
                expected |= {n.lower() for n in names}

        entry = {
            "type": "uwp",
            "display": display,
            "aumid": aumid,
            "expected_proc_names": expected,
            "expected_proc_paths": set(),
        }
        for alias in _add_aliases(display):
            _dedupe_push(index, alias, entry)
        
    log_aliases(index)  # Debug: log all aliases after UWP addition

    return index