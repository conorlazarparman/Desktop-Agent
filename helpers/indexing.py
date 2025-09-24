from helpers.linkProcessing import START_MENU_DIRS, _resolve_shortcut, _add_aliases, _enumerate_store_apps
from pathlib import Path

JUNK_WORDS = {"uninstall","help","readme","manual","documentation","support","website","release notes"}
UWP_PROC_HINTS = {
    "microsoft.windowsstore": {"winstore.app.exe","applicationframehost.exe"},
    "microsoft.windowscalculator": {"applicationframehost.exe"},
    "microsoft.todos": {"applicationframehost.exe"},
    "microsoft.edge": {"msedge.exe"},
}
# Hand-tuned aliases for common apps that don't map cleanly from exe -> spoken name
EXE_ALIAS = {
    "msedge.exe": {"edge", "microsoft edge"},
    "chrome.exe": {"chrome", "google chrome"},
    "code.exe": {"code", "vscode", "visual studio code"},
    "devenv.exe": {"visual studio", "vs", "visual studio 2022"},
    "winword.exe": {"word", "microsoft word"},
    "excel.exe": {"excel", "microsoft excel"},
    "powerpnt.exe": {"powerpoint", "microsoft powerpoint"},
    "vlc.exe": {"vlc", "vlc media player"},
    "robloxplayerbeta.exe": {"roblox", "roblox player"},
    "robloxstudioinstaller.exe": {"roblox studio"},   # installer; real proc differs at runtime
}

def build_name_index():
    """
    Returns: dict normalized_name -> [entry,...]
    entry for Win32 (.lnk):
      {"type":"lnk","display","lnk","target","args",
       "exe_path","exe_name","expected_proc_names","expected_proc_paths"}
    entry for UWP:
      {"type":"uwp","display","aumid","expected_proc_names","expected_proc_paths": set()}
    """
    index = {}

    # 1) Win32 apps from Start Menu shortcuts
    for root in START_MENU_DIRS:
        for lnk in Path(root).rglob("*.lnk"):
            try:
                info = _resolve_shortcut(lnk)
                base = Path(lnk).stem.lower()
                t = (info["target"] or "").lower()
                if any(w in base for w in JUNK_WORDS): 
                    continue
                if not t.endswith(".exe"):
                    continue  # skip docs/web/uninstallers

                exe_path = str(Path(info["target"]).resolve())
                exe_name = Path(exe_path).name.lower()

                entry = {
                    "type": "lnk",
                    "display": info["display"],
                    "lnk": info["lnk"],
                    "target": info["target"],
                    "args": info["args"],
                    "exe_path": exe_path,
                    "exe_name": exe_name,
                    "expected_proc_names": {exe_name},
                    "expected_proc_paths": {exe_path.lower()},
                }
                for alias in _add_aliases(info["display"]):
                    index.setdefault(alias, []).append(entry)
            except Exception:
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
            exists = any(e.get("display","").lower()==display.lower() for e in index.get(alias, []))
            if not exists:
                index.setdefault(alias, []).append(entry)

    return index
