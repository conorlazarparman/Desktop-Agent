from pathlib import Path
import os, sys, time, threading, queue, glob, ctypes
import win32com.client
import re

START_MENU_DIRS = [
    Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
    Path(os.environ["PROGRAMDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
]
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
VENDOR_PREFIXES = ("microsoft ", "microsoft® ", "ms ", "intel ", "intel® ", "dell ")
JUNK_WORDS = {
    "uninstall","help","readme","manual","documentation","support","website","release notes"
}
SKIP_TARGET_EXE = {"explorer.exe", "control.exe", "hh.exe"}  # not real app launches: again, looking for suitable one size fits all replacement to this. 


def _add_aliases(display_name: str, *, exe_name: str | None = None, alt_desc: str | None = None) -> set[str]:
    """Aliases: normalized display, vendor-stripped, optional description, exe-derived, and hand-tuned."""
    aliases: set[str] = set()

    def _n(x: str): return normalize(x)

    # 1) the display itself
    if display_name:
        aliases.add(_n(display_name))

        # 2) vendor-stripped form (e.g., 'Microsoft Edge' -> 'edge')
        low = display_name.lower()
        for pref in VENDOR_PREFIXES:
            if low.startswith(pref):
                aliases.add(_n(display_name[len(pref):]))
                break

    # 3) description as a backup alias, if present
    if alt_desc:
        aliases.add(_n(alt_desc))

    # 4) exe-derived aliases
    if exe_name:
        exe = exe_name.lower()
        base = exe[:-4] if exe.endswith(".exe") else exe        # 'msedge.exe' -> 'msedge'
        # generic exe base and number-stripped base (studio64 -> studio)
        aliases.add(_n(base))
        aliases.add(_n(re.sub(r"\d+$", "", base)))
        # hand-tuned
        if exe in EXE_ALIAS:
            aliases |= {_n(a) for a in EXE_ALIAS[exe]}

    # prune empties
    return {a for a in aliases if a}

def _dedupe_push(index: dict, alias: str, entry: dict):
    """Avoid duplicates per alias (by display + exe_path/aumid)."""
    bucket = index.setdefault(alias, [])
    sig = (entry["display"].lower(),
           entry.get("exe_path","").lower(),
           entry.get("aumid","").lower())
    for e in bucket:
        if (e["display"].lower(),
            e.get("exe_path","").lower(),
            e.get("aumid","").lower()) == sig:
            return
    bucket.append(entry)

def _resolve_shortcut(lnk_path: Path):
    """Prefer the .lnk filename (stem) for display; keep Description as an extra alias."""
    shell = win32com.client.Dispatch("WScript.Shell")
    sc = shell.CreateShortCut(str(lnk_path))
    stem = lnk_path.stem.strip()                  # <-- canonical display
    desc = (sc.Description or "").strip()         # <-- optional alias
    return {
        "display": stem,
        "alt_desc": desc,
        "lnk": str(lnk_path),
        "target": sc.Targetpath or "",
        "args": sc.Arguments or "",
        "startin": sc.WorkingDirectory or "",
    }

def _enumerate_store_apps():
    """Yield {'name': display_name, 'aumid': app_user_model_id} from shell:AppsFolder."""
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.NameSpace("shell:AppsFolder")
    items = folder.Items()
    for i in range(items.Count):
        try:
            it = items.Item(i)
            name = (it.Name or "").strip()          # defensive: handle None/whitespace
            aumid = (it.Path or "").strip()         # UWP looks like 'Vendor.App_...!App'
            if name and aumid and "!" in aumid:     # keep true UWP entries
                yield {"name": name, "aumid": aumid}
        except Exception:
            # some items can throw COM errors; skip them
            pass

def find_shortcuts():
    links = []
    for root in START_MENU_DIRS:
        for p in root.rglob("*.lnk"):
            links.append(p)
    return links

def normalize(name: str) -> str:
    n = name.lower()
    junk = ["(x64)", "(x86)", "inc.", "ltd.", "microsoft ", "corporation", "app"]
    for j in junk:
        n = n.replace(j, "")
    return n.strip()
