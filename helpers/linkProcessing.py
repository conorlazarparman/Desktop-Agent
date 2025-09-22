from pathlib import Path
import os, sys, time, threading, queue, glob, ctypes
import win32com.client

START_MENU_DIRS = [
    Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
    Path(os.environ["PROGRAMDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
]

def _add_aliases(name: str) -> set[str]:
    n = normalize(name)
    names = {n}
    if "google chrome" in n:
        names.add("chrome")
    if "microsoft edge" in n:
        names.add("edge")
    if "windows store" in n or "microsoft store" in n:
        names.update({"store", "microsoft store"})
    return names

def _resolve_shortcut(lnk_path: Path):
    shell = win32com.client.Dispatch("WScript.Shell")
    sc = shell.CreateShortCut(str(lnk_path))
    return {
        "display": (sc.Description or lnk_path.stem).strip(),
        "lnk": str(lnk_path),
        "target": sc.Targetpath or "",
        "args": sc.Arguments or "",
    }

def _enumerate_store_apps():
    """Return list of {'name': display_name, 'aumid': app_user_model_id} from shell:AppsFolder."""
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.NameSpace("shell:AppsFolder")
    items = folder.Items()
    apps = []
    # Iterate items; many are UWP, some are Win32. UWP ones have AUMID-like path with '!' bang.
    for i in range(items.Count):
        it = items.Item(i)
        name = it.Name  # display name shown in Start
        aumid = it.Path  # often looks like 'Microsoft.WindowsCalculator_8wekyb3d8bbwe!App'
        if name and aumid and "!" in aumid:  # heuristic to keep true UWP entries
            apps.append({"name": name, "aumid": aumid})
    return apps

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
