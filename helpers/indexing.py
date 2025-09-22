from helpers.linkProcessing import START_MENU_DIRS, _resolve_shortcut, _add_aliases, _enumerate_store_apps
from pathlib import Path

def build_name_index():
    index = {}  # normalized name -> list of launch specs

    # 1) Include classic desktop apps from Start Menu shortcuts
    for root in START_MENU_DIRS:
        for lnk in Path(root).rglob("*.lnk"):
            try:
                info = _resolve_shortcut(lnk)
                # Optional pruning: keep only real apps (EXE targets) and skip docs/uninstallers
                t = info["target"].lower()
                base = Path(lnk).stem.lower()
                junk_words = ("uninstall", "help", "readme", "manual", "documentation", "support", "website")
                if (not t.endswith(".exe")) or any(w in base for w in junk_words):
                    continue

                for alias in _add_aliases(info["display"]):
                    index.setdefault(alias, []).append({
                        "type": "lnk",
                        "display": info["display"],
                        "lnk": info["lnk"],
                        "target": info["target"],
                        "args": info["args"],
                    })
            except Exception:
                pass

    # 2) Include Microsoft Store (UWP) apps from AppsFolder
    for app in _enumerate_store_apps():
        display = app["name"].strip()
        for alias in _add_aliases(display):
            # Avoid duplicates if a Win32 alias with same display already exists
            exists = any(e.get("display", "").lower() == display.lower() for e in index.get(alias, []))
            if exists:
                continue
            index.setdefault(alias, []).append({
                "type": "uwp",
                "display": display,
                "aumid": app["aumid"],
            })

    return index
