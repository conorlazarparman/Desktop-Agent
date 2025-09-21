# jarvis_open_app.py
import os, sys, time, threading, queue, glob, ctypes
from pathlib import Path
from fuzzywuzzy import fuzz, process
import subprocess
import win32com.client
USE_WHISPER = True

def log_index(index):
    print("\n=== Indexed Apps ===")
    for name, shortcuts in sorted(index.items()):
        print(f"{name}:")
        for s in shortcuts:
            print(f"   -> {s}")
    print("====================\n")

def log_launch(entry):
    print("\n=== Launch Debug Info ===")
    print(f"Type:     {entry['type']}")
    print(f"Display:  {entry.get('display','')}")
    if entry["type"] == "lnk":
        print(f"LNK:      {entry['lnk']}")
        print(f"Target:   {entry.get('target','')}")
        print(f"Args:     {entry.get('args','')}")
    else:
        print(f"AUMID:    {entry['aumid']}")
    print("=========================\n")


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

def launch(entry):
    if entry["type"] == "lnk":
        # .lnk path opens the Win32 app
        os.startfile(entry["lnk"])
    elif entry["type"] == "uwp":
        # Launch Store app via AppsFolder + AUMID
        aumid = entry["aumid"]
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{aumid}"])
    else:
        raise ValueError(f"Unknown launch type: {entry['type']}")
# --- Speech front-end (choose ONE) ---
USE_WHISPER = True  # flip to True if you configured faster-whisper

def listen_text_once():
    # Whisper-only: no SpeechRecognition, no PyAudio
    import sounddevice as sd
    import numpy as np
    from faster_whisper import WhisperModel

    # init model once and cache on function attribute
    model = getattr(listen_text_once, "_model", None)
    if model is None:
        # pick a small model to start: "base" or "tiny"
        listen_text_once._model = model = WhisperModel("base", device="cpu", compute_type="int8")

    fs = 16000
    duration = 4.0  # seconds of audio per turn
    print("Listening… (Whisper)")
    audio = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype="float32")
    sd.wait()

    # Whisper expects either a path or a 1-D float32 array
    mono = audio.squeeze().astype(np.float32)

    segments, _ = model.transcribe(mono, language="en")
    text = " ".join(s.text for s in segments).strip()
    return text

def parse_and_act(utterance, index):
    u = utterance.lower().strip()
    if not u:
        return "Didn’t catch that."
    # simple intent
    if u.startswith("open "):
        app = u.replace("open ", "", 1)
        match, path = best_match(app, index)
        if not match:
            return f"Couldn’t find an app like “{app}”."
        launch(path)
        log_launch(path)
        return f"Opening {match}."
    return "Say: “open <app>”."

def main():
    print("Indexing apps...")
    index = build_name_index()
    print(f"Indexed {sum(len(v) for v in index.values())} shortcuts.")

    # log all apps/shortcuts
    log_index(index)
    
    print('Press Enter to talk; say: "open <app>" (Ctrl+C to quit).')
    while True:
        input()
        utt = listen_text_once()
        print("Heard:", utt)
        resp = parse_and_act(utt, index)
        print(resp)

if __name__ == "__main__":
    main()