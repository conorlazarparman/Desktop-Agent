# jarvis_open_app.py
import os, sys, time, threading, queue, glob, ctypes
from pathlib import Path
from fuzzywuzzy import fuzz, process
import subprocess
USE_WHISPER = True

def log_index(index):
    print("\n=== Indexed Apps ===")
    for name, shortcuts in sorted(index.items()):
        print(f"{name}:")
        for s in shortcuts:
            print(f"   -> {s}")
    print("====================\n")

def log_launch(shortcut_path):
    import win32com.client
    from pathlib import Path

    shell = win32com.client.Dispatch("WScript.Shell")
    sc = shell.CreateShortCut(shortcut_path)

    print("\n=== Launch Debug Info ===")
    print(f"LNK name:   {Path(shortcut_path).name}")
    print(f"LNK path:   {shortcut_path}")
    print(f"Target exe: {sc.Targetpath}")
    print("=========================\n")

START_MENU_DIRS = [
    Path(os.environ["APPDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
    Path(os.environ["PROGRAMDATA"]) / r"Microsoft\Windows\Start Menu\Programs",
]

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
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    index = {}
    for lnk in find_shortcuts():
        try:
            shortcut = shell.CreateShortCut(str(lnk))
            display = normalize(Path(lnk).stem)
            # add variants
            names = {display}
            # examples: add alias “chrome” for “google chrome”
            if "google chrome" in display:
                names.update({"chrome"})
            if "microsoft edge" in display:
                names.update({"edge"})
            for n in names:
                index.setdefault(n, []).append(str(lnk))
        except Exception:
            pass
    return index

def best_match(query, index):
    candidates = list(index.keys())
    if not candidates:
        return None, None
    match, score = process.extractOne(normalize(query), candidates, scorer=fuzz.token_set_ratio)
    if score < 70:
        return None, None
    return match, index[match][0]

def launch(shortcut_path):
    # os.startfile handles .lnk nicely
    os.startfile(shortcut_path)

# --- Speech front-end (choose ONE) ---
USE_WHISPER = False  # flip to True if you configured faster-whisper

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
