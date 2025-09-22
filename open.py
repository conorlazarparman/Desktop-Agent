# jarvis_open_app.py
import os, sys, time, threading, queue, glob, ctypes
from pathlib import Path
from fuzzywuzzy import fuzz, process
import subprocess
import win32com.client
USE_WHISPER = True
from helpers.logging import log_launch, log_index
from helpers.linkProcessing import START_MENU_DIRS, _resolve_shortcut, _add_aliases, _enumerate_store_apps, normalize
from helpers.indexing import build_name_index

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

    # opening
    if u.startswith("open "):
        app = u.replace("open ", "", 1)
        match, path = best_match(app, index)
        if not match:
            return f"Couldn’t find an app like “{app}”."
        launch(path)
        log_launch(path)
        return f"Opening {match}."
    
    #listing processes
    if u.startswith("list processes"):
        list_processes()
        return "Listed running processes."
    return "Say your commands."

import psutil
def list_processes():
    print("=== Running Processes ===")
    for p in psutil.process_iter(["pid","name","exe","username"]):
        try:
            print(p.info)
        except psutil.Error:
            pass

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