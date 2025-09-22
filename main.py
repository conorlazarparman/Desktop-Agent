# jarvis_open_app.py
import os
import subprocess
from helpers.logging import log_launch, log_index, list_processes
from helpers.linkProcessing import START_MENU_DIRS, _resolve_shortcut, _add_aliases, _enumerate_store_apps, normalize
from helpers.indexing import build_name_index
from helpers.searching import best_match
from helpers.listening import listen_text_once
from capabilities.launch import launch

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