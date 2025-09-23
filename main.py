# jarvis_open_app.py
from helpers.logging import log_launch, log_index, list_processes
from helpers.indexing import build_name_index
from helpers.searching import best_match
from helpers.listening import listen_text_once
from capabilities.launch import launch
from capabilities.close import close_entry

def parse_and_act(utterance, index):
    u = utterance.lower().strip()
    if not u:
        return "Didn’t catch that."

    # opening
    if u.startswith("open "):
        app = u.replace("open ", "", 1)
        display, entry = best_match(app, index)
        if not entry:
            return f"Couldn’t find an app like “{app}”."
        launch(entry)
        return f"Opening {display}."
    
    #closing
    if u.startswith("close "):
        app = u.replace("close ", "", 1)
        display, entry = best_match(app, index)
        if not entry:
            return f"Couldn’t find an app like “{app}”."
        terminated, killed = close_entry(entry, force=False)
        if terminated or killed:
            msg = f"Closed {display} ({terminated} terminated"
            if killed: msg += f", {killed} killed"
            msg += ")."
            return msg
        else:
            return f"Didn’t find any running processes for {display}."

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