import os
import subprocess
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
