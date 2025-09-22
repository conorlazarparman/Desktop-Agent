import psutil

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

def list_processes():
    print("=== Running Processes ===")
    for p in psutil.process_iter(["pid","name","exe","username"]):
        try:
            print(p.info)
        except psutil.Error:
            pass
