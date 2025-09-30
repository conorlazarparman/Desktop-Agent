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

def list_processes(index=None):
    print("=== Running Processes ===")
    for p in psutil.process_iter(["pid", "name", "exe", "username"]):
        try:
            info = p.info
            line = f"{info}"

            # If we have an index, check which entries this process could satisfy
            if index:
                pname = (info.get("name") or "").lower()
                pexe  = (info.get("exe") or "").lower()

                matched_aliases = []
                for alias, entries in index.items():
                    for e in entries:
                        if pname in e.get("expected_proc_names", set()):
                            matched_aliases.append(f"{alias} (name)")
                        if pexe and pexe in e.get("expected_proc_paths", set()):
                            matched_aliases.append(f"{alias} (path)")

                if matched_aliases:
                    line += "  <-- matches: " + ", ".join(matched_aliases)

            print(line)
        except psutil.Error:
            pass

def log_aliases(index):
    print("\n=== Aliases in Index ===")
    for alias, entries in sorted(index.items()):
        print(f"Alias: '{alias}'")
        for entry in entries:
            display = entry.get("display", "")
            etype = entry.get("type", "")
            exe_name = entry.get("exe_name", "")
            aumid = entry.get("aumid", "")
            print(f"   -> Type: {etype}, Display: '{display}', Exe: '{exe_name}', AUMID: '{aumid}'")
    print("========================\n")