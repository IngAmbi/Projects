def load_namedays(filename):
    nameday_dict = {}
    with open(filename, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split(" ", 1)
            if len(parts) < 2:
                continue
            date, names = parts
            for name in names.replace(" a ", "/").replace(",", "/").split("/"):
                name = name.strip().lower()
                if name:
                    nameday_dict[name] = date
    return nameday_dict