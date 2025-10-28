from openpyxl import load_workbook
import os

# --- Inställningar ---
INBOX = "inbox"
OUTBOX = "outbox"
RED_CODES = ["FF0000", "FF6666", "FFC7CE"]  # olika varianter av "röd" (Excel använder flera)
DAY_KEYWORDS = {
    "mån": "Monday",
    "tis": "Tuesday",
    "ons": "Wednesday",
    "tor": "Thursday",
    "fre": "Friday"
}


def is_red(cell):
    fill = cell.fill
    if fill and fill.start_color and fill.start_color.rgb:
        rgb = fill.start_color.rgb[-6:].upper()
        return rgb in RED_CODES
    return False


def detect_partial_presence(filepath):
    wb = load_workbook(filepath)
    ws = wb.active

    # --- Identifiera kolumner per dag ---
    day_columns = {}
    for idx, cell in enumerate(ws[1], start=1):
        if not cell.value:
            continue
        val = str(cell.value).lower()
        for sw, en in DAY_KEYWORDS.items():
            if sw in val:
                day_columns.setdefault(en, []).append(idx)

    print(f"Detected day-column mapping:\n{day_columns}\n")

    # --- Analys ---
    results = []

    for row in ws.iter_rows(min_row=2):  # hoppa över rubrikraden
        name = row[0].value
        if not name:
            continue
        for day, cols in day_columns.items():
            day_cells = [row[i - 1] for i in cols]
            reds = [is_red(c) for c in day_cells]
            # letar efter röd följd av icke-röd
            for i in range(len(reds) - 1):
                if reds[i] and not reds[i + 1]:
                    results.append((name, day))
                    break

    return results


def main():
    os.makedirs(OUTBOX, exist_ok=True)
    files = [f for f in os.listdir(INBOX) if f.endswith(".xlsx")]
    if not files:
        print("No Excel files found in inbox/")
        return

    for f in files:
        path = os.path.join(INBOX, f)
        print(f"Processing {f} ...")
        suspects = detect_partial_presence(path)

        if suspects:
            out_path = os.path.join(OUTBOX, f"partial_absence_{f}.txt")
            with open(out_path, "w", encoding="utf-8") as out:
                out.write("Students with red→present pattern:\n\n")
                for s in suspects:
                    out.write(f"{s[0]} – {s[1]}\n")
            print(f"→ Results saved to {out_path}")
        else:
            print("→ No partial absences detected.")


if __name__ == "__main__":
    main()