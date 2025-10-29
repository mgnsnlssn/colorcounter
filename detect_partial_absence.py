from openpyxl import load_workbook
import os

# --- Inställningar ---
INBOX = "inbox"
OUTBOX = "outbox"
RED_CODES = ["FF0000", "FF6666", "FFC7CE"]   # olika röda nyanser
YELLOW_CODES = ["FFFF00", "FFF200", "FFEB9C"]  # olika gula nyanser
DAY_KEYWORDS = {
    "mån": "Monday",
    "tis": "Tuesday",
    "ons": "Wednesday",
    "tor": "Thursday",
    "fre": "Friday"
}


def get_rgb(cell):
    fill = cell.fill
    if fill and fill.start_color and fill.start_color.rgb:
        return fill.start_color.rgb[-6:].upper()
    return None


def detect_yellow_to_red(filepath):
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

    results = []

    for row in ws.iter_rows(min_row=2):
        name = row[0].value
        if not name:
            continue

        for day, cols in day_columns.items():
            day_cells = [row[i - 1] for i in cols]
            colors = [get_rgb(c) for c in day_cells]

            # kolla om en gul följs av röd
            for i in range(len(colors) - 1):
                if (colors[i] in YELLOW_CODES) and (colors[i + 1] in RED_CODES):
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
        suspects = detect_yellow_to_red(path)

        if suspects:
            out_path = os.path.join(OUTBOX, f"yellow_to_red_{f}.txt")
            with open(out_path, "w", encoding="utf-8") as out:
                out.write("Students with yellow→red pattern (late → absent):\n\n")
                for s in suspects:
                    out.write(f"{s[0]} – {s[1]}\n")
            print(f"→ Results saved to {out_path}")
        else:
            print("→ No yellow→red transitions detected.")


if __name__ == "__main__":
    main()