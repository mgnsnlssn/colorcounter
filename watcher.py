import time
import os
import subprocess

INBOX = "inbox"
CHECK_INTERVAL = 5  # sekunder

def get_xlsx_files():
    return {f for f in os.listdir(INBOX) if f.endswith(".xlsx")}

def run_script(script_name):
    print(f"\n▶ Running {script_name} ...")
    subprocess.run(["python", script_name])

def main():
    print("👁️  Color Count Watcher is now watching 'inbox/' (Ctrl+C to stop)...")
    seen = get_xlsx_files()

    while True:
        time.sleep(CHECK_INTERVAL)
        current = get_xlsx_files()
        new_files = current - seen

        if new_files:
            for f in new_files:
                print(f"\n📂 New file detected: {f}")
                # Kör båda analyserna
                run_script("color_count_pro.py")
                run_script("detect_partial_absence.py")
            seen = current

if __name__ == "__main__":
    main()