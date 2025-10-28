# 🧮 Color Count Pro v3.0

A Python tool that automatically analyzes color-coded cells in Excel files and generates clear reports, stats, and trend charts.

⸻

## 🚀 Features
•	Watches the inbox/ folder for new .xlsx files
•	Counts green, yellow, and red cells per row
•	Creates a per-file stats sheet (text + color distribution)
•	Aggregates everything into Sammanställning.xlsx (summary workbook)
•	Draws trend charts with color-coded lines (green/yellow/red)
•	Fully automatic — just drop a file into inbox/

⸻

## ⚙️ Installation

### 1) Requirements
	•	Python 3.10+
	•	openpyxl

### 2) Clone the repo
```bash
git clone https://github.com/mgnsnlssn/colorcounter.git
cd colorcounter
```

### 3) Create a virtual environment
MacOS/Linux
```bash
python3 -m venv venv
source venv/bin/activate
```

Windows
```bash
python -m venv venv
venv\Scripts\activate
```
### 4) Install dependencies
```bash
pip install -r requirements.txt
```

## 🧑‍💻 Quick start
```bash
git clone https://github.com/mgnsnlssn/colorcounter.git
cd colorcounter
python3 -m venv venv && source venv/bin/activate     # macOS/Linux
# or: python -m venv venv && venv\Scripts\activate   # Windows
pip install -r requirements.txt
python color_count_pro.py
```
### 📂 How it works
 - Put one or more .xlsx files in inbox/.
 - Run python color_count_pro.py.
 - Results are written to outbox/ and to Sammanställning.xlsx (the aggregated summary workbook).
 - Each processed file also gets its own stats sheet with counts for green/yellow/red per row and a trend chart.

The tool reads cell fill colors (not text) and tallies them per row so you can quantify status, progress, or attendance at a glance.

### 📝 Typical workflow (SchoolSoft example)
 - In SchoolSoft: go to Närvaro → Elev → Elevnärvaro vecka
 - Copy/Paste to a spreadsheet and export to Excel
 - Name the file e.g. v42_y7.xlsx
 - Drop it into inbox/ → run the script → check outbox/ and Sammanställning.xlsx
