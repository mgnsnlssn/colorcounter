# 🧮 Color Count Pro v3.0
Ett Python-baserat verktyg som automatiskt analyserar färgmarkerade celler i Excel-filer och skapar tydliga rapporter, statistik och trenddiagram.

---

## 🚀 Funktioner
- Bevakar mappen **`inbox/`** efter nya `.xlsx`-filer  
- Räknar antal **gröna**, **gula** och **röda** celler per rad  
- Skapar **statistikblad** per fil (text + färgfördelning)  
- Sammanställer allt i **`Sammanställning.xlsx`**  
- Ritar **trenddiagram** med färgkodade linjer (grön/gul/röd)  
- Kör helt automatiskt – bara lägg filen i *inbox*  

---

## ⚙️ Installation

### 1️⃣ Krav
- **Python 3.10+**
- **openpyxl**

### 2️⃣ Klona projektet
git clone https://github.com/mgnsnlssn/colorcounter.git
cd colorcounter


### 3️⃣ Skapa virtuell miljö (för att hålla projektet rent)

Detta skapar en egen Python-miljö för projektet där endast nödvändiga paket installeras.  
Kör följande kommando i projektmappen:

macOS / Linux 👇
```bash
python3 -m venv venv
source venv/bin/activate
```
Windows 👇
```bash
python -m venv venv
venv\Scripts\activate
```
### 4️⃣ Installera beroenden
Projektet använder **openpyxl** för att läsa och skriva Excel-filer.

Installera alla beroenden med:
```bash
pip install -r requirements.txt

## 🧑‍💻 Testa själv
Kopiera och kör direkt i terminalen:
```bash
git clone https://github.com/mgnsnlssn/colorcounter.git
cd colorcounter
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python color_count_pro.py
```
### 💡 Första gången du kör?

1. Radera eventuell gammal `venv`-mapp.
2. Kör:
   ```bash
   python3 -m venv venv
   source venv/bin/activate      # macOS/Linux
   python -m pip install --upgrade pip
   python -m pip install -r requirements.txt

### Gå till SchoolSoft, Närvaro - Elev - Elevnärvaro vecka, Copy/Paste in i ett Spreadsheet - > Exportera
### Namnge filen v42_y7.xlsx eller nåt liknande och 
## Lägg en Excel-fil i inbox/, så skapas resultat i outbox/ automatiskt.

