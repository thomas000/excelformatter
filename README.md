# ExcelFormatter

**ExcelFormatter** er et enkelt Windows-program for å formatere og generere PDF-/Excel-lister til bruk i Matvareexpressen. Programmet gjør det lett å generere brødlister, fryselister og plukklister direkte fra rå Excel-data – helt uten behov for manuell formatering i Excel.

---

## 🚀 Funksjoner

- Automatisk formatering av lister for:
  - Brødliste (en rute per side)
  - Fryseliste (en rute per side)
  - Fryseliste Plukk (alle ruter på én side)
  - Bamaliste Plukk (grønn header, alle ruter på én side)
- Velg listetype og dato fra GUI
- Genererer pene PDF-filer og nye Excel-filer
- Rask og enkel å bruke for alle ansatte – ingen Excel-kunnskap nødvendig

---

## 📦 Kom i gang

### 1. Last ned

Du kan laste ned ferdig **.exe**-fil under [Releases](https://github.com/thomas000/ExcelFormatter/releases) (eller få den tilsendt fra en kollega).

### 2. Kjør programmet

- Dobbelklikk på `ExcelFormatter.exe`
- Velg inn-fil (Excel, f.eks. `PSR-BREAD-2025-08-01.xlsx`)
- Velg listetype fra menyen
- Velg hva du vil gjøre (print, lagre som PDF, lagre som Excel)

### 3. Følg instruksene i programmet

---

## 🛠️ Bygg fra kildekode (for utviklere)

Du trenger **Python 3.10+** og følgende pakker:

```bash
pip install -r requirements.txt
