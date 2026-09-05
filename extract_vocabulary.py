import openpyxl, sys, io
from collections import Counter
import unicodedata

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

file_path = r"C:\Users\said\Downloads\APPELS-DOFFRE-AVIS-DATTRIBUTION-2026-GAZON.xlsx"
wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

for sname in ["APPEL D'OFFRE 2026", "AVIS D'ATTRIBUTIONS 2026"]:
    ws = wb[sname]
    rows = list(ws.iter_rows(values_only=True))
    data = [r for r in rows if r and r[0] is not None and str(r[0]).strip().isdigit()]
    
    actions = Counter()
    proj_types = Counter()
    
    for r in data:
        a = str(r[2]).strip() if len(r) > 2 and r[2] else ""
        t = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        actions[a] += 1
        proj_types[t] += 1
        
    print(f"\n==================== {sname} ({len(data)} rows) ====================")
    print("ALL Actions (Col 3):")
    for k, v in actions.most_common():
        print(f"  {k:40} : {v}")
    print("\nALL Project Types (Col 5):")
    for k, v in proj_types.most_common():
        print(f"  {k:40} : {v}")
