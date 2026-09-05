import openpyxl, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

file_path = r"C:\Users\said\Downloads\APPELS-DOFFRE-AVIS-DATTRIBUTION-2026-GAZON.xlsx"
wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

for sname in ["APPEL D'OFFRE 2026", "AVIS D'ATTRIBUTIONS 2026"]:
    ws = wb[sname]
    print(f"\n--- {sname} (Rows 4 to 20) ---")
    rows = list(ws.iter_rows(min_row=4, max_row=20, values_only=True))
    for r in rows:
        if r[0] is not None:
            # Col 3 (Action), Col 5 (Type), Col 8 (Annonceur), Col 9 (Wilaya), Col 10 (Commune)
            print(f"N°{r[0]:<3} | Action (Col 3): {r[2]:<30} | Type (Col 5): {r[4]:<30} | Wilaya: {r[8]:<15} | Commune: {r[9]}")
