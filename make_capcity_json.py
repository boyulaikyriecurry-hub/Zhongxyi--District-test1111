from pathlib import Path
import pandas as pd
import json

DATA_DIR = Path(__file__).parent / "data"
DATA_XLSX = DATA_DIR / "放到網頁的data.xlsx"
OUTPUT_JSON = DATA_DIR / "capacity_map.json"

exclude_sheets = {
    "Summary",
    "ALL_LI_TOTAL",
    "Installed_Capacity_By_LI",
    "Installed_Capacity_Summary",
}

xls = pd.ExcelFile(DATA_XLSX)
capacity_map = {}

for sheet in xls.sheet_names:
    if sheet in exclude_sheets:
        continue

    df = pd.read_excel(DATA_XLSX, sheet_name=sheet)

    bess = 0.0
    pv = 0.0

    if "bess_capacity_kwh" in df.columns and not df.empty:
        bess = pd.to_numeric(df["bess_capacity_kwh"], errors="coerce").fillna(0).iloc[0]

    if "pv_capacity_kwh" in df.columns and not df.empty:
        pv = pd.to_numeric(df["pv_capacity_kwh"], errors="coerce").fillna(0).iloc[0]

    capacity_map[str(sheet).strip()] = {
        "bess_capacity_kwh": float(bess),
        "pv_capacity_kwh": float(pv),
    }

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(capacity_map, f, ensure_ascii=False, indent=2)

print("完成:", OUTPUT_JSON)
print("里數:", len(capacity_map))