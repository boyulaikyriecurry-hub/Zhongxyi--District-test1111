from pathlib import Path
import pandas as pd
import json

DATA_DIR = Path(__file__).parent / "data"

DATA_XLSX = DATA_DIR / "放到網頁的data.xlsx"
INPUT_GEOJSON = DATA_DIR / "zhongxi_li_simple.geojson"
OUTPUT_GEOJSON = DATA_DIR / "zhongxi_with_capacity.geojson"

EXCLUDE_SHEETS = {
    "Summary",
    "ALL_LI_TOTAL",
    "Installed_Capacity_By_LI",
    "Installed_Capacity_Summary",
}

def build_capacity_map():
    xls = pd.ExcelFile(DATA_XLSX)
    capacity_map = {}

    for sheet in xls.sheet_names:
        if sheet in EXCLUDE_SHEETS:
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

    return capacity_map

def main():
    with open(INPUT_GEOJSON, "r", encoding="utf-8") as f:
        geo = json.load(f)

    capacity_map = build_capacity_map()

    match_count = 0
    no_match = []

    for feature in geo.get("features", []):
        props = feature.setdefault("properties", {})
        village = str(props.get("VILLNAME", "")).strip()

        cap = capacity_map.get(village)
        if cap is None:
            props["village"] = village
            props["bess_capacity_kwh"] = 0.0
            props["pv_capacity_kwh"] = 0.0
            no_match.append(village)
        else:
            props["village"] = village
            props["bess_capacity_kwh"] = cap["bess_capacity_kwh"]
            props["pv_capacity_kwh"] = cap["pv_capacity_kwh"]
            match_count += 1

    with open(OUTPUT_GEOJSON, "w", encoding="utf-8") as f:
        json.dump(geo, f, ensure_ascii=False)

    print("完成輸出：", OUTPUT_GEOJSON)
    print("geojson feature 數：", len(geo.get("features", [])))
    print("成功對應里數：", match_count)
    print("未對應里名：", sorted(set(no_match)))

if __name__ == "__main__":
    main()