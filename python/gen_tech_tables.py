import sys
import json
import math
import pandas as pd

VARS_XLSX = "vars_parsing.xlsx"

def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if pd.isna(v):
        return None
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if hasattr(v, "item"):
        return v.item()
    return v

def main():
    if len(sys.argv) < 2:
        print("Usage: python gen_tech_tables.py <CABINET_ID>", file=sys.stderr)
        sys.exit(1)

    cabinet_id = sys.argv[1]
    df = pd.read_excel(VARS_XLSX)

    row = df[df["id"].astype(str) == str(cabinet_id)]
    if row.empty:
        print(f"[gen_tech_tables] CABINET_ID={cabinet_id} not found", file=sys.stderr)
        sys.exit(2)

    rec = {k: clean(v) for k, v in row.iloc[0].to_dict().items()}
    print(json.dumps(rec, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()