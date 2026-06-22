"""Full accuracy and feature validation for Supply Mapping Analyzer."""
from __future__ import annotations

import json
import math
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from analyzer import PREVIEW_COLUMNS, analyze_file_path

DATA = Path(__file__).parent / "Supply Mapping.xlsx"


def _sanitize_for_json(obj: Any) -> Any:
    if obj is None:
        return None
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_for_json(v) for v in obj]
    if isinstance(obj, (np.floating, float)):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return float(obj)
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.bool_,)):
        return bool(obj)
    if isinstance(obj, (pd.Timestamp, datetime, date)):
        return obj.isoformat()
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    return obj


def _preview_dataframe(df: pd.DataFrame, limit: int = 200) -> list[dict]:
    cols = [c for c in PREVIEW_COLUMNS if c in df.columns]
    subset = df[cols].head(limit).copy()
    for col in subset.columns:
        if pd.api.types.is_datetime64_any_dtype(subset[col]):
            subset[col] = subset[col].map(
                lambda x: x.isoformat() if pd.notna(x) else None
            )
    return _sanitize_for_json(subset.replace({np.nan: None}).to_dict(orient="records"))


def main() -> int:
    if not DATA.exists():
        print(f"FAIL: missing {DATA}")
        return 1

    result = analyze_file_path(str(DATA))
    df = result["processed_df"]
    validation = result["validation"]

    checks = [v for v in validation if v["Check"] != "OVERALL QUALITY SCORE"]
    failed = [v for v in checks if v["Status"] != "PASS"]
    overall = next(v for v in validation if v["Check"] == "OVERALL QUALITY SCORE")

    print("=" * 50)
    print("SUPPLY MAPPING ANALYZER - FULL AUDIT")
    print("=" * 50)
    print(f"Candidates: {len(df)}")
    print(f"Validation: {overall['Detail']}")
    print()

    for v in checks:
        icon = "PASS" if v["Status"] == "PASS" else "FAIL"
        print(f"  [{icon}] {v['Check']}: {v['Detail']}")

    if failed:
        print(f"\n{len(failed)} check(s) FAILED")
        return 1

    # API / JSON
    payload = {
        "stats": result["stats"],
        "charts": result["charts"],
        "validation": validation,
        "processed_preview": _preview_dataframe(df, 5),
    }
    text = json.dumps(_sanitize_for_json(payload))
    assert "NaN" not in text
    json.loads(text)
    print("\n  [PASS] API JSON serialization")

    # Excel
    out = Path(__file__).parent / "Supply_Mapping_Analyzed.xlsx"
    out.write_bytes(result["excel_bytes"])
    import pandas as pd

    xl = pd.ExcelFile(out)
    expected_sheets = {
        "PBI_Dashboard_Layout",
        "Power BI Instructions",
        "PBI_Summary",
        "PBI_Funnel",
        "PBI_Status",
        "PBI_Sourcer_Chart",
        "PBI_Candidates",
        "PBI_Stage_Transitions",
        "PBI_Bottleneck_Rounds",
        "PBI_Bottleneck_Designation",
        "PBI_Bottleneck_Sourcer",
        "PBI_Slow_Movers_30Plus",
        "PBI_Designation_Timing",
        "PBI_Sourcer_Timing",
        "Recruitment Timeline",
        "Validation Report",
        "Summary",
        "Sourcer Summary",
        "Sourcer x Skills",
        "Designation x Exp",
        "Assumptions",
    }
    missing = expected_sheets - set(xl.sheet_names)
    if missing:
        print(f"FAIL: missing sheets {missing}")
        return 1
    print("  [PASS] Excel sheets:", ", ".join(xl.sheet_names))

    proc = pd.read_excel(xl, "Recruitment Timeline", header=1)
    if "####" in proc.astype(str).to_string():
        print("FAIL: #### found in Excel")
        return 1
    print("  [PASS] No #### in Excel output")

    val_sheet = pd.read_excel(xl, "Validation Report")
    fails = val_sheet[val_sheet["Status"] == "FAIL"]
    if len(fails):
        print("FAIL: Validation sheet has failures")
        return 1
    print("  [PASS] Validation Report all PASS")

    print("\n" + "=" * 50)
    print("ALL CHECKS PASSED - READY FOR HR REVIEW")
    print("=" * 50)
    return 0


if __name__ == "__main__":
    sys.exit(main())
