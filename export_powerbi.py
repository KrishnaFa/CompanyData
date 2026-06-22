#!/usr/bin/env python3
"""
Generate Power BI-ready Excel and CSV pack from any Supply Mapping-format .xlsx file.
No HTML or web server required.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from analyzer import analyze_file_path, resolve_input_xlsx

DEFAULT_EXCEL = Path(__file__).parent / "Supply_Mapping_PowerBI_Ready.xlsx"
DEFAULT_ZIP = Path(__file__).parent / "Recruitment_PowerBI_DataPack.zip"


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Export recruitment analytics for Power BI from ANY Supply Mapping-format .xlsx file."
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=None,
        help=(
            "Path to any .xlsx file (Supply Mapping format). "
            "If omitted: uses newest file in uploads/ or project folder."
        ),
    )
    parser.add_argument(
        "--excel",
        default=str(DEFAULT_EXCEL),
        help="Output Excel path (default: ./Supply_Mapping_PowerBI_Ready.xlsx)",
    )
    parser.add_argument(
        "--zip",
        default=str(DEFAULT_ZIP),
        help="Output CSV zip path (default: ./Recruitment_PowerBI_DataPack.zip)",
    )
    args = parser.parse_args()

    try:
        input_path = resolve_input_xlsx(args.input)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Analyzing: {input_path.name}")
    print(f"Full path: {input_path}")
    try:
        result = analyze_file_path(str(input_path))
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return 1
    except Exception as exc:
        print(f"ERROR: Failed to analyze file: {exc}")
        return 1

    excel_path = Path(args.excel)
    zip_path = Path(args.zip)
    excel_path.write_bytes(result["excel_bytes"])
    zip_path.write_bytes(result["powerbi_csv_zip"])

    stats = result["stats"]
    slow = len(result["slow_movers"])
    primary = result["bottleneck_rounds"]
    primary_round = "—"
    if len(primary):
        flagged = primary[primary["Is Primary Bottleneck"] == True]  # noqa: E712
        if len(flagged):
            primary_round = flagged.iloc[0]["Transition"]

    print()
    print("=" * 56)
    print("POWER BI EXPORT COMPLETE")
    print("=" * 56)
    print(f"Input : {input_path.name}")
    print(f"Excel : {excel_path}")
    print(f"CSV   : {zip_path}")
    print()
    print(f"Candidates      : {stats['total_candidates']}")
    print(f"Reached Offer   : {stats['reached_offer']}")
    print(f"Slow movers 30+ : {slow}")
    print(f"Main bottleneck : {primary_round}")
    print()
    print("Works with ANY .xlsx that has Supply Mapping columns.")
    print("New file next time? Re-run this script — Power BI Refresh updates all charts.")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
