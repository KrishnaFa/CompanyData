#!/usr/bin/env python3
"""
Generate Power BI-ready Excel and CSV pack from Supply Mapping data.
No HTML or web server required — use this file for your manager's Power BI workflow.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from analyzer import analyze_file_path

DEFAULT_INPUT = Path(__file__).parent / "Supply Mapping.xlsx"
DEFAULT_EXCEL = Path(__file__).parent / "Supply_Mapping_PowerBI_Ready.xlsx"
DEFAULT_ZIP = Path(__file__).parent / "Recruitment_PowerBI_DataPack.zip"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export recruitment analytics for Power BI (no HTML dashboard needed)."
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_INPUT),
        help="Path to Supply Mapping.xlsx (default: ./Supply Mapping.xlsx)",
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

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}")
        return 1

    print(f"Analyzing {input_path.name}...")
    result = analyze_file_path(str(input_path))

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
    print(f"Excel : {excel_path}")
    print(f"CSV   : {zip_path}")
    print()
    print(f"Candidates      : {stats['total_candidates']}")
    print(f"Reached Offer   : {stats['reached_offer']}")
    print(f"Slow movers 30+ : {slow}")
    print(f"Main bottleneck : {primary_round}")
    print()
    print("NEXT STEPS FOR POWER BI:")
    print("  1. Open Power BI Desktop")
    print("  2. Get Data → Excel → select the Excel file above")
    print("  3. Load all PBI_* sheets")
    print("  4. Open sheet PBI_Dashboard_Layout — build visuals page by page")
    print("  5. Relate PBI_Stage_Transitions[Candidate Key] → PBI_Candidates[Candidate Key]")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
