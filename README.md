# Recruitment Pipeline Analyzer — Power BI Edition

Analyze **Supply Mapping** Excel data: predict missing **R1 interview dates** using a 7-level hierarchical model based on **Designation**, **Skills**, and **submission cohort**, then visualise stage timing (R0→R1→R2→Offer), bottlenecks, and slow-movers in **Power BI** — no HTML website required.

---

## What's New — R1 Date Prediction Model

Missing R1 dates are now **automatically predicted** rather than just gap-averaged.  
The model tries 7 levels from most-specific to least-specific:

| Level | Match Criteria | Method |
|-------|---------------|--------|
| **L1** | Same Designation + Skills + exact R0 date | Median actual R1 date of that exact batch |
| **L2** | Same Designation + Skills + R0 month | Median actual R1 date of that month-cohort |
| **L3** | Same Designation + R0 month | Median actual R1 date for designation in that month |
| **L4** | Same Designation + Skills | R0 + median R0→R1 gap for that pair |
| **L5** | Same Designation only | R0 + median R0→R1 gap for that designation |
| **L6** | Same Skills only | R0 + median R0→R1 gap for that skill set |
| **L7** | Global fallback | R0 + global median gap |

Every predicted date is labelled with its level in the **`R1 Assumption / Note`** column for full traceability in Power BI.  
Designation aliases are normalised automatically (`SSE → Senior Software Engineer`, `deveops → DevOps Engineer`, etc.).

---

## All Implemented Features

| Feature | Status | Power BI sheet |
|--------|--------|----------------|
| **R1 Date Prediction** (7-level hierarchy) | ✅ New | `PBI_Candidates`, `Recruitment Timeline` |
| Days R0→R1, R1→R2, R0→Offer by **Designation** | ✅ | `PBI_Bottleneck_Designation`, `PBI_Designation_Timing` |
| Same timing by **Sourcer** | ✅ | `PBI_Bottleneck_Sourcer`, `PBI_Sourcer_Timing` |
| Designations with **30+ day** delays | ✅ | `Count R0→R1 30+ Days`, `Any Stage 30+ Flag` |
| **Bottleneck #1** — Designation | ✅ | `PBI_Bottleneck_Designation` |
| **Bottleneck #2** — Round (R0→R1 vs R1→R2 vs R0→Offer) | ✅ | `PBI_Bottleneck_Rounds` |
| **Bottleneck #3** — Slowest sourcers | ✅ | `PBI_Bottleneck_Sourcer` |
| Slow mover list (30+ days) | ✅ | `PBI_Slow_Movers_30Plus` |
| KPI summary cards | ✅ | `PBI_Summary` |
| Recruitment funnel | ✅ | `PBI_Funnel` |
| Status breakdown | ✅ | `PBI_Status` |
| Sourcer bar chart data | ✅ | `PBI_Sourcer_Chart` |
| Per-candidate timeline | ✅ | `PBI_Candidates` |
| Stage transitions (long format for charts) | ✅ | `PBI_Stage_Transitions` |
| Dashboard build guide | ✅ | `PBI_Dashboard_Layout` |
| Data validation (13 checks) | ✅ | `Validation Report` |
| Power BI export (Excel + CSV zip) | ✅ | See output files below |

---

## Requirements

- **Python 3.9+**
- **Power BI Desktop** (free from Microsoft)
- Input file: **Any `.xlsx` file** with Supply Mapping columns (see below)

---

## Install Python packages (one time)

```bash
cd /path/to/CompanyData
python3 -m pip install -r requirements.txt
```

---

## How to Run

### Step 1 — Drop your Excel file

Place your `Supply Mapping.xlsx` (or any name) in the project folder or in the `uploads/` subfolder.

### Step 2 — Generate predictions + Power BI export

**Mac / Linux / terminal:**
```bash
python3 export_powerbi.py
```

**Windows (double-click):**
```
run.bat
```

**Custom input / output paths:**
```bash
python3 export_powerbi.py "C:\Data\New_Supply_Mapping.xlsx" \
  --excel "C:\Reports\PowerBI_Ready.xlsx" \
  --zip "C:\Reports\PowerBI_DataPack.zip"
```

### Step 3 — Validate data quality

```bash
python3 validate_all.py
```

Expected output: `ALL CHECKS PASSED - READY FOR HR REVIEW` (13/13 checks).

### Step 4 — Verify prediction model output (optional)

```bash
python3 scratch/verify_predictions.py
```

Shows how many candidates were predicted at each level (L1–L7) and 5 sample rows.

---

## Output files

| File | Purpose |
|------|---------|
| `Supply_Mapping_PowerBI_Ready.xlsx` | **Main file — import into Power BI** |
| `Recruitment_PowerBI_DataPack.zip` | Same data as CSVs (optional) |
| `Supply_Mapping_Analyzed.xlsx` | Created when you run `validate_all.py` |

---

## Power BI — Full Setup (First Time)

### Step 1 — Import data

1. Open **Power BI Desktop**.
2. **Home → Get data → Excel**.
3. Select **`Supply_Mapping_PowerBI_Ready.xlsx`**.
4. In the Navigator, tick **every sheet starting with `PBI_`** (13 sheets).
5. Click **Load**.

Sheets to load:
- `PBI_Dashboard_Layout`
- `PBI_Summary`
- `PBI_Funnel`
- `PBI_Status`
- `PBI_Sourcer_Chart`
- `PBI_Candidates`
- `PBI_Stage_Transitions`
- `PBI_Bottleneck_Rounds`
- `PBI_Bottleneck_Designation`
- `PBI_Bottleneck_Sourcer`
- `PBI_Slow_Movers_30Plus`
- `PBI_Designation_Timing`
- `PBI_Sourcer_Timing`

### Step 2 — Create one relationship

1. Open **Model** view (left sidebar).
2. Drag **`Candidate Key`** from `PBI_Stage_Transitions` onto **`Candidate Key`** in `PBI_Candidates`.
3. Cardinality: **Many to one** | Cross-filter: **Single**.

---

## Power BI — Build All Charts

### Page 1 — Executive Overview

| Visual | Table | Fields |
|--------|-------|--------|
| **Card** (repeat for each KPI) | `PBI_Summary` | Filter `Metric`, show `Value` |
| **Funnel** | `PBI_Funnel` | `Stage`, `Count`, sort by `Stage Order` |
| **Donut chart** | `PBI_Status` | `Status`, `Count` |
| **Clustered bar chart** | `PBI_Sourcer_Chart` | `Sourcer`, `Total Submitted`, `Overall Conversion %` |

Suggested KPI cards from `PBI_Summary`:
- Total Candidates
- Reached R1
- Reached R2
- Reached Offer
- Avg Days R0→R1
- Slow Movers 30+ Days

### Page 2 — Bottlenecks & Timing

| Visual | Table | Fields |
|--------|-------|--------|
| **Bar chart** | `PBI_Bottleneck_Rounds` | `Transition`, `Avg Days` |
| **Table** | `PBI_Bottleneck_Designation` | Designation, Avg Days R0→R1, Count 30+, Primary Bottleneck |
| **Table** | `PBI_Bottleneck_Sourcer` | Sourcer, Avg Days R0→R1, Count 30+, Primary Bottleneck |
| **Clustered column chart** | `PBI_Sourcer_Chart` | `Sourcer`, `Avg Days R0→R1`, `Avg Days R1→R2` |

### Page 3 — Candidate Detail & Slow Movers

| Visual | Table | Fields |
|--------|-------|--------|
| **Matrix** | `PBI_Stage_Transitions` | Rows: `Designation`, Columns: `Transition`, Values: Avg `Days` |
| **Table** | `PBI_Slow_Movers_30Plus` | Name, Sourcer, Designation, Slowest Transition, Slowest Days |
| **Table** | `PBI_Candidates` | Name, Sourcer, Days R0→R1, Days R1→R2, Any Stage 30+ Days |

### Page 4 — R1 Date Predictions (New)

| Visual | Table | Fields |
|--------|-------|--------|
| **Clustered bar chart** | `PBI_Designation_Timing` | `Designation`, `Avg Days R0→R1`, `Median Days R0→R1` |
| **Clustered bar chart** | `PBI_Sourcer_Timing` | `Sourcer`, `Avg Days R0→R1`, `Median Days R0→R1` |
| **Table** | `PBI_Candidates` | Name, Designation, Skills, R1 Date (if available), R1 Data Source, R1 Assumption / Note |
| **Slicer** | `PBI_Candidates` | `R1 Data Source` — filter by `Actual`, `Imputed`, `Yr Corrected` |

> **Tip:** Use the `R1 Data Source` slicer to compare actual vs predicted R1 dates side-by-side.

### Step 4 — Save the report

**File → Save as** → e.g. `Recruitment_Dashboard.pbix`

You only build charts **once**. New data updates all visuals automatically after Refresh.

---

## New Data — Update Charts Automatically

When someone uploads new Supply Mapping data:

```
New Supply Mapping.xlsx
      ↓
  python3 export_powerbi.py    (or run.bat on Windows)
      ↓
  Supply_Mapping_PowerBI_Ready.xlsx  (overwritten)
      ↓
  Power BI Desktop → Refresh
      ↓
  All visuals update automatically
```

Step-by-step:
1. Replace the old Excel with the new `Supply Mapping.xlsx`.
2. Run `python3 export_powerbi.py`.
3. Open `Recruitment_Dashboard.pbix` in Power BI Desktop.
4. Click **Home → Refresh**.
5. All KPI cards, funnel, bar charts, bottleneck tables, and slow-mover lists reload from the new file.

> **Important:** Keep the same output path (`Supply_Mapping_PowerBI_Ready.xlsx`) so Power BI does not ask to reconnect the data source.

---

## Null / blank values in Power BI

Some cells are intentionally blank:

- **R2 / Offer columns** — blank when the candidate did not reach that stage (correct).
- **Avg days by designation/sourcer** — blank when not enough valid date data for that group.

Sheets with **no nulls** (safe for direct visuals):  
`PBI_Summary`, `PBI_Funnel`, `PBI_Bottleneck_Rounds`, `PBI_Stage_Transitions`, `PBI_Status`, `PBI_Dashboard_Layout`.

In Power BI, blank numeric fields are ignored in averages — this is expected.

---

## Validate data quality

```bash
python3 validate_all.py
```

Expected: **ALL CHECKS PASSED - READY FOR HR REVIEW** (13/13)

---

## Project structure

```
CompanyData/
├── uploads/                          ← Drop ANY new .xlsx here (recommended)
├── Supply Mapping.xlsx               ← Example input (any name works)
├── export_powerbi.py                 ← Main script — run this
├── run.bat                           ← Windows double-click runner
├── analyzer.py                       ← Core analytics + 7-level prediction engine
├── validate_all.py                   ← Quality checks (13 checks)
├── requirements.txt                  ← Python dependencies
├── scratch/verify_predictions.py     ← Test: shows prediction level breakdown
├── Supply_Mapping_PowerBI_Ready.xlsx ← Output for Power BI
├── Recruitment_PowerBI_DataPack.zip  ← CSV export (optional)
└── README.md                         ← This file
```

---

## Quick reference

| Task | Command |
|------|---------|
| Export for Power BI | `python3 export_powerbi.py` |
| Validate quality | `python3 validate_all.py` |
| Test prediction model | `python3 scratch/verify_predictions.py` |
| Refresh charts after new data | Re-run export → Power BI **Refresh** |

---

## Notes

- **R0** = submission date (`Date of submission`). There is no separate "R0" column.
- **R2→Offer days** are not in the source file (no offer release date). R0→Offer Stage uses R2 as the last tracked milestone before offer.
- There is **no HTML dashboard** — all charts are built in **Power BI Desktop** using the exported sheets.
- The **R1 Assumption / Note** column in every sheet records exactly which prediction level (L1–L7) was used for each candidate.

---

## Support checklist

- [ ] Python installed and `pip install -r requirements.txt` done
- [ ] `export_powerbi.py` runs without errors
- [ ] `Supply_Mapping_PowerBI_Ready.xlsx` opens and has all `PBI_*` sheets
- [ ] `validate_all.py` outputs 13/13 PASS
- [ ] Power BI report saved as `.pbix`
- [ ] Test: replace input Excel → re-run export → **Refresh** in Power BI → numbers change

**All features are implemented and working.** New uploads only require re-running the export and clicking Refresh in Power BI — all charts and graphs update automatically.
