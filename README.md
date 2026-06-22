# Recruitment Pipeline Analyzer — Power BI Edition

Analyze **Supply Mapping** Excel data: stage timing (R0→R1→R2→Offer), bottlenecks by **designation**, **round**, and **sourcer**, plus 30+ day slow-mover flags. Output is ready for **Power BI** — no HTML website required.

---

## Final confirmation — all features implemented

| Feature | Status | Power BI sheet |
|--------|--------|----------------|
| Days R0→R1, R1→R2, R0→Offer by **Designation** | Done | `PBI_Bottleneck_Designation`, `PBI_Designation_Timing` |
| Same timing by **Sourcer** | Done | `PBI_Bottleneck_Sourcer`, `PBI_Sourcer_Timing` |
| Designations with **30+ day** delays | Done | `Count R0→R1 30+ Days`, `Any Stage 30+ Flag` |
| Sourcers with **30+ day** delays | Done | Same columns in sourcer sheets |
| **Bottleneck #1** — Designation | Done | `PBI_Bottleneck_Designation` |
| **Bottleneck #2** — Round (R0→R1 vs R1→R2 vs R0→Offer) | Done | `PBI_Bottleneck_Rounds` |
| **Bottleneck #3** — Slowest sourcers | Done | `PBI_Bottleneck_Sourcer` |
| Slow mover candidate list (30+ days) | Done | `PBI_Slow_Movers_30Plus` |
| KPI summary cards | Done | `PBI_Summary` |
| Recruitment funnel | Done | `PBI_Funnel` |
| Status breakdown | Done | `PBI_Status` |
| Sourcer bar chart data | Done | `PBI_Sourcer_Chart` |
| Per-candidate timeline | Done | `PBI_Candidates` |
| Stage transitions (long format for charts) | Done | `PBI_Stage_Transitions` |
| Dashboard build guide | Done | `PBI_Dashboard_Layout` |
| Data validation (13 checks) | Done | `Validation Report` |
| Power BI export (Excel + CSV zip) | Done | See output files below |

**Validation:** run `python3 validate_all.py` — all 13 quality checks must pass.

---

## Requirements

- **Python 3.9+**
- **Power BI Desktop** (free from Microsoft)
- Input file: **Supply Mapping.xlsx** (same format as current HR file)

### Install Python packages (one time)

```bash
cd /path/to/CompanyData
python3 -m pip install -r requirements.txt
```

---

## How to run

### Option A — Windows (easiest)

1. Put the new **Supply Mapping.xlsx** in this folder.
2. Double-click **`run.bat`**.
3. Wait for **"POWER BI EXPORT COMPLETE"**.

### Option B — Mac / Linux / terminal

```bash
cd /path/to/CompanyData
python3 export_powerbi.py
```

### Option C — Custom input or output paths

```bash
python3 export_powerbi.py "C:\Data\New_Supply_Mapping.xlsx" \
  --excel "C:\Reports\PowerBI_Ready.xlsx" \
  --zip "C:\Reports\PowerBI_DataPack.zip"
```

### Output files (created every run)

| File | Purpose |
|------|---------|
| `Supply_Mapping_PowerBI_Ready.xlsx` | **Main file** — import this into Power BI |
| `Recruitment_PowerBI_DataPack.zip` | Same data as CSV files (optional) |
| `Supply_Mapping_Analyzed.xlsx` | Created when you run `validate_all.py` |

---

## Power BI integration (first time — one-time setup)

Do this **once**. After that, only **Refresh** is needed when new data arrives.

### Step 1 — Generate data

Run `export_powerbi.py` or `run.bat` so `Supply_Mapping_PowerBI_Ready.xlsx` exists.

### Step 2 — Import into Power BI Desktop

1. Open **Power BI Desktop**.
2. **Home → Get data → Excel**.
3. Select **`Supply_Mapping_PowerBI_Ready.xlsx`**.
4. In Navigator, tick **every sheet that starts with `PBI_`** (all 13 sheets).
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

### Step 3 — Create one relationship

1. Open **Model** view (left sidebar).
2. Drag **`Candidate Key`** from `PBI_Stage_Transitions` onto **`Candidate Key`** in `PBI_Candidates`.
3. Set cardinality: **Many to one**, cross-filter: **Single**.

### Step 4 — Build dashboard pages (follow `PBI_Dashboard_Layout`)

Open the **`PBI_Dashboard_Layout`** sheet in Excel — each row is one visual to add.

#### Page 1 — Executive Overview

| Visual | Table | Fields |
|--------|-------|--------|
| **Card** (repeat) | `PBI_Summary` | Filter `Metric`, show `Value` |
| **Funnel** | `PBI_Funnel` | `Stage`, `Count`, sort by `Stage Order` |
| **Donut chart** | `PBI_Status` | `Status`, `Count` |
| **Clustered bar chart** | `PBI_Sourcer_Chart` | `Sourcer`, `Total Submitted`, `Overall Conversion %` |

Suggested KPI cards from `PBI_Summary`: Total Candidates, Reached R1, Reached R2, Reached Offer, Avg Days R0→R1, Slow Movers 30+ Days.

#### Page 2 — Bottlenecks

| Visual | Table | Fields |
|--------|-------|--------|
| **Bar chart** | `PBI_Bottleneck_Rounds` | `Transition`, `Avg Days` |
| **Table** | `PBI_Bottleneck_Designation` | Designation, Avg Days R0→R1, Count 30+, Primary Bottleneck |
| **Table** | `PBI_Bottleneck_Sourcer` | Sourcer, Avg Days R0→R1, Count 30+, Primary Bottleneck |
| **Clustered column chart** | `PBI_Sourcer_Chart` | `Sourcer`, `Avg Days R0→R1`, `Avg Days R1→R2` |

#### Page 3 — Detail & slow movers

| Visual | Table | Fields |
|--------|-------|--------|
| **Matrix** | `PBI_Stage_Transitions` | Rows: `Designation`, Columns: `Transition`, Values: Avg `Days` |
| **Table** | `PBI_Slow_Movers_30Plus` | Name, Sourcer, Designation, Slowest Transition, Slowest Days |
| **Table** | `PBI_Candidates` | Name, Sourcer, Days R0→R1, Days R1→R2, Any Stage 30+ Days |

### Step 5 — Save the report

**File → Save as** → e.g. `Recruitment_Dashboard.pbix`

You only build charts **once**. New data updates all visuals automatically after Refresh (see below).

---

## New data upload — automatic chart & graph update

When someone uploads **new Supply Mapping data**, follow this workflow. **All existing charts, bars, tables, and KPIs update automatically** — you do not rebuild the dashboard.

```
New Supply Mapping.xlsx
        ↓
  run export (run.bat or export_powerbi.py)
        ↓
  Supply_Mapping_PowerBI_Ready.xlsx  (overwritten with new numbers)
        ↓
  Power BI Desktop → Refresh
        ↓
  All visuals update automatically
```

### Step-by-step for new data

1. **Replace** the old file with the new **Supply Mapping.xlsx** in this project folder (or pass the new path to `export_powerbi.py`).
2. **Run export again:**
   - Windows: double-click **`run.bat`**
   - Mac/terminal: `python3 export_powerbi.py`
3. Open your saved **`Recruitment_Dashboard.pbix`** in Power BI Desktop.
4. Click **Home → Refresh** (or **Transform data → Refresh**).
5. All KPI cards, funnel, bar charts, bottleneck tables, and slow-mover lists **reload from the updated Excel file**.

> **Important:** Keep the **same Excel output path** (`Supply_Mapping_PowerBI_Ready.xlsx`) so Power BI does not ask to reconnect the data source. If you change the path, use **Transform data → Data source settings → Change source** once.

### Optional — scheduled refresh (Power BI Service)

If you publish the report to **Power BI Service** (cloud):

1. Upload `Supply_Mapping_PowerBI_Ready.xlsx` to **OneDrive** or **SharePoint**.
2. In Power BI Service, set the dataset **Scheduled refresh** (e.g. daily).
3. Automate `export_powerbi.py` with **Windows Task Scheduler** or a cron job so the Excel file is regenerated before refresh runs.

---

## Null / blank values in Power BI

Some cells are intentionally blank:

- **R2 / Offer columns** — blank when the candidate did not reach that stage (correct).
- **Avg days by designation/sourcer** — blank when there is not enough valid date data for that group.

Sheets with **no nulls** (safe for direct visuals): `PBI_Summary`, `PBI_Funnel`, `PBI_Bottleneck_Rounds`, `PBI_Stage_Transitions`, `PBI_Status`, `PBI_Dashboard_Layout`.

In Power BI, blank numeric fields are ignored in averages — this is expected.

---

## Validate data quality

Before sharing with HR or your manager:

```bash
python3 validate_all.py
```

Expected result: **ALL CHECKS PASSED - READY FOR HR REVIEW**

---

## Project structure

```
CompanyData/
├── Supply Mapping.xlsx              ← Input (HR upload)
├── export_powerbi.py                ← Main script — run this
├── run.bat                          ← Windows double-click runner
├── analyzer.py                      ← Core analytics engine
├── validate_all.py                  ← Quality checks
├── requirements.txt                 ← Python dependencies
├── Supply_Mapping_PowerBI_Ready.xlsx ← Output for Power BI
├── Recruitment_PowerBI_DataPack.zip  ← CSV export (optional)
└── README.md                        ← This file
```

---

## Quick reference

| Task | Command |
|------|---------|
| Export for Power BI | `python3 export_powerbi.py` |
| Validate quality | `python3 validate_all.py` |
| Refresh charts after new data | Re-run export → Power BI **Refresh** |

---

## Notes

- **R0** = submission date (`Date of submission`). There is no separate “RO” column.
- **R2→Offer days** are not in the source file (no offer release date). **R0→Offer Stage** uses R2 as the last tracked milestone before offer.
- There is **no HTML dashboard** — all charts are built in **Power BI Desktop** using the exported sheets.

---

## Support checklist for your manager

- [ ] Python installed and `pip install -r requirements.txt` done
- [ ] `export_powerbi.py` runs without errors
- [ ] `Supply_Mapping_PowerBI_Ready.xlsx` opens and has all `PBI_*` sheets
- [ ] Power BI report saved as `.pbix`
- [ ] Test: replace input Excel → re-run export → **Refresh** in Power BI → numbers change

**All requested features are implemented and working.** New uploads only require re-running the export and clicking Refresh in Power BI — all charts and graphs update automatically.
