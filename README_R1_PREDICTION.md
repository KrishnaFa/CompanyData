# R1 Date Prediction System

## Overview
This Python pipeline predicts missing R1 interview dates based on designation and other factors from the supply mapping data. It uses a rule-based approach where the median R1 date for each designation is used to fill missing values.

## Features
- **Smart Prediction**: Uses designation-specific median dates for accurate predictions
- **Date Validation**: Ensures no predicted date exceeds today's date
- **Standard Format**: All dates formatted to DD-MM-YYYY for consistency
- **Power BI Ready**: Output file can be directly imported into Power BI for visualization

## Installation

Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the prediction script:
```bash
python3 predict_r1_dates.py
```

## How It Works

1. **Data Loading**: Reads the supply mapping Excel file
2. **Data Cleaning**: Removes special characters (#, *) from text columns
3. **Preprocessing**: Converts R1 Date and Date of submission columns to datetime format
4. **Model Building**: Calculates median time gap (in days) between submission and R1 for each designation based on known dates
5. **Prediction**: Fills missing R1 dates using submission date + designation-specific days gap (falls back to overall median gap if designation not found)
6. **Validation**: 
   - Ensures R1 date is always after submission date (minimum 1 day gap)
   - Ensures all predicted dates are on or before today (24 Jun 2026)
   - Skips prediction for rows with future submission dates
7. **Formatting**: Converts all dates to DD-MM-YYYY format
8. **Output**: Saves the filled data to a new Excel file

## Input/Output

**Input**: `Supply Mapping.xlsx` (1331 rows, 37 columns)
- R1 Date column: 47 known dates, 1284 missing dates
- Date of submission: Used as base for prediction

**Output**: `Supply_Mapping_R1_Filled.xlsx`
- 1235 rows with R1 Date filled (1188 predicted + 47 original)
- 96 rows remain unfilled (future submission dates that cannot be predicted)
- All dates in DD-MM-YYYY format
- All R1 dates are after submission date
- All dates on or before 24 Jun 2026
- Ready for Power BI import

## Prediction Logic

The system uses a submission-based approach with business-friendly constraints:
- For each designation with known R1 dates, calculate the median time gap (days) between submission and R1
- **All gaps are capped to maximum 30 days** for realistic interview timelines
- When predicting missing dates: R1 Date = Submission Date + Designation's median days gap (capped to 30 days)
- If designation has no known dates, use the overall median gap (9 days)
- Ensures R1 date is always after submission date (minimum 1 day)
- Caps all predictions at today's date (24 Jun 2026) to ensure validity
- Skips rows with future submission dates (cannot predict past dates for future submissions)

## Power BI Integration

1. Open Power BI Desktop
2. Click "Get Data" → "Excel"
3. Select `Supply_Mapping_R1_Filled.xlsx`
4. Use the R1 Date field for visualizations and analysis

## Files

- `predict_r1_dates.py` - Main prediction script
- `requirements.txt` - Python dependencies
- `Supply_Mapping_R1_Filled.xlsx` - Output file with filled R1 dates

## Statistics

- **Total Rows**: 1,331
- **Original R1 Dates**: 47 (3.5%)
- **Predicted R1 Dates**: 1,188 (89.2%)
- **Unfilled R1 Dates**: 96 (7.2%) - future submission dates
- **Total Filled**: 1,235 (92.8%)
- **Unique Designations**: 34 with known R1 dates
- **Overall Median Gap**: 9 days between submission and R1
- **Max Gap**: 30 days (capped for business-friendly predictions)
- **Date Format**: DD-MM-YYYY
- **Max Date**: 24 Jun 2026
