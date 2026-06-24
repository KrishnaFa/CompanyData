"""
R1 Date Prediction System
Predicts missing R1 dates based on designation, skills, and other factors
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

def load_data(file_path):
    """Load the Excel file"""
    print(f"Loading data from {file_path}...")
    df = pd.read_excel(file_path)
    print(f"Loaded {len(df)} rows with {len(df.columns)} columns")
    return df

def clean_data(df):
    """Remove special characters like # from text columns"""
    print("\nCleaning data...")
    
    # Columns to skip (date columns)
    skip_columns = ['R1 Date', 'Date of submission', 'R2 Date', 'DOJ']
    
    # Get text columns (object dtype) excluding date columns
    text_columns = [col for col in df.select_dtypes(include=['object']).columns if col not in skip_columns]
    
    # Remove special characters from text columns
    for col in text_columns:
        df[col] = df[col].astype(str).str.replace('#', '', regex=False)
        df[col] = df[col].astype(str).str.replace('*', '', regex=False)
    
    print(f"  - Cleaned {len(text_columns)} text columns")
    
    return df

def preprocess_dates(df, date_column='R1 Date', submission_column='Date of submission', max_date=None):
    """Convert date columns to datetime and handle various formats"""
    print(f"\nPreprocessing date columns...")
    
    if max_date is None:
        max_date = datetime.now()
    
    # Convert R1 Date to datetime
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce', dayfirst=True)
    
    # Convert submission date to datetime
    df[submission_column] = pd.to_datetime(df[submission_column], errors='coerce', dayfirst=True)
    
    # Cap known R1 dates to be within 30 days of submission (business-friendly)
    valid_mask = df[date_column].notnull() & df[submission_column].notnull()
    for idx in df[valid_mask].index:
        submission = df.loc[idx, submission_column]
        r1 = df.loc[idx, date_column]
        days_gap = (r1 - submission).days
        if days_gap > 30:
            df.loc[idx, date_column] = submission + timedelta(days=30)
    
    # Cap known dates to max_date (today) to ensure no future dates
    df[date_column] = df[date_column].apply(lambda x: min(x, max_date) if pd.notna(x) else x)
    
    # Count missing and non-missing
    missing_count = df[date_column].isnull().sum()
    non_missing_count = df[date_column].notnull().sum()
    
    print(f"  - R1 Date non-null: {non_missing_count}")
    print(f"  - R1 Date null: {missing_count}")
    print(f"  - Dates capped to: {max_date.strftime('%d-%m-%Y')}")
    
    return df

def build_prediction_model(df):
    """
    Build a prediction model based on submission date and designation
    Calculates the typical time gap between submission and R1
    Caps gaps to maximum 30 days for business-friendly predictions
    """
    print("\nBuilding prediction model...")
    
    # Get rows with known R1 dates and submission dates
    known = df[df['R1 Date'].notnull() & df['Date of submission'].notnull()].copy()
    
    if len(known) == 0:
        print("  ERROR: No known R1 dates with submission dates found!")
        return None, None, None
    
    print(f"  - Training samples: {len(known)}")
    
    # Calculate time gap in days between submission and R1
    known['days_gap'] = (known['R1 Date'] - known['Date of submission']).dt.days
    
    # Cap gaps to maximum 30 days (reasonable interview timeline)
    known['days_gap_capped'] = known['days_gap'].clip(upper=30)
    
    # Calculate median days gap per designation (capped)
    designation_gaps = known.groupby('Current Designation')['days_gap_capped'].median()
    print(f"  - Unique designations with R1 dates: {len(designation_gaps)}")
    
    # Calculate overall median days gap as fallback (capped)
    overall_gap = known['days_gap_capped'].median()
    print(f"  - Overall median days gap (capped to 30 days): {overall_gap:.0f} days")
    
    return designation_gaps, overall_gap, known

def predict_missing_dates(df, designation_gaps, overall_gap, max_date=None):
    """
    Predict missing R1 dates using submission date + days gap
    Ensures no predicted date exceeds max_date (today by default)
    """
    print("\nPredicting missing R1 dates...")
    
    if max_date is None:
        max_date = datetime.now()
    print(f"  - Maximum allowed date: {max_date.strftime('%d-%m-%Y')}")
    
    # Get rows with missing R1 dates but valid submission dates that are not in the future
    missing_mask = df['R1 Date'].isnull() & df['Date of submission'].notnull() & (df['Date of submission'] <= max_date)
    missing_count = missing_mask.sum()
    
    future_submission_mask = df['R1 Date'].isnull() & df['Date of submission'].notnull() & (df['Date of submission'] > max_date)
    future_count = future_submission_mask.sum()
    if future_count > 0:
        print(f"  - Skipping {future_count} rows with future submission dates (> max_date)")
    
    if missing_count == 0:
        print("  - No missing dates to predict!")
        return df
    
    print(f"  - Predicting for {missing_count} rows...")
    
    # Predict based on submission date + designation gap
    predicted_dates = []
    for idx, row in df[missing_mask].iterrows():
        designation = row['Current Designation']
        submission_date = row['Date of submission']
        
        # Get days gap for this designation
        if pd.notna(designation) and designation in designation_gaps:
            days_gap = designation_gaps[designation]
        else:
            days_gap = overall_gap
        
        # Calculate R1 date as submission date + days gap
        predicted_date = submission_date + timedelta(days=int(days_gap))
        
        # If predicted date exceeds max_date, adjust the gap to fit within limit
        if predicted_date > max_date:
            max_allowed_gap = (max_date - submission_date).days
            if max_allowed_gap > 0:
                predicted_date = submission_date + timedelta(days=max_allowed_gap)
            else:
                # If submission is already at or past max_date, use max_date - 1 day
                predicted_date = max_date - timedelta(days=1)
        
        # Ensure R1 date is at least 1 day after submission date
        min_r1_date = submission_date + timedelta(days=1)
        if predicted_date < min_r1_date:
            predicted_date = min_r1_date
        
        # Final cap to ensure date doesn't exceed max_date
        if predicted_date > max_date:
            predicted_date = max_date
        
        predicted_dates.append(predicted_date)
    
    # Fill missing dates
    df.loc[missing_mask, 'R1 Date'] = predicted_dates
    
    print(f"  - Successfully predicted {len(predicted_dates)} dates")
    
    return df

def format_dates(df, date_column='R1 Date', output_format='%d-%m-%Y'):
    """Format dates to specified format (dd-mm-yyyy)"""
    print(f"\nFormatting dates to {output_format}...")
    
    # Convert to string in desired format
    df[date_column] = df[date_column].dt.strftime(output_format)
    
    print(f"  - Dates formatted successfully")
    
    return df

def save_output(df, output_path):
    """Save the filled data to Excel"""
    print(f"\nSaving output to {output_path}...")
    df.to_excel(output_path, index=False)
    print(f"  - File saved successfully")

def main():
    # Configuration
    input_file = '/Users/apple/Downloads/Kamal/CompanyData/Supply Mapping.xlsx'
    output_file = '/Users/apple/Downloads/Kamal/CompanyData/Supply_Mapping_R1_Filled.xlsx'
    date_column = 'R1 Date'
    
    # Maximum allowed date (today, date only without time)
    max_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
    print("="*60)
    print("R1 DATE PREDICTION SYSTEM")
    print("="*60)
    
    # Load data
    df = load_data(input_file)
    
    # Clean data (remove special characters)
    df = clean_data(df)
    
    # Preprocess dates (cap to today)
    df = preprocess_dates(df, date_column=date_column, submission_column='Date of submission', max_date=max_date)
    
    # Build prediction model
    designation_gaps, overall_gap, known_data = build_prediction_model(df)
    
    if designation_gaps is None:
        print("\nERROR: Cannot build prediction model. Exiting.")
        return
    
    # Predict missing dates
    df = predict_missing_dates(df, designation_gaps, overall_gap, max_date)
    
    # Format dates to dd-mm-yyyy
    df = format_dates(df, date_column)
    
    # Save output
    save_output(df, output_file)
    
    print("\n" + "="*60)
    print("PREDICTION COMPLETE")
    print("="*60)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Total rows: {len(df)}")
    print(f"R1 Date column filled successfully")
    print("="*60)

if __name__ == "__main__":
    main()
