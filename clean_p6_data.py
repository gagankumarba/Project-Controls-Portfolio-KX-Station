import pandas as pd
import os

# 1. Setup File Paths
script_dir = os.path.dirname(os.path.abspath(__file__))
# Verify the filename matches your P6 export file exactly
input_file = os.path.join(script_dir, 'KX-P8-Activities.xlsx') 
output_file = os.path.join(script_dir, 'Cleaned_Project_Data.csv')

print(f"Processing file: {input_file}")

try:
    # 2. Load the Excel File
    # skiprows=1 skips the 'User: Admin' row
    df = pd.read_excel(input_file, skiprows=1)

    # Clean whitespace from headers just in case
    df.columns = df.columns.str.strip()

    print("Found these columns in your Excel file:")
    print(df.columns.tolist()) 

    # 3. Rename Columns (Updated to match your specific P6 Export)
    df.rename(columns={
        'Activity ID': 'ActivityID',
        'Activity Name': 'ActivityName',
        'Activity Status': 'Status',
        'WBS Code': 'WBS',
        '(*)Start': 'Start_Date',
        '(*)Finish': 'Finish_Date',
        '(*)BL1 Start': 'Baseline_Start',
        '(*)BL1 Finish': 'Baseline_Finish',
        '(*)Budgeted Total Cost(£)': 'Budget',
        '(*)Planned Value Cost(£)': 'PV',
        '(*)Earned Value Cost(£)': 'EV',
        '(*)Actual Total Cost(£)': 'AC',
        'Activity % Complete(%)': 'Percent_Complete'
    }, inplace=True)

    # --- CHECK IF RENAME WORKED ---
    if 'PV' not in df.columns:
        print("\nCRITICAL ERROR: Column mapping failed.")
        print("Python could not find '(*)Planned Value Cost(£)' to rename.")
        print("Please check the 'Found these columns' list above again.")
        raise KeyError("Column mapping failed.")

    # 4. Fill Missing Values
    df.fillna({
        'PV': 0, 'EV': 0, 'AC': 0, 'Budget': 0
    }, inplace=True)

    # 5. Calculate SPI and CPI
    df['SPI'] = df.apply(lambda x: round(x['EV'] / x['PV'], 2) if x['PV'] > 0 else 1.0, axis=1)
    df['CPI'] = df.apply(lambda x: round(x['EV'] / x['AC'], 2) if x['AC'] > 0 else 1.0, axis=1)

    # 6. Calculate Variance (in Days)
    # Convert date columns to datetime objects
    df['Finish_Date'] = pd.to_datetime(df['Finish_Date'])
    df['Baseline_Finish'] = pd.to_datetime(df['Baseline_Finish'])
    
    # Calculate Variance: (Baseline - Finish). Negative number means delay.
    df['Finish_Variance_Days'] = (df['Baseline_Finish'] - df['Finish_Date']).dt.days

    # 7. Save to CSV
    df.to_csv(output_file, index=False)

    print("\nSUCCESS! Data cleaned and saved to 'Cleaned_Project_Data.csv'")
    print(f"Processed {len(df)} activities.")

except FileNotFoundError:
    print(f"ERROR: Could not find the file '{input_file}'. Check the filename.")
except Exception as e:
    print(f"\nERROR: {e}")
    