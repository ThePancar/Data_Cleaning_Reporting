import pandas as pd
import glob

# Order Flow Raporunun sadelestirilmesi icin kullaniliyor

directory_path = 'C:/Users/semih/OneDrive/Desktop/Order Flow Reports/Total Reports/'

# List all Excel files in the directory
excel_files = glob.glob(directory_path + "*.xlsx")

combined_df = pd.DataFrame()

# Fixed exchange rates
usd_to_sar_rate = 3.7575
eur_to_sar_rate = 4.0585

# Process each Excel file
for file_path in excel_files:
    try:
        # Load the Excel file and check available sheets
        excel_file = pd.ExcelFile(file_path)
        if "Order Flow Report KASKTAS" in excel_file.sheet_names:
            # Load the Excel file with the correct header row
            df = pd.read_excel(file_path, sheet_name="Order Flow Report KASKTAS", header=1)
            
            # Filter the rows where 'PO Status' is 'APPROVED'
            filtered_df = df[df['PO Status'] == 'APPROVED']
            print(f"After filtering 'PO Status': {filtered_df.shape}")

            # Further filter the rows where 'Received Ratio' is 1 (which represents 100%)
            filtered_df = filtered_df[filtered_df['Received Ratio'] != 1]
            print(f"After filtering 'Received Ratio': {filtered_df.shape}")

            # Further filter the rows where 'PO No' contains 'SO'
            filtered_df = filtered_df[~filtered_df['PO No'].str.contains('SO', na=False)]
            print(f"After filtering 'PO No': {filtered_df.shape}")

            # Delete specific words' Row
            keywords = ['drink', 'Groce', "groce", 'Drink', 'Food','food', 'manpower', 'Manpower', 'Market', 'market']
            filtered_df = filtered_df[~filtered_df['PO Title'].str.contains('|'.join(keywords), case=False, na=False)]
            print(f"After filtering 'PO Title': {filtered_df.shape}")

            # Append the filtered data to the combined DataFrame
            combined_df = pd.concat([combined_df, filtered_df], ignore_index=True)
            print(f"Combined DataFrame shape: {combined_df.shape}")
        else:
            print(f"Sheet 'Order Flow Report KASKTAS' not found in file {file_path}. Available sheets: {excel_file.sheet_names}")

    except Exception as e:
        print(f"An error occurred while processing the file {file_path}: {e}")

# Drop columns from index 1 to 19 (inclusive), 33 to 36 (inclusive), 39 to 55 (inclusive), and 61 to 62 (inclusive) after merging
if not combined_df.empty:
    columns_to_drop = combined_df.columns[list(range(1, 19)) + list(range(33, 36)) + list(range(39, 55)) + list(range(61, 62))]
    combined_df = combined_df.drop(columns=columns_to_drop)
    print(f"Shape after dropping columns: {combined_df.shape}")
else:
    print("Combined DataFrame is empty after filtering. No columns to drop.")

# Convert the Unit Price based on the Currency Code directly in the combined_df
combined_df.loc[combined_df['Currency Code'] == 'USD', 'Unit Price'] *= usd_to_sar_rate
combined_df.loc[combined_df['Currency Code'] == 'EUR', 'Unit Price'] *= eur_to_sar_rate

# Save the combined data to a new Excel file
combined_file_path = directory_path + 'Total Combined Warehouse Report.xlsx'
combined_df.to_excel(combined_file_path, index=False)

print(f"Combined cleaned data saved to {combined_file_path}")
