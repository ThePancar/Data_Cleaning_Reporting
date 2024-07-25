import pandas as pd
import glob

# Directory containing the Excel files
directory_path = 'C:/Users/semih/OneDrive/Desktop/Semi Klasor/VS Codes/Python/Kasktas Arabia Works/Order Flow Reports/'

# List all Excel files in the directory
excel_files = glob.glob(directory_path + "*.xlsx")

# Initialize an empty DataFrame to hold the combined data
combined_df = pd.DataFrame()

# Process each Excel file
for file_path in excel_files:
    # Load the Excel file with the correct header row
    df = pd.read_excel(file_path, sheet_name="Order Flow Report KASKTAS", header=1)
    
    # Filter the rows where 'PO status' is 'APPROVED'
    filtered_df = df[df['PO Status'] == 'APPROVED']
    
    # Further filter the rows where 'Received Ratio' is 1 (which represents 100%)
    filtered_df = filtered_df[filtered_df['Received Ratio'] != 1]
    
    # Further filter the rows where 'PO No' contains 'SO'
    filtered_df = filtered_df[~filtered_df['PO No'].str.contains('SO', na=False)]
    
    # Further filter the rows where 'PO Approval Date' is before 2024
    filtered_df['PO Approval Date'] = pd.to_datetime(filtered_df['PO Approval Date'], errors='coerce')
    filtered_df = filtered_df[filtered_df['PO Approval Date'].dt.year >= 2024]
    
    # Append the filtered data to the combined DataFrame
    combined_df = pd.concat([combined_df, filtered_df], ignore_index=True)

# Drop columns from index 1 to 19 (inclusive), 33 to 36 (inclusive), 39 to 55 (inclusive), and 61 to 62 (inclusive) after merging
columns_to_drop = combined_df.columns[list(range(1, 19)) + list(range(33, 36)) + list(range(39, 55)) + list(range(61, 62))]
combined_df = combined_df.drop(columns=columns_to_drop)

# Save the combined data to a new Excel file
combined_file_path = directory_path + 'Combined_Order_Flow_Report_KASKTAS_Cleaned.xlsx'
combined_df.to_excel(combined_file_path, index=False)

print(f"Combined cleaned data saved to {combined_file_path}")
