import pandas as pd

#Haftalik hazirlanan Ambar girislerinin KPI Raporu, 1 hafta icerisinde ne kadarlik mal girisi yapilmis, ne kadar beklemede

# Define the paths to the Excel files
old_report_path = "C:/Users/xxx/xxx/xxx/xxx/xxx/xxx/xxx/xxx/23072024/23072024 Combined Warehouse Report.xlsx"
new_report_path = "C:/Users/xxx/xxx/xxx/xxx/xxx/xxx/xxx/xxx/23072024/29072024 Combined Warehouse Report.xlsx"

# Load the data from the first sheet of both files
df_old = pd.read_excel(old_report_path, sheet_name=0)
df_new = pd.read_excel(new_report_path, sheet_name=0)

# Extract the relevant columns from both reports
columns_to_extract = ['PO No', 'PO Item No', 'PO Title', 'PO Desc', 'PO Approval Date', 'PO Qty', 'Unit Price', 'To be Rcv \nQty']

# Extract the columns from the old and new reports
old_report_columns = df_old[columns_to_extract].rename(columns={'To be Rcv \nQty': 'To be Rcv Qty Old'})
new_report_columns = df_new[columns_to_extract].rename(columns={'To be Rcv \nQty': 'To be Rcv Qty New'})

# Merge both reports based on "PO No" and "PO Item No"
comparison_df = pd.merge(
    old_report_columns,
    new_report_columns,
    on=['PO No', 'PO Item No'],
    how='outer',
    suffixes=('_Old', '_New')
)

# Add a column to calculate the ratio from old to new and check if it's newly added
def calculate_ratio(row):
    old_qty = row['To be Rcv Qty Old']
    new_qty = row['To be Rcv Qty New']
    po_qty = row['PO Qty_New']
    
    if pd.isna(old_qty) and not pd.isna(new_qty):
        if pd.isna(po_qty) or po_qty == 0:
            return '0.00% done', 'Newly Added'
        return f"{((po_qty - new_qty) / po_qty) * 100:.2f}% done", 'Newly Added'
    elif pd.isna(new_qty) or new_qty == 0:
        return '100% done', ''
    else:
        return f"{(1 - new_qty / old_qty) * 100:.2f}% done", ''

comparison_df[['Completion Ratio', 'New Item Status']] = comparison_df.apply(
    lambda row: pd.Series(calculate_ratio(row)), axis=1
)

# Add a column to indicate if the "To be Rcv Qty" values are the same or different
def determine_status(row):
    if row['New Item Status'] == 'Newly Added':
        return ''
    return 'Unchanged' if row['To be Rcv Qty Old'] == row['To be Rcv Qty New'] else 'Changed'

comparison_df['Status'] = comparison_df.apply(determine_status, axis=1)

# Ensure the relevant details are filled for new items and blank for old details
for column in ['PO Title', 'PO Desc', 'PO Approval Date', 'PO Qty', 'Unit Price']:
    comparison_df[column] = comparison_df.apply(
        lambda row: row[f'{column}_New'] if pd.isna(row[f'{column}_Old']) else row[f'{column}_Old'], axis=1
    )

# Ensure 'To be Rcv Qty Old' is blank for new items
comparison_df['To be Rcv Qty Old'] = comparison_df.apply(
    lambda row: '' if row['New Item Status'] == 'Newly Added' else row['To be Rcv Qty Old'], axis=1
)

# Add a new column "Project Code" that extracts the first 5 digits from "PO No"
comparison_df['Project Code'] = comparison_df['PO No'].str[:5]

# Reorder columns to place "Project Code" at the beginning
comparison_df = comparison_df[['Project Code', 'PO No', 'PO Item No', 'Unit Price', 'PO Title', 'PO Desc', 'PO Approval Date', 'PO Qty', 'To be Rcv Qty Old', 'To be Rcv Qty New', 'Status', 'Completion Ratio', 'New Item Status']]

# Save the comparison result to a new Excel file
comparison_output_path = "C:/Users/semih/OneDrive/Desktop/Semi Klasor/VS Codes/Python/Kasktas Arabia Works/Order Flow Reports/Comparison_Report.xlsx"

# Convert completion ratio to numeric values for aggregation
comparison_df['Completion Ratio Numeric'] = comparison_df['Completion Ratio'].str.rstrip('% done').astype(float)

# Calculate the average completion ratio for each project
summary_df = comparison_df.groupby('Project Code')['Completion Ratio Numeric'].mean().reset_index()
summary_df.rename(columns={'Completion Ratio Numeric': 'Average Completion Ratio (%)'}, inplace=True)

# Write both comparison and summary to Excel
with pd.ExcelWriter(comparison_output_path) as writer:
    comparison_df.to_excel(writer, sheet_name='Comparison Report', index=False)
    summary_df.to_excel(writer, sheet_name='Summary Report', index=False)

# Output the differences to the console
print(comparison_df)
print("\nSummary Report:")
print(summary_df)

