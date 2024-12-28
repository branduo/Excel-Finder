import pandas as pd
import glob

# Define categories
desired_columns = ["part", "description", "cost", "odm/site", "supplier", "varianceofcost", "specifications existing (yes/no)"]

# Path to folder containing Excel files
folder_path = "./Downloads/"
excel_files = glob.glob(f"{folder_path}/*.csv")
print(f"Excel files found: {excel_files}")

# List to hold DataFrames
dataframes = []

# Loop through each Excel file
for file in excel_files:
    for file in excel_files:
        print(f"Reading file: {file}")
        all_sheets = pd.read_excel(file, sheet_name=None)
        if not all_sheets:
            print(f"No sheets found in file: {file}")

    print(f"Processing file: {file}")
    # Read all sheets in the file
    all_sheets = pd.read_excel(file, sheet_name=None)
    
    for sheet_name, df in all_sheets.items():
        print(f"  Processing sheet: {sheet_name}")
        
        # Standardize column names (example mapping)
        column_mapping = {
            "Part Number": "part",
            "Item Description": "description",
            "Unit Cost": "cost",
            "ODM/Location": "odm/site",
            "Supplier Name": "supplier",
            "Cost Variance": "varianceofcost",
            "Specs Existing": "specifications existing (yes/no)"
        }
        
        # Rename columns based on mapping
        df = df.rename(columns=column_mapping)
        
        # Keep only desired columns
        df = df[[col for col in desired_columns if col in df.columns]]
        
        # Normalize 'specifications existing (yes/no)'
        df["specifications existing (yes/no)"] = df["specifications existing (yes/no)"].str.strip().str.lower().map({"yes": "Yes", "no": "No"})
        
        # Add source file and sheet for traceability
        df["Source File"] = file
        df["Sheet Name"] = sheet_name
        
        # Append to list of DataFrames
        dataframes.append(df)

# Combine all DataFrames into one
#combined_df = pd.concat(dataframes, ignore_index=True)

# Save to an Excel file
output_file = "comparison_results.xlsx"
combined_df.to_excel(output_file, index=False)

print(f"Comparison results saved to '{output_file}'.")