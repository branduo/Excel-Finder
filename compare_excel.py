import os
import pandas as pd

def search_excel_files_and_create_summary(output_path):
    # Define the directory to search for Excel files
    search_directory = input("Enter the directory to search for Excel files: ")

    # Define the output columns
    output_columns = []

    # Initialize an empty DataFrame for consolidated data
    #consolidated_data = pd.DataFrame(columns=output_columns)
    quanta_data = pd.DataFrame(columns=output_columns)
    compal_data = pd.DataFrame(columns=output_columns)

    # Search for Excel files
    for root, dirs, files in os.walk(search_directory):
        for file in files:
            if file.endswith(('.csv')):
                file_path = os.path.join(root, file)
                try:
                    for encoding in ['utf-8', 'ISO-8859-1', 'Windows-1252']:
                        try:
                            df = pd.read_csv(file_path, encoding = encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        raise UnicodeDecodeError("All encodings failed.")
                
                    keywords = ['part', 'description', 'cost', 'price', 'supplier', 'variance', 'spec']
                    output_columns = [col for col in df.columns if any(keyword.casefold() in col.casefold() for keyword in keywords)]

                    relevant_data = df[output_columns]
                    print(f"Selected columns in {file}: {output_columns}")
                    print(relevant_data.head())

                    # Separate data based on ODM and append to respective DataFrames
                    file_name = file.lower()
                    if 'quanta' in file_name:
                        quanta_data = pd.concat([quanta_data, relevant_data], ignore_index=True)
                        print(f"Added data from {file} to 'Quanta' sheet")

                    if 'compal' in file_name:
                        compal_data = pd.concat([compal_data, relevant_data], ignore_index=True)
                        print(f"Added data from {file} to 'Compal' sheet")

                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

            if file.endswith(('.xlsx')):
                file_path = os.path.join(root, file)
                try:
                    # Read the Excel file
                    df = pd.read_excel(file_path, engine='openpyxl')

                    keywords = ['part', 'description', 'cost', 'price', 'supplier', 'variance', 'spec']
                    output_columns = [col for col in df.columns if any(keyword.casefold() in col.casefold() for keyword in keywords)]

                    relevant_data = df[output_columns]
                    print(f"Selected columns in {file}: {output_columns}")
                    print(relevant_data.head())

                    # Separate data based on ODM and append to respective DataFrames
                    file_name = file.lower()
                    if 'quanta' in file_name:
                        quanta_data = pd.concat([quanta_data, relevant_data], ignore_index=True)
                        print(f"Added data from {file} to 'Quanta' sheet")

                    if 'compal' in file_name:
                        compal_data = pd.concat([compal_data, relevant_data], ignore_index=True)
                        print(f"Added data from {file} to 'Compal' sheet")

                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

    with pd.ExcelWriter('consolidated_data.xlsx', engine = 'openpyxl') as writer:
        quanta_data.to_excel(writer, sheet_name = 'Quanta')
        compal_data.to_excel(writer, sheet_name = 'Compal')
    print(f"Consolidated Excel file created at: {output_path}")

# Specify the output Excel file path
output_file_path = "consolidated_data.xlsx"

# Run the function
search_excel_files_and_create_summary(output_file_path)