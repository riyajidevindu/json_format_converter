import pandas as pd
import json

def convert_excel_to_json(excel_file_path, output_json_path):
    """
    Converts an Excel file with multiple sheets to a JSON file in the specified format.

    Args:
        excel_file_path (str): The path to the input Excel file.
        output_json_path (str): The path to the output JSON file.
    """
    all_data = []
    xls = pd.ExcelFile(excel_file_path)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Check for required columns
        required_columns = ['Original', 'Need Anonymization', 'Anonymized', 'PII Identifires', 'PII Value']
        if not all(col in df.columns for col in required_columns):
            print(f"Sheet '{sheet_name}' is missing one or more required columns. Skipping.")
            continue

        for index, row in df.iterrows():
            # Handle 'Need Anonymization' column.
            need_anonymization_raw = str(row['Need Anonymization']).strip().lower()
            is_anonymization_needed = need_anonymization_raw in ['yes', 'true', '1']
            need_anonymization_output = 'yes' if is_anonymization_needed else 'no'

            # Ensure 'Original' column is not empty
            original_input = row['Original']
            if pd.isna(original_input):
                print(f"Skipping row {index+2} in sheet '{sheet_name}' due to empty 'Original' value.")
                continue

            if is_anonymization_needed:
                anonymized_text = row['Anonymized']
                # If 'Anonymized' is empty but required, treat it as an empty string.
                if pd.isna(anonymized_text):
                    anonymized_text = ""

                pii_identifiers = row['PII Identifires']
                if pd.isna(pii_identifiers):
                    pii_identifiers_list = []
                else:
                    pii_identifiers_list = [item.strip() for item in str(pii_identifiers).split(',')]

                pii_values = row['PII Value']
                if pd.isna(pii_values):
                    pii_values_list = []
                else:
                    pii_values_list = [item.strip() for item in str(pii_values).split(',')]
            else:
                # If anonymization is not needed, the anonymized text is the original text,
                # and PII fields are empty.
                anonymized_text = original_input
                pii_identifiers_list = []
                pii_values_list = []

            json_record = {
                "instruction": "Detect and anonymize PII in the user input while preserving intent. Return the anonymized prompt along with identified PII.",
                "input": original_input,
                "output": {
                    "need_anonymization": need_anonymization_output,
                    "anonymized": anonymized_text,
                    "pii_identifiers": pii_identifiers_list,
                    "pii_values": pii_values_list
                }
            }
            all_data.append(json_record)

    with open(output_json_path, 'w') as f:
        json.dump(all_data, f, indent=2)

if __name__ == "__main__":
    # You can change the file names here.
    # The script will look for this file in the same folder.
    excel_file = 'Dataset.xlsx'
    output_json = 'output.json'
    
    try:
        convert_excel_to_json(excel_file, output_json)
        print(f"Successfully converted {excel_file} to {output_json}")
    except FileNotFoundError:
        print(f"Error: The file '{excel_file}' was not found. Please make sure the file is in the same directory as the script, or provide the full path.")
    except Exception as e:
        print(f"An error occurred: {e}")
