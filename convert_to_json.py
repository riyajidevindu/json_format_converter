import pandas as pd
import json
import csv
import io

def convert_excel_to_json(excel_file_path, output_json_path):
    """
    Converts an Excel file with multiple sheets to a JSON file in the specified format.

    Args:
        excel_file_path (str): The path to the input Excel file.
        output_json_path (str): The path to the output JSON file.
    """
    anonymized_data = []
    non_anonymized_data = []
    total_converted_rows = 0
    xls = pd.ExcelFile(excel_file_path)

    print("Starting conversion...")
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Check for required columns
        required_columns = ['Original', 'Need Anonymization', 'Anonymized', 'PII Identifiers', 'PII Value']
        if not all(col in df.columns for col in required_columns):
            print(f"Sheet '{sheet_name}' is missing one or more required columns. Skipping.")
            continue

        sheet_converted_rows = 0
        for index, row in df.iterrows():
            # Handle 'Need Anonymization' column.
            need_anonymization_raw = str(row['Need Anonymization']).strip().lower()
            is_anonymization_needed = need_anonymization_raw in ['yes', 'true', '1']
            need_anonymization_output = 'yes' if is_anonymization_needed else 'no'

            # Ensure 'Original' column is not empty
            original_input = row['Original']
            if pd.isna(original_input):
                # Skipping rows with no original text.
                continue

            if is_anonymization_needed:
                anonymized_text = row['Anonymized']
                # If 'Anonymized' is empty but required, treat it as an empty string.
                if pd.isna(anonymized_text):
                    anonymized_text = ""

                pii_identifiers_list = []
                if not pd.isna(row['PII Identifiers']):
                    try:
                        # Attempt to parse the string as a JSON array
                        loaded_identifiers = json.loads(str(row['PII Identifiers']))
                        if isinstance(loaded_identifiers, list):
                            pii_identifiers_list = loaded_identifiers
                        else:
                            # If the loaded data is not a list (e.g., a single number or string), wrap it in a list.
                            pii_identifiers_list = [loaded_identifiers]
                    except json.JSONDecodeError:
                        # If it's not a valid JSON array, it might be a simple comma-separated string.
                        # We will use the CSV reader as a fallback.
                        reader = csv.reader(io.StringIO(str(row['PII Identifiers'])))
                        pii_identifiers_list = [item.strip() for item in next(reader)]

                pii_values_list = []
                if not pd.isna(row['PII Value']):
                    try:
                        # Attempt to parse the string as a JSON array
                        loaded_values = json.loads(str(row['PII Value']))
                        if isinstance(loaded_values, list):
                            pii_values_list = loaded_values
                        else:
                            # If the loaded data is not a list (e.g., a single number or string), wrap it in a list.
                            pii_values_list = [loaded_values]
                    except json.JSONDecodeError:
                        # Fallback to CSV reader for simple comma-separated values.
                        reader = csv.reader(io.StringIO(str(row['PII Value'])))
                        pii_values_list = [item.strip() for item in next(reader)]

                # --- Data Validation Check ---
                if len(pii_identifiers_list) != len(pii_values_list):
                    print(f"\n--- WARNING: Mismatch Found ---")
                    print(f"Sheet: '{sheet_name}', Excel Row: {index + 2}")
                    print(f"  Original Input: {original_input[:100]}...")
                    print(f"  PII Identifiers ({len(pii_identifiers_list)}): {pii_identifiers_list}")
                    print(f"  PII Values ({len(pii_values_list)}): {pii_values_list}")
                    print(f"  SUGGESTION: Check this row in your Excel file. The number of identifiers and values do not match.")
                    print(f"---------------------------------\n")
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
            if is_anonymization_needed:
                anonymized_data.append(json_record)
            else:
                non_anonymized_data.append(json_record)
            sheet_converted_rows += 1
        
        print(f"- Sheet '{sheet_name}': Converted {sheet_converted_rows} rows.")
        total_converted_rows += sheet_converted_rows

    all_data = anonymized_data + non_anonymized_data
    with open(output_json_path, 'w') as f:
        json.dump(all_data, f, indent=2)
    
    return total_converted_rows

if __name__ == "__main__":
    # You can change the file names here.
    # The script will look for this file in the same folder.
    excel_file = 'Dataset.xlsx'
    output_json = 'output.json'
    
    try:
        total_rows = convert_excel_to_json(excel_file, output_json)
        print(f"\nSuccessfully converted a total of {total_rows} rows from '{excel_file}' to '{output_json}'")
    except FileNotFoundError:
        print(f"Error: The file '{excel_file}' was not found. Please make sure the file is in the same directory as the script, or provide the full path.")
    except Exception as e:
        print(f"An error occurred: {e}")
