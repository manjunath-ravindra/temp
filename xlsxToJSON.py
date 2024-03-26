import pandas as pd
import json
import os

def row_to_json(row, json_directory):
    # Extract values from columns C to MX
    values = row.iloc[2:362]  # Exclude the first two columns (assuming 0-based indexing)

    # Convert non-null numeric values to a list
    values_list = [float(value) for value in values if pd.notnull(value) and isinstance(value, (int, float))]

    # Create dictionary with 'radius' key and values list
    data = {'radius': values_list}
    
    # Get the index of the row
    index = row.name
    
    # Create a filename based on the row index
    filename = os.path.join(json_directory, f'row_{index}.json')
    
    # Write the data to a JSON file
    with open(filename, 'w') as jsonfile:
        jsonfile.write(json.dumps(data, indent=4))

def xlsx_to_json(xlsx_file, json_directory):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(xlsx_file, header=None)  # Avoids treating the first row as header
    print(df)

    # Drop rows where all values are NaN (empty rows)
    df = df.dropna(how='all')

    # Iterate over each row in the DataFrame, excluding the first row
    for index, row in df.iloc[1:].iterrows():
        # Convert and save each row as a separate JSON file
        row_to_json(row, json_directory)

if __name__ == "__main__":
    xlsx_file = ".\\Input\\temp.xlsx"  # Change this to your XLSX file
    json_directory = '.\\Output\\json_files'  # Directory to store individual JSON files

    # Create the directory if it doesn't exist
    os.makedirs(json_directory, exist_ok=True)

    xlsx_to_json(xlsx_file, json_directory)
