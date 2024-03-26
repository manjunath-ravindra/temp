import pandas as pd
import json
import os

def row_to_json(row, json_directory):
    # Convert the row to a dictionary
    data = row.to_dict()
    
    # Get the index of the row
    index = row.name
    
    # Create a filename based on the row index
    filename = os.path.join(json_directory, f'row_{index}.json')
    
    # Write the data to a JSON file
    with open(filename, 'w') as jsonfile:
        jsonfile.write(json.dumps(data, indent=4))

def xlsx_to_json(xlsx_file, json_directory):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(xlsx_file)
    print(df)

    # Drop rows where all values are NaN (empty rows)
    df = df.dropna(how='all')

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        # Convert and save each row as a separate JSON file
        row_to_json(row, json_directory)

if __name__ == "__main__":
    xlsx_file = ".\\Input\\temp.xlsx"  # Change this to your XLSX file
    json_directory = '.\\Output\\'  # Directory to store individual JSON files

    # Create the directory if it doesn't exist
    os.makedirs(json_directory, exist_ok=True)

    xlsx_to_json(xlsx_file, json_directory)

