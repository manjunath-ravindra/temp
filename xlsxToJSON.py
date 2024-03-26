import pandas as pd
import json
import os

def row_to_json(row, json_directory, filename):
    # Extract values from columns C to MX
    values = row.iloc[2:362]  # Exclude the first two columns (assuming 0-based indexing)

    # Convert non-null numeric values to a list
    values_list = [float(value) for value in values if pd.notnull(value) and isinstance(value, (int, float))]

    # Create dictionary with 'radius' key and values list
    data = {"fileInfo":{
        "Ver":"1.0"
        },
    "projectInfo":{
        "project_num":"11",
        "project_name":"abc",
        "customer_name":"Reliance",
        "project_manager":"Robin"
        },
    "pipeinfo":{
        "group_id":"SC20L",
        "pipe_name":"0021",
        "pipe_end":"EndB",
        "pipe_length":"20",
        "pipe_OD":"508",
        "material":"Glass",
        "thickness":"12",
        "type":"Single",
        "sample_size":"360",
        "location":{
            "site_name":"Some yard",
            "latitude":"100",
            "longitude":"100"
            },
        "timestamp":"1702903141",
        "technician_name":"Fake Tech"
        },
        'radius': values_list}
    
    # Write the data to a JSON file
    with open(filename, 'w') as jsonfile:
        jsonfile.write(json.dumps(data, indent=4))

def xlsx_to_json(xlsx_file, json_directory):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(xlsx_file, header=None)  # Avoids treating the first row as header
    # print(df)

    # Drop rows where all values are NaN (empty rows)
    df = df.dropna(how='all')

    # Iterate over each row in the DataFrame, excluding the first row
    for index, row in df.iloc[1:].iterrows():
        # Generate custom filename based on the row index
        pipe_ids = df.iloc[1:, 0].tolist()
        pipe_ends = df.iloc[1:, 1].tolist()
        # print(pipe_ends)

        # Iterate through each file name in the list
        for i in range(len(pipe_ids)):
            # Replace forward slashes with underscores and remove spaces
            pipe_ids[i] = pipe_ids[i].replace('/', '_').replace(' ', '')
            pipe_ids[i]= pipe_ids[i]+ "End"+ pipe_ends[i]

        print(pipe_ids)

        filename = os.path.join(json_directory, f'{pipe_ids[index-4]}.json')
        # Convert and save each row as a separate JSON file with custom name
        row_to_json(row, json_directory, filename)

if __name__ == "__main__":
    xlsx_file = ".\\Input\\temp.xlsx"  # Change this to your XLSX file
    json_directory = '.\\Output\\json_files'  # Directory to store individual JSON files

    # Create the directory if it doesn't exist
    os.makedirs(json_directory, exist_ok=True)

    xlsx_to_json(xlsx_file, json_directory)
