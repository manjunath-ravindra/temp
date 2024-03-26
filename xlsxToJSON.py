import pandas as pd
import json

def xlsx_to_json(xlsx_file,json_file):
        # Read the Excel file into a DataFrame
    df = pd.read_excel(xlsx_file)
    print(df)

    # Convert the DataFrame to a list of dictionaries
    data = df.to_dict(orient='records')

    # Write the data to a JSON file
    with open(json_file, 'w') as jsonfile:
        jsonfile.write(json.dumps(data, indent=4))

if __name__ == "__main__":
    xlsx_file = ".\\Input\\temp.xlsx"  # Change this to your XLSX file
    json_file = '.\\Output\\output.json'  # Change this to your desired JSON file

    xlsx_to_json(xlsx_file,json_file)