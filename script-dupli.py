import requests
from dotenv import load_dotenv
import os
import googlemaps
import pandas as pd


load_dotenv()

api_key = os.getenv('API_KEY')

def get_school():
  df = pd.read_excel("gmandal.xlsx")
  schools = df["SCHOOL NAME"].tolist()
  return schools

# Function to insert data into Excel sheet
def insert_data_into_excel(iterable_object, excel_file, column_name, column_number):
    try:
        # Read the existing Excel file into a DataFrame
        df = pd.read_excel(excel_file)
        
        # Check if the specified column name already exists, if not, create it
        if column_name not in df.columns:
            df[column_name] = None
        
        # Iterate through the iterable object and insert data into the specified column
        for i, data in enumerate(iterable_object):
            df.at[i, column_name] = data
        
        # Write the updated DataFrame back to the Excel file
        df.to_excel(excel_file, index=False)
        print(f"Data inserted into '{column_name}' in '{excel_file}' successfully.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def get_distance(origin, destination):
    client = googlemaps.Client(key=api_key)
    try:
      distance_matrix = client.distance_matrix(origins=[origin], destinations=[destination])
      dis = distance_matrix["rows"][0]["elements"][0]["distance"]["value"]
      dis = dis/1000
      return dis
    except KeyError:
      return "Not Found"

if __name__ == "__main__":
  origin = "18.43455403564981, 79.10711311605102"
  schools = get_school()
  schools = [schools[item] + " ,Ganneruvaram" for item in range(1,len(schools))]
  
  destination = list()
  distances = list()
  
  print(schools)
  
  insert_data_into_excel(schools, 'dist-ganneruvaram.xlsx', "School Name", 2)
    
  for i in range(len(destination)):
      distances.append(get_distance(origin, schools[i]))