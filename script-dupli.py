import requests
from dotenv import load_dotenv
import os
import googlemaps
import pandas as pd


load_dotenv()

api_key = os.getenv('API_KEY')

# def get_mandal():
#   df = pd.read_excel("schools.xlsx")
#   mandals = df["MANDAL"].tolist()
#   return mandals

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
  schools = [item + " ,Ganneruvaram" for item in schools]
#   mandals = get_mandal()
  
  destination = list()
  distances = list()

#   insert_data_into_excel(mandals, 'output1.xlsx', "Mandal Name", 2)
  insert_data_into_excel(schools, 'dist-ganneruvaram.xlsx', "School Name", 2)
    
  for i in range(len(destination)):
      distances.append(get_distance(origin, schools[i]))
      
#   print(len(distances))
  
#   insert_data_into_excel(distances, 'dist-ganneruvaram.xlsx', "Distance", 6)
  
  
#   for i in range(len(mandals)):
#     if (mandals[i] == "CHIGURUMAMIDI"):
#       destination.append(f"{schools[i]},Chigurumamidi")
#   for i in range(len(mandals)):
#     if (mandals[i] == "CHOPPADANDI"):
#       destination.append(f"{schools[i]},Choppadandi")
#   for i in range(len(mandals)):
#     if (mandals[i] == "ELLANDAKUNTA"):
#       destination.append(f"{schools[i]},Ellandakunta")
#   for i in range(len(mandals)):
#     if (mandals[i] == "GANGADHARA"):
#       destination.append(f"{schools[i]},Gangadhara")
#   for i in range(len(mandals)):
#     if (mandals[i] == "GANNERVARAM"):
#       destination.append(f"{schools[i]},Gannervaram")
#   for i in range(len(mandals)):
#     if (mandals[i] == "HUZURABAD"):
#       destination.append(f"{schools[i]},Huzurabad")
#   for i in range(len(mandals)):
#     if (mandals[i] == "KARIMNAGAR"):
#       destination.append(f"{schools[i]},Karimnagar")
#   for i in range(len(mandals)):
#     if (mandals[i] == "KARIMNAGAR (RURAL)"):
#       destination.append(f"{schools[i]},Karimnagar (Rural)")
#   for i in range(len(mandals)):
#     if (mandals[i] == "KOTHAPALLE"):
#       destination.append(f"{schools[i]},Kothapalle")
#   for i in range(len(mandals)):
#     if (mandals[i] == "MANAKONDUR"):
#       destination.append(f"{schools[i]},Manakondur")
#   for i in range(len(mandals)):
#     if (mandals[i] == "RAMADUGU"):
#       destination.append(f"{schools[i]},Ramadugu")
#   for i in range(len(mandals)):
#     if (mandals[i] == "SAIDAPUR (V)"):
#       destination.append(f"{schools[i]},Saipadur")
#   for i in range(len(mandals)):
#     if (mandals[i] == "SHANKARAPATNAM"):
#       destination.append(f"{schools[i]},Shankarapatnam")
#   for i in range(len(mandals)):
#     if (mandals[i] == "THIMMAPUR"):
#       destination.append(f"{schools[i]},Thimmapur")
#   for i in range(len(mandals)):
#     if (mandals[i] == "VEENAVANKA"):
#       destination.append(f"{schools[i]},Veenavanka")