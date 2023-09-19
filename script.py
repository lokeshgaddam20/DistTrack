import requests
from dotenv import load_dotenv
import os
import googlemaps
import pandas as pd
import openpyxl

load_dotenv()

api_key = os.getenv('API_KEY')

def get_mandal():
  df = pd.read_excel("schools.xlsx")
  mandals = df["MANDAL"].tolist()
  return mandals

def create_excel_sheet(iterable_object_data, column_name, column_no):
  wb = openpyxl.Workbook()
  ws = wb.create_sheet()

  ws.cell(row=1, column=column_no).value = column_name

  for row_index, row_data in enumerate(iterable_object_data):
    ws.cell(row=row_index + 2, column=column_no).value = row_data
  wb.save('output.xlsx')

def get_school():
  df = pd.read_excel("schools.xlsx")
  schools = df["School Name"].tolist()
  return schools
  
def get_distance(origin, destination):

  client = googlemaps.Client(key=api_key)
  distance_matrix = client.distance_matrix(origins=[origin], destinations=[destination])
  distance = distance_matrix["rows"][0]["elements"][0]["distance"]["value"]
  distance = distance/1000
  return distance

if __name__ == "__main__":
  origin = "18.43455403564981, 79.10711311605102"
  schools = get_school()
  mandals = get_mandal()
  destination = list()
  create_excel_sheet(mandals, "Mandal Name",2)
  create_excel_sheet(schools, "School Name",4)
  # for i in range(len(mandals)):
  #     destination.append(f"{schools[i]},Thimmapur")
  # for j in range(len(destination)):
  #     print(get_distance(origin, destination[j]))
    
    
    # if (mandals[i] == "THIMMAPUR"):




  # print(destination)
#   print(f"The distance between {origin} and {destination} is {distance/1000} kilometers.")