import requests
from dotenv import load_dotenv
import os
import googlemaps
import getpara as gp
import pandas as pd


load_dotenv()

api_key = os.getenv('API_KEY')

def get_mandal():
  df = pd.read_excel("schools.xlsx")
  mandals = df["MANDAL"].tolist()
  print(mandals)
#   return mandals

def get_school():
  df = pd.read_excel("schools.xlsx")
  schools = df["School Name"].tolist()
  return schools
  
def get_distance(origin, destination):

  client = googlemaps.Client(key=api_key)
  distance_matrix = client.distance_matrix(origins=[origin], destinations=[destination])
  distance = distance_matrix["rows"][0]["elements"][0]["distance"]["value"]
  return distance

if __name__ == "__main__":
  origin = "18.43455403564981, 79.10711311605102"
#   school = get_school()
  mandals = get_mandal()
  destination = list()
  print(mandals)
#   for i in range(len(destination)):
#     if (mandals[i] == "Thimmapur"):
#         destination[i].replace(" ", "+")
#         destiantion[i] = f"{destiantion[i]} +','' + f'{mandals[i]}'"
#         print(destination[i])
    # distance = get_distance(origin, destination[i])

#   print(f"The distance between {origin} and {destination} is {distance/1000} kilometers.")
