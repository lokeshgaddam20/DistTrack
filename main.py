import requests
from dotenv import load_dotenv
import os

load_dotenv()

api_key = os.getenv('API_KEY')

def get_lat_lng(place_name):
  """Gets the latitude and longitude of a place using the place name as search.

  Args:
    place_name: The name of the place to find the latitude and longitude of.

  Returns:
    A tuple containing the latitude and longitude of the place, or None if the place
    could not be found.
  """
  
  url = f"https://maps.googleapis.com/maps/api/geocode/json?address={placename}&key="+api_key
  response = requests.get(url)
  if response.status_code == 200:
    data = response.json()
    if data["status"] == "OK":
      return data["results"][0]["geometry"]["location"]["lat"], data["results"][0]["geometry"]["location"]["lng"]
  return None

# Example usage:

lat, lng = get_lat_lng("Pocharam, India")

print(lat, lng)
