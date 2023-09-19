import pandas as pd

def get_destination():
  df = pd.read_excel("schools.xlsx")
  destinations = df["School Name"].tolist()
  return destinations

print(get_destination())


# Here are some code snippets from other files of the repo:et_destination()