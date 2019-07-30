from tkinter import *
from tkinter.filedialog import askopenfilename
import os
import pandas as pd
import xlrd

from Helpers.AverageIntakes import intake_averages
from Helpers.Globals import *
from Helpers.ResizeColumn import resize_columns
from Helpers.SpaceRows import space_rows

Tk().withdraw()
filename = askopenfilename()

if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
    sys.exit("Did not Choose an Excel Sheet")

directory = os.path.dirname(filename)
file = os.path.basename(directory)

# Create a Folder if its not already made to hold full intakes.
intake_path = r"C:/VantagePoints/Full-Intakes/"
if not os.path.exists(intake_path):
    os.makedirs(intake_path)

excels = []

# Get the path of each file and put it in a list.
for f in os.listdir(directory):
    if ".xls" in f and "Full" not in f:
        excels.append(os.path.abspath(os.path.join(directory, f)))
    else:
        continue

# Excel is not in a readable to the program form
database = [pd.ExcelFile(name) for name in excels]

# This pulls from each sheet.
full_data = [pd.read_excel(book, header=0, na_filter=False, sheet_name=0) for book in database]
short_data = [pd.read_excel(book, header=0, na_filter=False, sheet_name=1) for book in database]

# Combine the DataFrames into one for each. And Reset indexes for rows. Otherwise they will be wrong.
full_data = pd.concat(full_data).reset_index(drop=True)
short_data = pd.concat(short_data).reset_index(drop=True)

sort_race = short_data.copy()
sort_age = sort_race.copy()

sort_age.sort_values(by="Age", inplace=True)
# Create a category for Race to be sorted by
sort_race["Race"] = pd.Categorical(sort_race["Race"], ["W", "B", "L", "A", "NA", "O", "N/A"])
sort_race.sort_values(by="Race", inplace=True)

# Make an Excel book
full_data.to_excel(intake_path + "Total-Intakes.xlsx", sheet_name="Full Data", na_rep="NA")

# Allows me to write multiple sheets.
with pd.ExcelWriter(intake_path + "Intakes.xlsx") as writer:
    short_data.reset_index(drop=True).to_excel(writer, sheet_name="Unsorted", na_rep="NA")
    sort_race.reset_index(drop=True).to_excel(writer, sheet_name="Races", na_rep="NA")
    sort_age.reset_index(drop=True).to_excel(writer, sheet_name="Ages", na_rep="NA")

# Resize all the columns
resize_columns(intake_path + "Intakes.xlsx")
resize_columns(intake_path + "Total-Intakes.xlsx")

# Make a list of all lists for organizing you need. In order of sheets. So if Race is 2nd sheet. Do Race then age
# If Age is 3rd.
lists = [RACES, AGES]
groups = ["race", "age"]
# Place spaces between rows where they need to go for readability.
space_rows(intake_path + "Intakes.xlsx", lists, groups)

intake_averages(intake_path + "Intakes.xlsx", ["bdi", "ace", "cage", "bai"])

