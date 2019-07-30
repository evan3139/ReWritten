import os
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from docx import *
import xlsxwriter

from Helpers.CreateQuiz import create_quiz
from Helpers.Globals import HEADERS
from Helpers.Header import *
from Helpers.ResizeColumn import resize_columns
from Helpers.SheetContents import fill_intake_sheet

# Opens up file explorer

Tk().withdraw()
filename = askopenfilename()
if not filename.endswith(".docx"):
    sys.exit("Did not choose a .docx file.")

# This will get the name of the directory allowing it to be looped through all the files
directory = os.path.dirname(filename)
file = os.path.basename(directory)

intakepath = r'C:/VantagePoints/Intakes/'
if not os.path.exists(intakepath):
    os.makedirs(intakepath)
quizpath = r'C:/VantagePoints/Quiz-Template/'
if not os.path.exists(quizpath):
    os.makedirs(quizpath)

# Create the documents.
doc = os.path.join(quizpath + file + "-QuizTemplate.docx")
intakes = os.path.join(intakepath + file + ".xlsx")

# Create the one workbook, but two sheets.
workbook = xlsxwriter.Workbook(intakes)
full_sheet = workbook.add_worksheet()
score_sheet = workbook.add_worksheet()
workbook.close()

create_intake_header(filename, intakes, HEADERS)

for files in os.listdir(directory):
    if ".docx" in files:
        fname = directory + "/" + files
        fill_intake_sheet(fname, intakes, HEADERS)
    else:
        continue

resize_columns(intakes)

# This is function is super poorly coded because it was a time saver that isnt fully
# relevant to this. You can look through the code and change it to how you need it if you want
# Or you can make each quiz by hand. I think you'll see why I made it if you try that lol.
create_quiz(intakes, quizpath + "/" + file + ".docx", file)







