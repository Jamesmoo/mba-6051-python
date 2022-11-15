from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import xlrd as xlrd

# read excel file links:
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
# https://xlrd.readthedocs.io/en/latest/
# https://www.digitalocean.com/community/tutorials/pandas-read_excel-reading-excel-file-in-python

#global variables
pathToFile = "C:\\my-sync\\MBA 6051\\Group Project - Student Grades\\Student Grades Data.xlsx"
allColumns = ['gender', 'ethnicity', 'parental level of education', 'lunch',
              'test preparation course', 'math score', 'reading score', 'writing score']
def start():
    sheet_101_gender_male()

def sheet_101_gender_male():
    # header is: [excel row] - 1   => use Compsci where index starts at 0
    excel = pd.read_excel(pathToFile, sheet_name="101 - Gender Male", header=13, usecols=allColumns)
    print(excel)

