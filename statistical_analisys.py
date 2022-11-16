from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import xlrd as xlrd
from enum import Enum

# read excel file links:
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
# https://xlrd.readthedocs.io/en/latest/
# https://www.digitalocean.com/community/tutorials/pandas-read_excel-reading-excel-file-in-python
# https://www.youtube.com/watch?v=FniLzpaSFGk
# https://pandas.pydata.org/pandas-docs/stable/user_guide/10min.html


#global variables
pathToFile = "C:\\my-sync\\MBA 6051\\Group Project - Student Grades\\Student Grades Data.xlsx"
allColumns = ['gender', 'ethnicity', 'parental level of education', 'lunch',
              'test preparation course', 'math score', 'reading score', 'writing score']

class ExcelColumns(Enum):
    gender = 'gender'
    ethnicity = 'ethnicity'
    education = 'parental level of education'
    lunch = 'lunch'
    preparation = 'test preparation course'
    math = 'math score'
    reading = 'reading score'
    writing = 'writing score'


class ExcelSheets(Enum):
    template = 0
    index = 1
    rawdata = 2
    gender_male_101 = 3
    gender_female_102 = 4
    parental_education_103 = 5
    ethnicity_group_a_104 = 6
    ethnicity_group_b_105 = 7
    ethnicity_group_c_106 = 8
    ethnicity_group_d_107 = 9
    count_by_lunch_108 = 10
    count_by_preparation_109 = 11
    ethnicity_education_lunch_110 = 12
    education_preparation_lunch_111 = 13
    count_preparation_education_lunch = 14
    m1_math_score_frequency = 15
    m2_reading_score_frequency = 16
    m3_writing_score_frequency = 17
    m4_score_correlation = 18
    m5_population_proportions = 19


# NOTE: all sheets must have headers in row 10
# manual sheets have headers in row 1
# this must be constant on all excel sheet
def loadExcelSheet(sheetnumber):
    table_header = 9
    if sheetnumber.value > 14:
        table_header = 0
    return pd.read_excel(pathToFile, sheet_name=sheetnumber.value, header=table_header)

def start():
    # each sheet analysis will be its own definition
    sheet_101_gender_male()

def sheet_101_gender_male():
    print('===========================')
    print('Sheet: 101 Gender Male')
    excel = loadExcelSheet(ExcelSheets.gender_male_101)

    # gets all the values in a  column, but it can only get 1 at a time
    math_scores = excel.loc[:,ExcelColumns.math.value]
    writing_scores = excel.loc[:,ExcelColumns.writing.value]
    reading_scores = excel.loc[:,ExcelColumns.reading.value]

    print(math_scores)
    print(writing_scores)
    print(reading_scores)

    #print(math_scores)

def sheet_102_gender_female():
    print('===========================')
    print('Sheet: 102 Gender Female')
    excel = loadExcelSheet(ExcelSheets.gender_female_102)
