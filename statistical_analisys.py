from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import xlrd as xlrd
from enum import Enum
import os

# read excel file links:
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
# https://xlrd.readthedocs.io/en/latest/
# https://www.digitalocean.com/community/tutorials/pandas-read_excel-reading-excel-file-in-python
# https://www.youtube.com/watch?v=FniLzpaSFGk
# https://pandas.pydata.org/pandas-docs/stable/user_guide/10min.html


#global variables
folderPath = "C:\\my-sync\\MBA 6051\\Group Project - Student Grades\\"
excelFileName = "Student Grades Data.xlsx"
completeExcelPath = folderPath + excelFileName

outputFilePath = folderPath + "stats_calculation.txt"
if os.path.exists(outputFilePath):
    os.remove(outputFilePath)
outputFile = open(outputFilePath, "a")

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

def save(text):
    parsed_text = str(text)
    print(parsed_text)
    outputFile.write(parsed_text)
    newline()

def savesameline(textOne, textTwo):
    parsed_text = str(textOne) + ' ' + str(textTwo)
    print(parsed_text)
    outputFile.write(parsed_text)
    newline()
def newline():
    print('\n')
    outputFile.write("\n")

def loadExcelSheet(sheetnumber):
    table_header = 9
    if sheetnumber.value > 14:
        table_header = 0
    return pd.read_excel(completeExcelPath, sheet_name=sheetnumber.value, header=table_header)

def start():
    # each sheet analysis will be its own definition
    sheet_101_gender_male()
    sheet_102_gender_female()

    outputFile.close()

def summary_calc(excel, title):
    # gets all the values in a  column, but it can only get 1 at a time
    math_scores = excel.loc[:,ExcelColumns.math.value]
    writing_scores = excel.loc[:,ExcelColumns.writing.value]
    reading_scores = excel.loc[:,ExcelColumns.reading.value]

    #its a pandas series object
    #savesameline('Type of Object - just to check',str(type(math_scores)))

    savesameline('Math Scores Mean:', np.mean(math_scores))
    savesameline('Math Scores Standard Deviation:', np.std(math_scores))
    savesameline('Math Scores Median:', np.median(math_scores))
    savesameline('Math Scores Variance:', np.var(math_scores))
    savesameline('Math Scores Minimum:', math_scores.min())
    savesameline('Math Scores Max:', math_scores.max())
    savesameline('Math Scores range:', int(math_scores.max()) - int(math_scores.min()))
    savesameline('Math Scores Count:', math_scores.size)
    savesameline('Math Scores Quartile 1:', math_scores.quantile(.25))
    savesameline('Math Scores Quartile 2:', math_scores.quantile(.5))
    savesameline('Math Scores Quartile 3:', math_scores.quantile(.75))
    newline()

    savesameline('Writing Scores Mean:', np.mean(writing_scores))
    savesameline('Writing Scores Standard Deviation:', np.std(writing_scores))
    savesameline('Writing Scores Median:', np.median(writing_scores))
    savesameline('Writing Scores Variance:', np.var(writing_scores))
    savesameline('Writing Scores Minimum:', writing_scores.min())
    savesameline('Writing Scores Max:', writing_scores.max())
    savesameline('Writing Scores range:', int(writing_scores.max()) - int(writing_scores.min()))
    savesameline('Writing Scores Count:', writing_scores.size)
    savesameline('Writing Scores Quartile 1:', writing_scores.quantile(.25))
    savesameline('Writing Scores Quartile 2:', writing_scores.quantile(.5))
    savesameline('Writing Scores Quartile 3:', writing_scores.quantile(.75))
    newline()

    savesameline('Reading Scores Mean:', np.mean(reading_scores))
    savesameline('Reading Scores Standard Deviation:', np.std(reading_scores))
    savesameline('Reading Scores Median:', np.median(reading_scores))
    savesameline('Reading Scores Variance:', np.var(reading_scores))
    savesameline('Reading Scores Minimum:', reading_scores.min())
    savesameline('Reading Scores Max:', reading_scores.max())
    savesameline('Reading Scores range:', int(reading_scores.max()) - int(reading_scores.min()))
    savesameline('Reading Scores Count:', reading_scores.size)
    savesameline('Reading Scores Quartile 1:', reading_scores.quantile(.25))
    savesameline('Reading Scores Quartile 2:', reading_scores.quantile(.5))
    savesameline('Reading Scores Quartile 3:', reading_scores.quantile(.75))
    newline()

def sheet_101_gender_male():
    save('==========================================')
    save('Sheet: 101 Gender Male')
    save('==========================================')
    excel = loadExcelSheet(ExcelSheets.gender_male_101)
    math_scores = excel.loc[:,ExcelColumns.math.value]
    writing_scores = excel.loc[:,ExcelColumns.writing.value]
    reading_scores = excel.loc[:,ExcelColumns.reading.value]

    summary_calc(excel)


def sheet_102_gender_female():
    save('==========================================')
    save('Sheet: 102 Gender Female')
    save('==========================================')
    excel = loadExcelSheet(ExcelSheets.gender_female_102)
    summary_calc(excel)