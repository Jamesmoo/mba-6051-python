from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import scipy as stats
import seaborn as sns
from enum import Enum
import os

"""
These are links for doing things in python
read excel file links:
https://www.geeksforgeeks.org/reading-excel-file-using-python/
https://xlrd.readthedocs.io/en/latest/
https://www.digitalocean.com/community/tutorials/pandas-read_excel-reading-excel-file-in-python
https://www.youtube.com/watch?v=FniLzpaSFGk
https://pandas.pydata.org/pandas-docs/stable/user_guide/10min.html

outliers 
https://hersanyagci.medium.com/detecting-and-handling-outliers-with-pandas-7adbfcd5cad8
https://linuxhint.com/pandas-remove-outliers/
"""

#global variables
folderPath = "C:\\my-sync\\MBA 6051\\Group Project - Student Grades\\"
excelFileName = "Student Grades Data.xlsx"
completeExcelPath = folderPath + excelFileName

outputFilePath = folderPath + "python_stats_calculation.txt"
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
    #print(parsed_text)
    outputFile.write(parsed_text)
    newline()

def savesameline(textOne, textTwo):
    parsed_text = str(textOne) + ' ' + str(textTwo)
    #print(parsed_text)
    outputFile.write(parsed_text)
    newline()
def newline():
    #print('\n')
    outputFile.write("\n")

def loadExcelSheet(sheetnumber):
    table_header = 9
    if sheetnumber.value > 14:
        table_header = 0
    return pd.read_excel(completeExcelPath, sheet_name=sheetnumber.value, header=table_header)

def start():
    print('Processing Started')
    # each sheet analysis will be its own definition
    math_reading_writing('Sheet: 101 Gender Male', loadExcelSheet(ExcelSheets.gender_male_101))
    math_reading_writing('Sheet: 102 Gender Female', loadExcelSheet(ExcelSheets.gender_female_102))
    math_reading_writing('Sheet: 104 Ethnicity Group A', loadExcelSheet(ExcelSheets.ethnicity_group_a_104))
    math_reading_writing('Sheet: 105 Ethnicity Group B', loadExcelSheet(ExcelSheets.ethnicity_group_b_105))
    math_reading_writing('Sheet: 106 Ethnicity Group C', loadExcelSheet(ExcelSheets.ethnicity_group_c_106))
    math_reading_writing('Sheet: 107 Ethnicity Group D', loadExcelSheet(ExcelSheets.ethnicity_group_d_107))
    outputFile.close()
    print('processing completed')


def summary_calc(excel, title):
    #print(type(excel), 'struct')
    savesameline(title + ' Scores Mean:', np.mean(excel))
    savesameline(title + ' Scores Standard Deviation:', np.std(excel))
    savesameline(title + ' Scores Median:', np.median(excel))
    savesameline(title + ' Scores Variance:', np.var(excel))
    savesameline(title + ' Scores Minimum:', excel.min())
    savesameline(title + ' Scores Max:', excel.max())
    savesameline(title + ' Scores range:', int(excel.max()) - int(excel.min()))
    savesameline(title + ' Scores Count:', excel.size)

    std = excel.std()
    mean = np.mean(excel)
    z_score = (excel - mean) / std
    # print(z_score, 'z-score')

    outliers = excel[abs(z_score) > 3]
    # print(outliers)
    #
    quantile_one = excel.quantile(.25)
    quantile_three = excel.quantile(.75)
    inter_quantile_range = quantile_three - quantile_one
    upper_threshold = quantile_three + 1.5 * inter_quantile_range
    lower_threshold = quantile_one - 1.5 * inter_quantile_range

    no_outliers = excel[abs(z_score) <= 3]
    print(no_outliers)

    savesameline(title + 'q1', quantile_one)
    savesameline(title + 'q3', quantile_three)
    savesameline(title + 'iqr', inter_quantile_range)
    savesameline(title + 'upper', upper_threshold)
    savesameline(title + 'lower', lower_threshold)

    savesameline(title + ' Scores Quartile 1:', excel.quantile(.25))
    savesameline(title + ' Scores Quartile 2:', excel.quantile(.5))
    savesameline(title + ' Scores Quartile 3:', excel.quantile(.75))
    newline()

def math_reading_writing(title, excel):
    save('==========================================')
    save(title)
    save('==========================================')

    math_scores = excel.loc[:, ExcelColumns.math.value]
    writing_scores = excel.loc[:, ExcelColumns.writing.value]
    reading_scores = excel.loc[:, ExcelColumns.reading.value]

    summary_calc(math_scores, 'Math')
    summary_calc(writing_scores, 'Writing')
    summary_calc(reading_scores, 'Reading')

    """
        stats.stats.zscore()
        a_no_outliers = a[(np.abs(stats.zscore(a)) < 3)]
        print(a_no_outliers)
    """

    print(type(excel))
    # outliers = excel[(np.abs(stats.stats.zscore(excel)) < 3).all(axis=1)]
    # print(np.abs(stats.stats.zscore(excel)))
    # excel_no_outliers = excel[(np.abs(stats.stats.zscore(excel)) < 3)]
    # print(excel_no_outliers)


    #outliers
    summary_calc(math_scores.dropna(), 'Math - Outliers removed -')
    summary_calc(writing_scores.dropna(), 'Writing - Outliers removed -')
    summary_calc(reading_scores.dropna(), 'Reading - Outliers removed -')
