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

drop a value from a series
https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.Series.drop.html
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
    group_a_parent_ed_115 = 20
    group_b_parent_ed_116 = 21
    group_c_parent_ed_117 = 22
    group_d_parent_ed_118 = 23
    ethnicity_group_e_119 = 24
    group_e_parent_ed_118 = 25
    random_sample_20 = 26
    random_sample_30 = 27

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
    math_reading_writing('Python Check - Sheet: 101 Gender Male', loadExcelSheet(ExcelSheets.gender_male_101))
    math_reading_writing('Python Check - Sheet: 102 Gender Female', loadExcelSheet(ExcelSheets.gender_female_102))
    math_reading_writing('Python Check - Sheet: 104 Ethnicity Group A', loadExcelSheet(ExcelSheets.ethnicity_group_a_104))
    math_reading_writing('Python Check - Sheet: 105 Ethnicity Group B', loadExcelSheet(ExcelSheets.ethnicity_group_b_105))
    math_reading_writing('Python Check - Sheet: 106 Ethnicity Group C', loadExcelSheet(ExcelSheets.ethnicity_group_c_106))
    math_reading_writing('Python Check - Sheet: 107 Ethnicity Group D', loadExcelSheet(ExcelSheets.ethnicity_group_d_107))
    """
        had to do a workaround 
        for some reason python could not pull the data from the sheet
        so i copy pasted the table from 119 ethenicity E to the template and then had python
        pull the data from the template and it worked and crunched the numbers        
    """
    math_reading_writing('Python Check - Sheet: 119 Ethnicity Group E', loadExcelSheet(ExcelSheets.template))
    outputFile.close()
    print('processing completed')


def summary_calc(excel, title):
    savesameline('====> Calculating: ', title)

    np_mean = np.mean(excel)
    np_std = np.std(excel)
    np_median = np.median(excel)
    np_variance = np.var(excel)
    ex_min = excel.min()
    ex_max = excel.max()
    ex_range = int(excel.max()) - int(excel.min())
    ex_size = excel.size

    savesameline('==> record count: ', ex_size)

    newline()

    savesameline('no outliers removed - straight calculation - numpy calculated', '')
    savesameline(title + ' Scores Mean:', np_mean)
    savesameline(title + ' Scores Standard Deviation:', np_std)
    savesameline(title + ' Scores Median:', np_median)
    savesameline(title + ' Scores Variance:', np_variance)
    savesameline(title + ' Scores Minimum:', ex_min)
    savesameline(title + ' Scores Max:', ex_max)
    savesameline(title + ' Scores range:', ex_range)
    savesameline(title + ' Scores Count:', ex_size)

    newline()

    std = excel.std()
    mean = np.mean(excel)
    z_score = (excel - mean) / std
    outliers = excel[abs(z_score) > 3]

    quantile_one = excel.quantile(.25)
    quantile_three = excel.quantile(.75)
    inter_quantile_range = quantile_three - quantile_one
    upper_threshold = quantile_three + 1.5 * inter_quantile_range
    lower_threshold = quantile_one - 1.5 * inter_quantile_range

    excel_array = np.array(excel.values.tolist())

    savesameline(title + ' inter quartile range - manually calculated:', inter_quantile_range)
    savesameline(title + ' quartile upper threshold - manually calculated:', upper_threshold)
    savesameline(title + ' quartile lower threshold - manually calculated:', lower_threshold)

    numpy_quantile_25 = str(np.percentile(excel_array, 25))
    numpy_quantile_50 = str(np.percentile(excel_array, 50))
    numpy_quantile_75 = str(np.percentile(excel_array, 75))

    newline()

    savesameline(title + ' numpy - 25 quantile:', numpy_quantile_25)
    savesameline(title + ' numpy - 50 quantile:', numpy_quantile_50)
    savesameline(title + ' numpy - 75 quantile:', numpy_quantile_75)

    newline()

    # manual removal of outliers  to check if outliers are being automatically removed on the above code by numpy
    no_outliers_values = excel
    if outliers.size:
        savesameline('found these outliers: ', '')
        savesameline('', outliers)
        no_outliers_values = excel.drop(outliers)

    newline()
    outliers_removed_array = np.array(no_outliers_values.values.tolist())

    manual_quantile25 = str(np.quantile(outliers_removed_array, .25))
    manual_quantile50 = str(np.quantile(outliers_removed_array, .50))
    manual_quantile75 = str(np.quantile(outliers_removed_array, .75))

    savesameline(title + ' - outliers manually removed - 25 percentile:', manual_quantile25)
    savesameline(title + ' - outliers manually removed - 50 percentile:', manual_quantile50)
    savesameline(title + ' - outliers manually removed - 75 percentile:', manual_quantile75)

    newline()

    savesameline(title + ' quantile 25 - numpy: ' + numpy_quantile_25 + ' -- manual: ' + manual_quantile25, '')
    savesameline(title + ' quantile 50 - numpy: ' + numpy_quantile_50 + ' -- manual: ' + manual_quantile50, '')
    savesameline(title + ' quantile 75 - numpy: ' + numpy_quantile_75 + ' -- manual: ' + manual_quantile75, '')

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
        
    outliers = excel[(np.abs(stats.stats.zscore(excel)) < 3).all(axis=1)]
    print(np.abs(stats.stats.zscore(excel)))
    excel_no_outliers = excel[(np.abs(stats.stats.zscore(excel)) < 3)]
    print(excel_no_outliers)
    """
