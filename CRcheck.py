
import openpyxl
import numpy as np
import pandas as pd
import matplotlib as plt
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import sys
import xlsxwriter


while True:
    
    sys.setrecursionlimit(5000)
    print('please enter file location')
    file = input()
    print('please select first sheet')  
    firstsheet = input()
    print('please select second sheet')
    secondsheet = input()
    print('please enter your preferred column name,(ssno,name,salary)')
    searchcolumn = input()
    print('please enter preferred level of accuracy, an integer between 1 and 100')
    accuracy = input()

 #Extract work sheets and add a coluln where names are put in lower case and stripped of spaces

    wb1 =pd.read_excel(file, sheet_name = [str(firstsheet),str(secondsheet)],header = 0)
    ws1=wb1[firstsheet].loc[:,('ssno','name','salary')]
    ws1[searchcolumn] = ws1.apply(lambda x: str.lower(x[searchcolumn]), axis = 1)
    ws1[searchcolumn] = ws1.apply(lambda x: str.strip(x[searchcolumn]), axis = 1)

    ws2=wb1[secondsheet].loc[:,('ssno','name','salary')]
    ws2[searchcolumn] = ws2.apply(lambda x: str.lower(x[searchcolumn]), axis = 1)
    ws2[searchcolumn] = ws2.apply(lambda x: str.strip(x[searchcolumn]), axis = 1)

# Function to create dictionaries of index, name pairs

    Allnames={}

    def characterCounter(conReport):
        namesList = {}
        for i in conReport.index:
            v = conReport.at[(i,searchcolumn)]
             
            namesList.update({i:v})
       
        return  namesList               
 
 # Create name pairs for two sheets

    Allnames1  = characterCounter(ws1)
    Allnames2 = characterCounter(ws2)
 
 # Do one to many matching from sheet1 to sheet2 and vice versa and look for set difference 
    exactmatch1 = {}
    exactmatch2 = {}
    for j,name1 in Allnames1.items():
        for k,name2 in Allnames2.items():
            if fuzz.token_sort_ratio(name1,name2)>=int(accuracy):
       
                exactmatch1.update({j:name1})
           
    leftDiff ={}

    for x,y in Allnames1.items():
        if y not in exactmatch1.values():
            leftDiff.update({x+1:y})

    
    for j,name1 in Allnames1.items():
        for k,name2 in Allnames2.items():
            if fuzz.token_sort_ratio(name1,name2)>=int(accuracy):
       
                exactmatch2.update({k:name2})
                


    rightDiff = {}

    for x,y in Allnames2.items():
        if y not in exactmatch2.values():
            rightDiff.update({x+1:y})

#Print out results
#        
    print('\n names in '+str(firstsheet)+' and not in '+str(secondsheet)+' are given below:')
    print(leftDiff)

    print('\n names in '+str(secondsheet)+' and not in '+str(firstsheet)+' are given below:')
    print(rightDiff)


#   create and setup worksheets with headings for results  
    workbook = xlsxwriter.Workbook('Result.xlsx')
    worksheet1 = workbook.add_worksheet()
    worksheet1.write(0, 0, 'row number')
    worksheet1.write(0, 1, 'name')
    worksheet1.write(0, 2, 'salary')

    worksheet2 = workbook.add_worksheet()
    worksheet2.write(0, 0, 'row number')
    worksheet2.write(0, 1, 'name')
    worksheet3.write(0, 2, 'salary')

    worksheet3 = workbook.add_worksheet()
    worksheet3.write(0, 0, 'row number')
    worksheet3.write(0, 1, 'name')
    worksheet3.write(0, 2, 'salary')
    
    # write out the matched, left difference and right difference to 3 excel sheets

    row = 1
    col = 0
    
    for key in leftDiff.keys():
        
        worksheet1.write(row,col,key)
        worksheet1.write(row,col+1,leftDiff[key])
        worksheet1.write(row, col+2, ws1.loc[key, 'salary'])
        row += 1
    
    
    row = 1
    col = 0
    
    for key in rightDiff.keys():
      
        worksheet2.write(row,col,key)
        worksheet2.write(row,col+1,rightDiff[key])
        worksheet2.write(row, col+2, ws2.loc[key, 'salary'])
        row += 1
            
    row = 1
    col = 0
    
    for key in exactmatch1.keys():
       
        worksheet3.write(row,col,key)
        worksheet3.write(row,col+1,exactmatch1[key])
        worksheet3.write(row, col+2, ws1.loc[key, 'salary'])
        row += 1
        
    workbook.close() 

