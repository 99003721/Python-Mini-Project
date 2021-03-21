# from openpyxl import load_workbook
import openpyxl as xl
import pandas as pd 
import os
InputListPath = []
keysearch=[]
select = 0

masterpath = "D:\\Python-Mini-Project\\MasterSheet.xlsx"



def checkstatus():
    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']
    mastermaxcol = mastersheet.max_column
    mastermaxrow = mastersheet.max_row    
    mastersheet.delete_rows(2, mastersheet.max_row - 1)
    loadmaster.save(masterpath)
currrow = 0
currcol = 0

def print_header(lpath):
    loaded_workbook1 = xl.load_workbook(lpath)
    numsheet = loaded_workbook1.sheetnames
    loadingsheet = loaded_workbook1[numsheet[0]]

    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']

    max_row_master = mastersheet.max_row
    max_row_master = max_row_master+4
    print(max_row_master)
    max_col_sheet1 = loadingsheet.max_column
    print(max_col_sheet1) 

    for c in range(1, max_col_sheet1+1):     # changed to max_col_sheet1 in lace 4
        mastersheet.cell(row = max_row_master, column= c).value = loadingsheet.cell(row= 1, column= c).value
    currcol = mastersheet.max_column
    # mastersheet.font  = Font(color="FF0000")
    loadmaster.save(masterpath)
    return currcol



def searchandprint(path):
    # loading the workbook with given path
    print(path)
    lpath = path
    loaded_workbook1 = xl.load_workbook(path)
    numsheet = loaded_workbook1.sheetnames
    loadingsheet = loaded_workbook1[numsheet[0]]
    colloadingsheet = loadingsheet.max_column
    print(numsheet)
    lensheet = len(numsheet)
    
    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']

    colcurr = print_header(lpath)
    savecol = colcurr
    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']
   
    
    a = 1
    mastercol = 0
    print(lensheet)
    for i in range(0, lensheet):
        print("here")
        activesheet = loaded_workbook1[numsheet[i]]

        mastermaxcol = mastersheet.max_column
        mastermaxrow = mastersheet.max_row
        print(mastermaxcol)
          # needed to store the current position
        print("savecolumn",savecol)
        # mastermaxcol = mastermaxcol+1  #????

        temp_c = mastermaxcol+1

        # for masrow in range(4, colloadingsheet+1):   # copying the header from 4 col
        #     mastermaxcol = mastermaxcol+1
        #     mastersheet.cell(row= mastermaxrow , column= mastermaxcol).value = activesheet.cell(row= 1, column = masrow).value
        #     print("now you are here")
        
        maxrows = activesheet.max_row
        maxcol = activesheet.max_column
        print(maxrows,maxcol)
        # masterbook.save(UserInput.outputPath[0])
        for rows in range(2, maxrows+1):
            if select == 1: 
                for col in range(1,3):
                    cellvalue1 = activesheet.cell(row= rows, column= col)
                    if str(cellvalue1.value) == str(keysearch[0]): 
                        
                        if a == 1:   # printing name once
                            mastermaxrow = mastersheet.max_row
                            mastermaxrow = mastermaxrow+1
                            for r in range(1, 4):
                               mastersheet.cell(row= mastermaxrow, column = r).value = activesheet.cell(row= rows, column= r).value 
                               r = r+1
                            a = a+1
                        mastermaxrow = mastersheet.max_row +1
                        k = 4
                        for temp in range(4, maxcol+1):
                            mastersheet.cell(row= mastermaxrow, column = k).value = activesheet.cell(row= rows, column= temp).value
                            k = k+1 
            elif select == 0: 
                for col in range(1,3):
                    cellvalue1 = activesheet.cell(row= rows, column= col)
                    if str(cellvalue1.value) == str(keysearch[0]): 
                        
                        if a == 1:   # printing name once
                            mastermaxrow = mastersheet.max_row
                            mastermaxrow = mastermaxrow+1
                            for r in range(1, 4):
                               mastersheet.cell(row= mastermaxrow, column = r).value = activesheet.cell(row= rows, column= r).value 
                               r = r+1
                            a = a+1
                        mastermaxrow = mastersheet.max_row
                        k = 4
                        for temp in range(4, maxcol+1):
                            mastersheet.cell(row= mastermaxrow, column = k).value = activesheet.cell(row= rows, column= temp).value
                            k = k+1 
                  
        
    loadmaster.save(str('MasterSheet.xlsx'))
    # if 'Sheet1' in loadmaster.sheetnames:
    #     ref = loadmaster['Sheet1']
    #     loadmaster.remove(ref) 
    #     loadmaster.close()

                  
def userpathinput():
    path = input("Enter your Woorkbook Path : ")
    InputListPath.append(path)
    # choice = input("Want to add more path :: Y/N: ")
    while True:
        choice = input("Do You Want To Add More Workbook :: Y/N: ")
        if choice == "Y" or choice == "y":
            path = input("Paste new path here: ")
            InputListPath.append(path)
        elif choice == "N" or choice == "n":
            break


def Inputsearchkey():
    
    d1a = input (" Do you want to: A)\033[1m Search By PS Number :\033[0m B) \033[1mSearch By User Name :\033[0m [A/B]? :  ")
    if d1a == "A" or d1a == "a":

        user_PS_number = int(input("Enter PS number : "))
        keysearch.append(user_PS_number)  
        return 1 
    if d1a == "B" or d1a == "b": 
        user_name = str(input("Enter user name : "))
        keysearch.append(user_name)
        return 0
userpathinput()
select = Inputsearchkey()
lenpathlist = len(InputListPath)
checkstatus()
for i in range(0,lenpathlist):
    searchandprint(InputListPath[i])
