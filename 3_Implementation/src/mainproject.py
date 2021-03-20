# from openpyxl import load_workbook
import openpyxl as xl
import pandas as pd 
import os

InputListPath = []
keysearch=[]
select = 0

#loadmaster = load_workbook("D:\\Temppp\\MasterSheet.xlsx")
masterpath = "D:\\Temppp\\MasterSheet.xlsx"



def checkstatus():
    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']
    mastermaxcol = mastersheet.max_column
    mastersheet.delete_rows(1, mastersheet.max_row - 1)


   


def searchandprint(path):
    # loading the workbook with given path
    print(path)
    loaded_workbook1 = xl.load_workbook(path)
    numsheet = loaded_workbook1.sheetnames
    loadingsheet = loaded_workbook1[numsheet[0]]
    colloadingsheet = loadingsheet.max_column
    print(numsheet)
    lensheet = len(numsheet)
    
    loadmaster = xl.load_workbook(masterpath)
    mastersheet = loadmaster['Sheet1']
    mastermaxcol = mastersheet.max_column
    mastermaxrow = mastersheet.max_row
    mastermaxrow = mastermaxrow+1

    var =mastermaxrow
    var = var+1
    for masrow in range(1, 4):
    
        mastersheet.cell(row= var, column= masrow).value = loadingsheet.cell(row= var, column = masrow).value
        

    
    mastercol = 0
    for i in range(0, lensheet):
        
        activesheet = loaded_workbook1[numsheet[i]]

        mastermaxcol = mastersheet.max_column
        mastermaxcol = mastermaxcol+1
        for masrow in range(4, colloadingsheet+1):
            mastersheet.cell(row=1 , column= mastermaxcol).value = activesheet.cell(row= 1, column = masrow).value
            mastermaxcol = mastermaxcol+1

        maxrows = activesheet.max_row
        maxcol = activesheet.max_column + 1
        print(maxrows,maxcol)
        k = mastercol + 1
        print(select)
        # masterbook.save(UserInput.outputPath[0])
        for rows in range(2, maxrows+1):
            if select == 1: 
                mastercol = mastersheet.max_column
                for col in range(1,3):
                    cellvalue1 = activesheet.cell(row= rows, column= col)
                    if str(cellvalue1.value) == str(keysearch[0]): 
                        mastermaxrow = mastersheet.max_row
                        l = mastermaxrow+1
                        for temp in range(1, maxcol):
                            mastersheet.cell(row= l, column = k).value = activesheet.cell(row= rows, column= temp).value
                            k = k+1   
            elif select == 0:
                mastercol = mastersheet.max_column
                for col in range(2,3):
                    cellvalue1 = activesheet.cell(row= rows, column= col)
                    if str(cellvalue1.value) == keysearch[0]:
                        l= mastermaxrow
                        for temp in range(1, maxcol):
                            mastersheet.cell(row= l, column = k).value = activesheet.cell(row= rows, column= temp).value
                            k = k+1
                  
        
    loadmaster.save(str('MasterSheet.xlsx'))
    if 'Sheet1' in loadmaster.sheetnames:
        ref = loadmaster['Sheet1']
        loadmaster.remove(ref) 
        loadmaster.close()

                    
def userpathinput():
    path = input("Enter your Path : ")
    InputListPath.append(path)
    # choice = input("Want to add more path :: Y/N: ")
    while True:
        choice = input("Add multiple path :: Y/N: ")
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
