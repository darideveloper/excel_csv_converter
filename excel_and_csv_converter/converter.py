#! python3
# Conver each sheet for xlsx file to csv files

import os, openpyxl, csv, sys

# Red the terminal
helpMenssage = """Write the path of the .xlsx file, and opcional the destination 
folder to the .csv files (example: main.py "user/myfolder/file.xlsx" "user/folderCSV"). 
If you only type the xlsx file path, the csv files will make in the parent folder (example: main.py "user/myfolder/file.xlsx")
If you have more that one file, type the parent folder of all xlsx files (example: main.py "user/myfolder") 
"""

if len(sys.argv) == 1: 
    print ('The program need more arguments. ' + helpMenssage)
    sys.exit()
elif len(sys.argv) == 2:  
    pathXlsx = sys.argv[1]
    pathDestiny = os.path.dirname(pathXlsx)
elif len(sys.argv) == 3:  
    pathXlsx = sys.argv[1]
    pathDestiny = sys.argv[2]
else: 
    print ('To much argument. ' + helpMenssage)
    sys.exit()

# Verify the paths
if not os.path.exists (pathXlsx) or not os.path.exists (pathDestiny): 
    print ('Check your paths')
    sys.exit()

def convertXlsxToCsv (excelFile): 
    """ Convert each sheet of a xlsx file to a csv file"""
    wb = openpyxl.load_workbook(excelFile)

    for sheetName in wb.sheetnames: 
        # Loop through every sheet in the workbook
        sheet = wb[sheetName]

        #Crate the csv filename drom the xlsx file name and the sheet title.
        csvFile = excelFile[:-5] + '_' + sheetName + '.csv'
        csvPath = os.path.join(pathDestiny, csvFile)
        outputFile = open(csvPath, 'w', newline='')

        #crate the csv.writer object for this csv file
        outputWriter = csv.writer(outputFile)

        # Loop though every row un the sheet
        for rowNum in range (1, sheet.max_row + 1): 
            rowData = [] #Appen each cell to this list
            #Loop trough each cell in the row
            for colNum in range (1, sheet.max_column + 1): 
                #Appen each cell's data to row data
                rowData.append(sheet.cell(rowNum, colNum).value)            
            #Write  the rowData list to the csv file.
            outputWriter.writerow(rowData)
        outputFile.close()
        print ("File %s generated." % (csvPath))

if pathXlsx.endswith('.xlsx'): 
    convertXlsxToCsv (pathXlsx)
else:
    for currentExcelFile in os.listdir(pathXlsx): 
        # Only convert xlsx files
        if str(currentExcelFile).endswith('.xlsx'): 
            convertXlsxToCsv (currentExcelFile)
            

