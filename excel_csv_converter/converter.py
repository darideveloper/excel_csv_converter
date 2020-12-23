#! python3
# Conver each sheet for xlsx file to csv files

import os, openpyxl, csv, sys, errno

class Convert (): 
    """
    Convert each page from xlsx file, to csv files,
    or insert csv information in xlsx file
    """

    def __init__ (self, file_csv = "", file_xlsx = ""):
        """ 
        Constructor of class. Get the from and to path files
        """

        self.file_csv = file_csv
        self.file_xlsx = file_xlsx

        if self.file_csv != "": 
            self.__verify_path (self.file_csv) 
            self.__verify_extension (self.file_csv, ".csv")
        
        if self.file_xlsx != "": 
            self.__verify_path (self.file_xlsx)
            self.__verify_extension (self.file_xlsx, ".xlsx")

    def __verify_path (self, path):
        """
        Verify is the from file and the to file path exist in the pc
        """ 

        if not os.path.exists (path): 
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), path)
        
    def __verify_extension (self, file, extension):
        """
        Verify is the from file and the to file path exist in the pc
        """ 

        if not file.endswith (extension): 
            raise ValueError(errno.ENOENT, os.strerror(errno.ENOENT), file)

    def xlsx_to_csv (self, destionarion_folder): 
        """
        Convert each page from csv file to csv files
        """


        self.__verify_path (destionarion_folder)
    
        wb = openpyxl.load_workbook(self.file_xlsx)

        # Loop through every sheet in the workbook
        for sheetName in wb.sheetnames: 
            sheet = wb[sheetName]

            #Crate the csv filename drom the xlsx file name and the sheet title.
            csvFile = os.path.basename(self.file_xlsx)[:-5] + '_' + sheetName + '.csv'
            csvPath = os.path.join(destionarion_folder, csvFile)
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

            


