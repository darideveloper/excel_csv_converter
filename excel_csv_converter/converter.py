#! python3
# Conver each sheet for xlsx file to csv files

import os, openpyxl, csv, sys, errno

class Verify (): 
    """
    Main class to verify paths and extentions
    """

    def verify_path (self, path):
        """
        Verify is the from file and the to file path exist in the pc
        """ 

        if not os.path.exists (path): 
            raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), path)
        
    def verify_extension (self, file, extension):
        """
        Verify is the from file and the to file path exist in the pc
        """ 

        if not file.endswith (extension): 
            raise ValueError(errno.ENOENT, os.strerror(errno.ENOENT), file)


class Xlsx_to_csv (Verify): 
    """
    Convert each page from xlsx file, to csv files
    """

    def __init__ (self, file_xlsx, folder_destination):
        """ 
        Constructor of class. Get the from and to path files
        """

        self.file_xlsx = file_xlsx
        self.folder_destination = folder_destination

        if self.file_xlsx != "": 
            super().verify_path (self.file_xlsx)
            super().verify_extension (self.file_xlsx, ".xlsx")

        if self.folder_destination != "": 
            # Make folder if it doesn't exist

            try: 
                super().verify_path (self.folder_destination)
            except: 
                os.makedirs (folder_destination)

        self.xlsx_to_csv ()
    

    def xlsx_to_csv (self): 
        """
        Convert each page from csv file to csv files
        """
    
        wb = openpyxl.load_workbook(self.file_xlsx)

        # Loop through every sheet in the workbook
        for sheetName in wb.sheetnames: 
            sheet = wb[sheetName]

            #Crate the csv filename drom the xlsx file name and the sheet title.
            csvFile = os.path.basename(self.file_xlsx)[:-5] + '_' + sheetName + '.csv'
            csvPath = os.path.join(self.folder_destination, csvFile)
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
            file_name = os.path.basename (csvPath)
            print ('File "{}" generated.'.format (file_name))

class Csv_to_xlsx (Verify): 
    """
    Convert file csv to xlsx document, or insert csv data in existing file
    """

    def __init__ (self, file_csv, file_xlsx_destination):
        """ 
        Constructor of class. Get the from and to path files
        """

        self.file_csv = file_csv
        self.file_xlsx_destination = file_xlsx_destination


        if self.file_csv != "": 
            super().verify_path (self.file_csv)
            super().verify_extension (self.file_csv, ".csv")

        if self.file_xlsx_destination != "": 
            super().verify_extension (self.file_xlsx_destination, ".xlsx")

            # Make file if it dosen't exist

            try: 
                super().verify_path (os.path.dirname(self.file_xlsx_destination))
            except: 
                os.makedirs (os.path.dirname (self.file_xlsx_destination))

        self.csv_to_xlsx ()

    def csv_to_xlsx (self):
        """
        Convert csv file to xlsx file or insert data
        """ 

        # Read csv information
        file_csv = open (self.file_csv, 'r')
        reader = csv.reader (file_csv)
        data_csv = list (reader)

        # Verify if file exist
        if os.path.isfile (self.file_xlsx_destination): 
            wb = openpyxl.load_workbook (self.file_xlsx_destination)
            sheet_name = os.path.basename (self.file_csv) [:-4]

            sheets_names = wb.sheetnames
            last_sheet_number = ""

            # if sheet with the same already exist in the document, 
            #   get the number of last sheet
            for current_sheet_name in sheets_names: 
                if str(current_sheet_name).startswith (sheet_name):
                    number_sheet = str(current_sheet_name)[-1:]

                    try: 
                        number_sheet = int (number_sheet)
                        if number_sheet >= last_sheet_number:
                            last_sheet_number = number_sheet+1
                    except: 
                        last_sheet_number = 1
                    

            # Rename sheet
            if last_sheet_number != "":
                sheet_name += str(last_sheet_number) 
                
            
            # Select sheet
            sheet = wb.create_sheet (sheet_name)

            self.__write_data (sheet, data_csv)

            # Save sheet
            document_name = os.path.basename (self.file_xlsx_destination)
            print ('Data written in existing document: "{}" as sheet: "{}"'.format (document_name, sheet_name))
            wb.save (self.file_xlsx_destination)
        else: 
            # Make new file       
            wb = openpyxl.Workbook ()
            sheet_name = os.path.basename (self.file_csv) [:-4]
            sheet = wb["Sheet"]
            sheet.title = sheet_name

            self.__write_data (sheet, data_csv)

            # Save sheet
            document_name = os.path.basename (self.file_xlsx_destination)
            print ('Data written in new document: "{}" as sheet: "{}"'.format (document_name, sheet_name))
            wb.save (self.file_xlsx_destination)
        
            

    def __write_data (self, sheet, data): 
        """
        Write data in specific sheet of workbook
        """

        # Write information
        for row in data: 
            for cell in row:
                sheet.cell (data.index(row)+1, row.index (cell)+1).value = cell


