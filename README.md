# Excel csv converter
This is a project without GUI. 
**Convert** each page of a **Excel (xlsx)** file, or path with xlsx files, **to cvs files**.

# Install
``` bash
$
```


# How to use

``` python
# import converter
from excel_csv_converter import converter

# CONVERT XLSX DOCUMENTS TO CSVs

file_xlsx = "c:\\my_file.xlsx"
folder_destination = "c:\\my_folder"

my_converter = converter.Xlsx_to_csv (file_xlsx, folder_destination)

# INSERT CSV FILE IN XLSX DOCUMENT

file_csv = "c:\\my_file.csv"
file_xlsx_destination = "c:\\my_file.xlsx"

my_converter = converter.Csv_to_xlsx (file_csv, file_xlsx_destination)
```