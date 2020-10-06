# excel to csv
## Description
Convert each page of a Excel (xlsx) file, or path with xlsx files, to cvs files. 
## How to use
### Install modules
**$ pip3 install openpyxl**
### Run the program
**$ main.py "usr/myPath/file.xlsx" "usr/destinyPath"** # Convert one file and save the csv files in specific folder
**$ main.py "usr/myPath/file.xlsx"** # Save the csv files in parent folder
**$ main.py "usr/myPath" "usr/destinyPath"** # Convert all xlsx files froma path and save the csv files in specific folder
**$ main.py "usr/myPath"** # Convert all xlsx files froma path and save the csv in parent folder