# excel to csv
## Description
This is a project without GUI. 
**Convert** each page of a **Excel (xlsx)** file, or path with xlsx files, **to cvs files**. 
## How to use
### Install modules
```bash
$ pip3 install openpyxl
```
### Run the program
```bash
# Convert one file and save the csv files in specific folder
$ python3 main.py "usr/myPath/file.xlsx" "usr/destinyPath"** 

# Save the csv files in parent folder
$ python3 main.py "usr/myPath/file.xlsx"** 

# Convert all xlsx files froma path and save the csv files in specific folder
$ python3 main.py "usr/myPath" "usr/destinyPath"** 

# Convert all xlsx files from a path and save the csv in parent folder
$ python3 main.py "usr/myPath"** 
```
