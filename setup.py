#! python3
import os
from setuptools import setup, find_packages

# read the contents of your README file
this_directory = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(this_directory, 'README.md')) as f:
    long_description = f.read()

setup (
	name = "excel_csv_converter",
	version = "1.0.1", 
	description = "Convert each page from xlsx file to csv files, convert file csv to xlsx document, insert csv data in existing file",
	long_description = long_description,
	long_description_content_type = "text/markdown",
	author = "Dari Developer",
	author_email = "hernandezdarifrancisco@gmail.com",
	license = "MIT",
	keywords = "xlsx, csv, excel, extract, convert",
	project_urls = {
		"Documentation": "https://github.com/DariHernandez/excel_csv_converter/blob/master/README.md",
		"Funding": "https://www.paypal.com/paypalme/FranciscoDari",
		"Source": "https://github.com/DariHernandez/excel_csv_converter/"
		},
	packages = find_packages(include=["excel_csv_converter", "excel_csv_converter.*"]),
	install_requires = ["openpyxl"],
	python_requires = ">=3.7"
)
