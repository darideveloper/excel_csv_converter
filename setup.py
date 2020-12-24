#! python3
import os
from setuptools import setup, find_packages

# read the contents of your README file
this_directory = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(this_directory, 'README.md')) as f:
    long_description = f.read()

setup (
	name = "excel_csv_converter",
	version = "1.0.0", 
	description = "Package that allows, with regular expressions, to extract from texts, phone numbers and emails",
	long_description = long_description,
	long_description_content_type = "text/markdown",
	author = "Dari Developer",
	author_email = "hernandezdarifrancisco@gmail.com",
	license = "MIT",
	keywords = "extract, re, phones, emails",
	project_urls = {
		"Documentation": "https://github.com/DariHernandez/phone_and_email_extractor/blob/master/README.md",
		"Funding": "https://www.paypal.com/paypalme/FranciscoDari",
		"Source": "https://github.com/DariHernandez/phone_and_email_extractor"
		},
	packages = find_packages(include=["excel_csv_converter", "excel_csv_converter.*"]),
	install_requires = ["pyperclip"],
	python_requires = ">=3.7"
)
