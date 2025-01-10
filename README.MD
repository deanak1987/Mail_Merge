# Doc Merge

This python program automatically sends emails to a list of recipients and attaches a document containing a preset content.

## How to Use
1. Must have a CSV file containing recipient info i.e. name, email address, etc. An example is produced using CsvFileCreator.py - simply change the values or use your own CSV creator.
2. A template document containing placeholders for the required information
3. An email account that is preconfigured to allow the program to send the email from, i.e. sending address, password, connection, etc. Gmail is used as an example.
4. A message title and body to address the recipient. An example is given.
5. Run DocMerge.py file to auto send messages

## Files
CsvFileCreator.py - A simple python file for creating CSV file

DocMerge.py - Main python program for creating and sending email with attached document

README.MD - This file with project info

recipients.csv - CSV file containing recipient info

requirements.txt - Required dependencies for program to operate

template.docx - a word doc that is altered accordingly to contain recipient info