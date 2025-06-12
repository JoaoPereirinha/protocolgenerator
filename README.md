Excel to Word Document Automation
This repository contains a Python script developed for a university project that automates the process of transferring data from an Excel database to specific fields in Word document templates.

Overview
Filling out repetitive Word documents with data from spreadsheets is time-consuming and prone to errors. This script streamlines that process by automatically populating predefined fields in Word files using data from an Excel sheet, improving both speed and accuracy.

Features
Reads data from Excel (.xlsx) files

Automatically fills corresponding fields in Word (.docx) templates

Batch processing for multiple records

Each generated document is automatically named after the value in the "nome_aluno" column in the Excel database

Easy to configure for different template structures

User-friendly and fully written in Python

Requirements
The Excel file must include a column named nome_aluno. The value in this column will be used as the file name for each generated Word document.

Adjust the script to ensure that your template field names match your Excel columns.

Technologies Used
Python

openpyxl — for handling Excel files

python-docx — for editing Word documents

How to Use
Clone the repository:

bash
git clone https://github.com/your-username/excel-to-word-automation.git
cd excel-to-word-automation
Install the required libraries:

bash 
pip install openpyxl python-docx
Place your Excel database and Word template in the project directory.

Adjust the script (e.g., field names and paths) as needed.

Run the script:

bash
python main.py
The generated Word documents will be saved in the output folder, each named after the corresponding nome_aluno value.

Customization
Make sure your Excel column names and Word placeholders match.

Adapt the script for different template formats as needed.

License
This project was developed for educational purposes at university.
