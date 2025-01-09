import os
import subprocess
import sys
import pandas as pd
import csv
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import re
from datetime import datetime


# Installing missing libraries using this function
def install_missing_packages():
    
    requirements_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requirements.txt")

    if os.path.exists(requirements_file):
        print(f"Found requirements.txt at: {requirements_file}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", requirements_file])
            print("All required packages are installed.")
        except subprocess.CalledProcessError as e:
            print(f"Error occurred while installing packages: {e}")
    else:
        print("No requirements.txt file found. Please ensure it is in the same directory as the script.")

install_missing_packages()


# Function to read pdf from the data
def extract_date_and_value(pdf_path):
    reader = PdfReader(pdf_path)
    text = " ".join(page.extract_text() for page in reader.pages)

    # To extract billing amounts from the pdf
    value = "Unknown"
    if "Gross Amount incl. VAT" in text:
        value_start = text.find("Gross Amount incl. VAT") + len("Gross Amount incl. VAT")
        value = text[value_start:].split()[0].strip()
        if "€" in text[value_start:]:
            value = "€ " + text[value_start:].split()[0].strip()
    elif "Total" in text:
        value_start = text.find("Total") + len("Total")
        value = text[value_start:].split()[0].strip()
        if "USD" in text[value_start:]:
            value =  text[value_start:].split()[1].strip()
            
            

    # to extract dates from the pdf invoices
    date = "Unknown"
    # Matching the format for sample_invoice_1
    date_match_month_name = re.search(r"(\d{1,2}[.\s][.\s][A-Za-zäöüÄÖÜß]+[.\s]\d{4})(?=\s|[A-Z])", text)
    if date_match_month_name:
        date = date_match_month_name.group(1)
    else:
        # Matching the format for sample_invoice_2
        date_match_english_month = re.search(r"\b([A-Za-z]+\s\d{1,2},\s\d{4})\b", text)
        if date_match_english_month:
            date = date_match_english_month.group(1)
        else:
            # To match generic dates if needed
            date_match_generic = re.search(r"\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b", text)
            if date_match_generic:
                date = date_match_generic.group(1)

    return date, value

# Function to get output to Excel and CSV files
def process_pdfs(output_csv, output_excel):
    """Process all PDFs in the current folder and create CSV and Excel files."""
    data = []
    current_folder = os.getcwd()  # Get the current working directory

    for file_name in os.listdir(current_folder):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(current_folder, file_name)
            date, value = extract_date_and_value(file_path)
            data.append({"File Name": file_name, "Date": date, "Value": value})

    # Creating the CSV file to store our data
    csv_file_path = os.path.join(current_folder, output_csv)
    with open(csv_file_path, mode="w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=["File Name", "Date", "Value"], delimiter=';')
        writer.writeheader()
        writer.writerows(data)

    # Creating the Excel sheet
    excel_file_path = os.path.join(current_folder, output_excel)
    df = pd.DataFrame(data)
    workbook = Workbook()

    # Adding the data sheet
    data_sheet = workbook.active
    data_sheet.title = "Data"
    for row in dataframe_to_rows(df, index=False, header=True):
        data_sheet.append(row)

    # Creating the pivot table
    pivot_sheet = workbook.create_sheet(title="Pivot Table")
    pivot_df = df.groupby(["Date", "File Name"]).sum().reset_index()
    for row in dataframe_to_rows(pivot_df, index=False, header=True):
        pivot_sheet.append(row)

    # Some styling of the pivot table
    table = Table(displayName="PivotTable", ref=f"A1:{chr(65 + len(pivot_df.columns) - 1)}{len(pivot_df) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    pivot_sheet.add_table(table)

    workbook.save(excel_file_path)

if __name__ == "__main__":
    output_csv = "data_csv.csv"
    output_excel = "data_excel.xlsx"

    print("Processing PDFs in the current folder...")
    process_pdfs(output_csv, output_excel)
    print(f"CSV and Excel files created: {output_csv}, {output_excel}")
