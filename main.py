#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Import the required modules
import os
import glob
from openpyxl import Workbook
from PyPDF2 import PdfReader

def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf = PdfReader(file)
        text = ''
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text += page.extract_text()
        return text

def save_text_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    for row, text in enumerate(data, start=1):
        ws.cell(row=row, column=1, value=text)
    wb.save(output_file)
    print(f"Data saved to {output_file}")

def main(folder_path, output_file):
    pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))
    if not pdf_files:
        print("No PDF files found in the specified folder.")
        return
    text_data = []
    for file in pdf_files:
        text = extract_text_from_pdf(file)
        text_data.append(text)
    save_text_to_excel(text_data, output_file)

# Specify the folder path containing the PDF files
folder_path = 'test_files'

# Specify the output Excel file path
output_file = 'output.xlsx'

if __name__ == '__main__':
    main(folder_path, output_file)

