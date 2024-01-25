import os
from docx import Document
from PyPDF2 import PdfMerger
from datetime import datetime
import subprocess
import comtypes.client # word to PDF

from selenium import webdriver
from selenium.webdriver.edge.service import Service
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

import time


# Input parent directory path
directory_path = r'\\SepehrNAS\Thick Volume\CAREER\SEPEHR\Job Related\Resume & Cover Letter\APPLICATIONS\2024'
# Specify the path to Edge WebDriver executable
edge_driver_path = r'C:\Users\sepeh\OneDrive\Documents\Git\CoverLetterUpdater\msedgedriver.exe'

def replace_text_in_doc(doc, old_text, new_text):
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text

def save_as_pdf(word_file, pdf_file):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(word_file)
    doc.SaveAs(pdf_file, FileFormat=17)  # 17 represents the wdFormatPDF constant
    doc.Close()
    word.Quit()
    
def merge_pdfs(pdf_list, output):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output)
    merger.close()
    
def screenshot_webpage(url, save_path):
    driver = webdriver.Edge() 
    driver.get(url)
    
    # Set browser window size to page width and height
    required_width = driver.execute_script('return document.body.parentNode.scrollWidth')
    required_height = driver.execute_script('return document.body.parentNode.scrollHeight')
    driver.set_window_size(required_width, required_height)

    # Take screenshot and save
    driver.save_screenshot(save_path)
    driver.quit()
    
def save_webpage_as_html(url, save_path):
    # Set up Edge WebDriver
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service)

    driver.get(url)
    
    # # Wait for a specific element that indicates the page has loaded
    # WebDriverWait(driver, 10).until(
    # EC.presence_of_element_located((By.ID, "element_id"))
    # )
    
    # Wait 3 seconds for everything to load
    time.sleep(3)  
    
    with open(save_path, "w", encoding="utf-8") as file:
        file.write(driver.page_source)
        
    driver.quit()
    
# User Input
# url = input("Enter job posting url: ")
company = input("Enter the company name: ")
position = input("Enter the position: ")

# Get current date
current_date = datetime.now().strftime("%B %d, %Y")

# Create new folder for the application
company_folder = os.path.join(directory_path, company, f'{position} - {current_date}')
os.makedirs(company_folder, exist_ok=True)

# Define file paths
cover_letter_template = os.path.join(directory_path, 'Sepehr Salimi - Cover Letter.docx')
temp_docx = os.path.join(company_folder, f'Sepehr Salimi - Cover Letter - {company} - {position}.docx')
temp_pdf = os.path.join(company_folder, f'Sepehr Salimi - Cover Letter - {company} - {position}.pdf')
final_pdf = os.path.join(company_folder, f'Sepehr Salimi - Application - {company} - {position}.pdf')
resume_pdf = os.path.join(directory_path, 'Sepehr Salimi - Resume.pdf')
final_png = os.path.join(company_folder, f'Job Posting - {company} - {position}.png')
final_html = os.path.join(company_folder, f'Job Posting - {company} - {position}.html')


# Process Document
doc = Document(cover_letter_template)
replace_text_in_doc(doc, "COMPANY", company)
replace_text_in_doc(doc, "POSITION", position)
replace_text_in_doc(doc, "DATE", current_date)
doc.save(temp_docx)

# Convert to PDF and merge
save_as_pdf(temp_docx, temp_pdf)
merge_pdfs([temp_pdf, resume_pdf], final_pdf)

# # Screnshot job posting
# screenshot_webpage(url, final_png)

# # Save webpage as HTML
# save_webpage_as_html(url, final_html)

print("Done! Good luck!")