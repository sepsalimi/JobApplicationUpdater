import os
from docx import Document
from PyPDF2 import PdfMerger
from datetime import datetime
import subprocess
import comtypes.client # word to PDF
import shutil

from selenium import webdriver
from selenium.webdriver.edge.service import Service
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

import time

# Input your name
Name = 'Sepehr Salimi'
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
    
# Reduces file name size if required
def create_shortened_paths_if_needed():
    global cover_letter_docx, final_coverletter, final_application_pdf, final_resume_pdf
    if len(cover_letter_file_path) > 83:
        cover_letter_docx = cover_letter_docx_shortened
        final_coverletter = final_coverletter_shortened
        final_application_pdf = final_application_pdf_shortened
        final_resume_pdf = final_resume_pdf_shortened    
   
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
job_title = input("Enter the RE: Position: ")
role = input("... excitement that I submit my application for the ____ position: ")
skill_role = input("Join company..., further advancing my skills in: ")

# Get current date
CurrentDateCoverLetter = datetime.now().strftime("%B %d, %Y")
CurrentDateFileName = datetime.now().strftime("%Y%m%d")

# Create new folder for the application
company_folder = os.path.join(directory_path, "Applications Sent Out", company, f'{job_title} - {CurrentDateFileName}')
os.makedirs(company_folder, exist_ok=True)

cover_letter_file_path = f'{Name} - Cover Letter - {company} - {job_title}.pdf' # used to check if file name length is appropriate

# Define file paths
cover_letter_template = os.path.join(directory_path, f'{Name} - Cover Letter.docx')
cover_letter_docx = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {job_title}.docx')
resume_pdf = os.path.join(directory_path, f'{Name} - Resume.pdf')

final_coverletter = os.path.join(company_folder, cover_letter_file_path) 

final_application_pdf = os.path.join(company_folder, f'{Name} - Application - {company} - {job_title}.pdf')
final_resume_pdf = os.path.join(company_folder, f'{Name} - Resume - {company} - {job_title}.pdf')
final_png = os.path.join(company_folder, f'Job Posting - {company} - {job_title}.png')
final_html = os.path.join(company_folder, f'Job Posting - {company} - {job_title}.html')

cover_letter_docx_shortened  = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {role}.docx')
final_coverletter_shortened = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {role}.pdf')
final_application_pdf_shortened = os.path.join(company_folder, f'{Name} - Application - {company} - {role}.pdf')
final_resume_pdf_shortened = os.path.join(company_folder, f'{Name} - Resume - {company} - {role}.pdf')
final_png = os.path.join(company_folder, f'Job Posting - {company} - {role}.png')
final_html = os.path.join(company_folder, f'Job Posting - {company} - {role}.html')


# Process Document
doc = Document(cover_letter_template)
replace_text_in_doc(doc, "COMPANY", company)
replace_text_in_doc(doc, "JOBTITLE", job_title)
replace_text_in_doc(doc, "POSITION", role)
replace_text_in_doc(doc, "DATE", CurrentDateCoverLetter)
replace_text_in_doc(doc, "SKILL", skill_role)


# Adjust paths if needed before assigning file names
create_shortened_paths_if_needed()

# Copy the resume PDF to the new location with the new name
shutil.copy(resume_pdf, final_resume_pdf)

# save new .docx cover letter
doc.save(cover_letter_docx)

# Convert to PDF and merge
save_as_pdf(cover_letter_docx, final_coverletter)
merge_pdfs([final_coverletter, resume_pdf], final_application_pdf)

# # Screnshot job posting
# screenshot_webpage(url, final_png)

# # Save webpage as HTML
# save_webpage_as_html(url, final_html)

print("Done! Good luck!")