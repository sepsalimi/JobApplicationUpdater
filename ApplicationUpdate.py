#region Imports
import os, time, shutil
import comtypes.client # word to PDF
from docx import Document
from PyPDF2 import PdfMerger
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.edge.service import Service
# import subprocess
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
#endregion

##################################
########## Instructions ########## 

## 1) 
# Anywhere in cover letter that is COMPANY's, re-write as COMPS. This will handle companies that end with 's'.
## ie. which aligns well with COMPS culture

## 2) 
# Input the following information:
Name = 'Sepehr Salimi'

# Input parent directory path where cover letter and resume exists
directory_path = r'\\SepehrNAS\Thick Volume\CAREER\SEPEHR\Job Related\Resume & Cover Letter\APPLICATIONS\2024'

## 3) 
# Name cover letter template as {Name} - Cover Letter.docx and place in parent directory to match the following:
cover_letter_template_name = f'{Name} - Cover Letter.docx'

# Name resume as {Name} - Resume.pdf and place in parent directory to match the following:
resume_pdf_name = f'{Name} - Resume.pdf'

# Only for screenshots of job posting: specify the path to Edge WebDriver executable
# edge_driver_path = r'C:\Users\sepeh\OneDrive\Documents\Git\CoverLetterUpdater\msedgedriver.exe'

##################################

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
    global cover_letter_file_path, cover_letter_docx, final_coverletter, final_application_pdf, final_resume_pdf
    if len(cover_letter_file_path) > 83 or ('/' in cover_letter_file_path):
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
    
def request_file_name():
    return input(f"Company or title name incompatible with Windows. What do you want to save the file as? \n{Name} - Cover Letter - ____________________: ")

def main_process():
    try:
        save_as_pdf(cover_letter_docx, final_coverletter)
        merge_pdfs([final_coverletter, resume_pdf], final_application_pdf)
    except Exception as e:
        print(f"An error occurred: {e}")
        new_file_name = request_file_name()
    
def replace_text(doc, company, job_title, role, CurrentDateCoverLetter, skill_role):
       
    # Correct the possessive form of the company name depending on if last letter is an 's'
    if company[-1].lower() == 's':
        company_possessive = f"{company}'"
        replace_text_in_doc(doc, "COMPS", company_possessive)
    else:
        company_possessive = f"{company}'s"
        replace_text_in_doc(doc, "COMPS", company_possessive)
      
    replace_text_in_doc(doc, "COMPANY", company)
    replace_text_in_doc(doc, "JOBTITLE", job_title)
    replace_text_in_doc(doc, "POSITION", role)
    replace_text_in_doc(doc, "DATE", CurrentDateCoverLetter)
    replace_text_in_doc(doc, "SKILL", skill_role)     
        
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

# used to check if file name length is valid
cover_letter_file_path = f'{Name} - Cover Letter - {company} - {job_title}.pdf'

# Define file paths
cover_letter_template = os.path.join(directory_path, cover_letter_template_name)
cover_letter_docx = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {job_title}.docx')
resume_pdf = os.path.join(directory_path, resume_pdf_name)

# Final file name pathways
final_coverletter = os.path.join(company_folder, cover_letter_file_path) 
final_application_pdf = os.path.join(company_folder, f'{Name} - Application - {company} - {job_title}.pdf')
final_resume_pdf = os.path.join(company_folder, f'{Name} - Resume - {company} - {job_title}.pdf')
final_png = os.path.join(company_folder, f'Job Posting - {company} - {job_title}.png')
final_html = os.path.join(company_folder, f'Job Posting - {company} - {job_title}.html')

# Shortened file name pathways
cover_letter_docx_shortened  = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {role}.docx')
final_coverletter_shortened = os.path.join(company_folder, f'{Name} - Cover Letter - {company} - {role}.pdf')
final_application_pdf_shortened = os.path.join(company_folder, f'{Name} - Application - {company} - {role}.pdf')
final_resume_pdf_shortened = os.path.join(company_folder, f'{Name} - Resume - {company} - {role}.pdf')
final_png = os.path.join(company_folder, f'Job Posting - {company} - {role}.png')
final_html = os.path.join(company_folder, f'Job Posting - {company} - {role}.html')

# Process Document
doc = Document(cover_letter_template)
replace_text(doc, company, job_title, role, CurrentDateCoverLetter, skill_role)

# Adjust paths if needed before assigning file names
create_shortened_paths_if_needed()

# Copy the resume PDF to the new location with the new name
shutil.copy(resume_pdf, final_resume_pdf)

# save new .docx cover letter
doc.save(cover_letter_docx)

# Convert to PDF and merge
main_process()

# # Screnshot job posting
# screenshot_webpage(url, final_png)

# # Save webpage as HTML
# save_webpage_as_html(url, final_html)

print("Done! Good luck!\n")