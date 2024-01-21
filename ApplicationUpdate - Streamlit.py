### This is still a work in progress. Please use the ApplicationUpdate.py

## To run, run "streamlit run CodeUpdater-Streamlit.py" in terminal ##


import streamlit as st
import os
from docx import Document
from PyPDF2 import PdfMerger
from datetime import datetime
import subprocess
import comtypes.client  # word to PDF
import altair as alt

def replace_text_in_doc(doc, old_text, new_text):
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text

def save_as_pdf(word_file, pdf_file): # Requires Microsoft Word installed on computer for this to work
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


def main():
    st.title("Cover Letter Updater")

    # User Input
    company = st.text_input("Enter the company name:")
    position = st.text_input("Enter the position:")
    submit_button = st.button("Update Cover Letter")

    if submit_button and company and position:
        # Input parent directory path
        directory_path = r'\\SepehrNAS\Thick Volume\CAREER\SEPEHR\Job Related\Resume & Cover Letter\APPLICATIONS\2024'

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

        # Process Document
        doc = Document(cover_letter_template)
        replace_text_in_doc(doc, "COMPANY", company)
        replace_text_in_doc(doc, "POSITION", position)
        replace_text_in_doc(doc, "DATE", current_date)
        doc.save(temp_docx)

        # Convert to PDF and merge
        save_as_pdf(temp_docx, temp_pdf)
        merge_pdfs([temp_pdf, resume_pdf], final_pdf)

        st.success("Done! Your cover letter and resume have been updated and merged successfully.")

if __name__ == "__main__":
    main()
