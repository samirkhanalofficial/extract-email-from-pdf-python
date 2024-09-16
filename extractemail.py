import PyPDF2
import re
from docx import Document


def extract_emails_from_pdf(pdf_path):
    emails = set()
    pdf_reader = PyPDF2.PdfReader(open(pdf_path, 'rb'))
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        emails.update(re.findall(
            r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', page_text))
    return emails


def save_emails_to_docx(emails, docx_path):
    doc = Document()
    doc.add_paragraph(', '.join(emails))
    doc.save(docx_path)


pdf_path = input("File name of PDF (must be placed in same foder) : ")
docx_path = 'emails.docx'
emails = extract_emails_from_pdf(pdf_path)
save_emails_to_docx(emails, docx_path)
print(f"Emails extracted and saved to {docx_path}")
