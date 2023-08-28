from PyPDF2 import PdfReader
from docx import Document

# Initialize a Word Document
doc = Document()

# Open the PDF file
with open('filepath.pdf', 'rb') as pdf_file:
    pdf_reader = PdfReader(pdf_file)
    
    # Loop through each page in the PDF
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        
        # Extract text from the page
        text = page.extract_text()
        
        # Add the text to the Word document
        doc.add_paragraph(text)
        
        # Optionally, add a page break after each PDF page
        doc.add_page_break()

# Save the Word document
doc.save('filepath.docx')
