from pypdf import PdfReader, PdfWriter
from pypdf.annotations import FreeText

pdf_path = r"C:\Users\Angelo\Desktop\Holiday Stuff\Automation\sources_pdfs\pefs.pdf"
reader = PdfReader(pdf_path)
page = 0
writer = PdfWriter()
pages = reader.pages[0]
writer.add_page(pages)
annotation = FreeText(
                    text="rattus rattus",
                    rect=(100, 0, 200, 200), # look for a way to conver from pixel space to document space
                    font="Arial",              
                    bold=False,
                    italic=False,
                    font_size="8pt",
                    font_color="#ff0000",
                    border_color="#ff0000",
                    background_color="#FFFFFF"
                )
annotation.flags = 4
writer.add_annotation(page_number=0, annotation=annotation)

    # Write the annotated file to disk
with open("annotated-pdf.pdf", "wb") as fp:
    writer.write(fp)