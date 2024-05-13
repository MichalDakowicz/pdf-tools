from pdf2image import convert_from_bytes
import os

class pdfTools:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file

    def export_first_page_as_png(self):
        images = convert_from_bytes(open(self.pdf_file, 'rb').read(), first_page=1, last_page=1)
        try:
            images[0].save(f"{self.pdf_file[:-4]}.png", 'PNG')
            print(f"First page of the PDF saved as {self.pdf_file[:-4] + '.png'}")
        except Exception as e:
            print(f"Error: {e}")

for file in os.listdir():
    if file.endswith(".pdf"):
        tools = pdfTools(file)
        tools.export_first_page_as_png()