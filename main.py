from pdf2image import convert_from_bytes
import os
import zipfile

class pdfTools:
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        
    def export_as_png(self):
        images = convert_from_bytes(open(self.pdf_file, 'rb').read())
        if not os.path.exists(self.pdf_file[:-4]):
            os.makedirs(self.pdf_file[:-4])
        for i, image in enumerate(images):
            try:
                image.save(f"{self.pdf_file[:-4]}/{i+1}.png", 'PNG')
                print(f"Page {i} of the PDF saved as {self.pdf_file[:-4] + '/' + str(i+1) + '.png'}")
            except Exception as e:
                print(f"Error: {e}")
        zipf = zipfile.ZipFile(f"{self.pdf_file[:-4]}.zip", 'w', zipfile.ZIP_DEFLATED)
        for root, dirs, files in os.walk(self.pdf_file[:-4]):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), os.path.join(self.pdf_file[:-4], '..')))
        zipf.close()
        os.rename(f"{self.pdf_file[:-4]}.zip", f"{self.pdf_file[:-4]}/{self.pdf_file[:-4]}.zip")
        print(f"All pages of the PDF saved as {self.pdf_file[:-4]}/{self.pdf_file[:-4] + '.zip'}")

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
        tools.export_as_png()
        tools.export_first_page_as_png()