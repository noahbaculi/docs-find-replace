from PyPDF2 import PdfReader, PdfFileWriter

replacements = [("position", "LALALA")]

pdf = PdfReader("cover_letter_test.pdf", "rb")
writer = PdfFileWriter()

# breakpoint()
# dir(pdf.pages[0].get_contents()[0].get_object())


for page in pdf.pages:
    contents = page.get_contents()[0].get_object().get_data()
    for (a, b) in replacements:
        contents = contents.replace(a.encode("utf-8"), b.encode("utf-8"))
    breakpoint()
    print(page.getContents()[0].get_object().decoded_self)
    page.getContents()[0].get_object().decoded_self.set_data(contents)
    writer.addPage(page)

with open("modified.pdf", "wb") as f:
    writer.write(f)
