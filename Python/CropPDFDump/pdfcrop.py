import os
from os import listdir
from os.path import isfile, join
from pathlib import Path
from PyPDF2 import PdfFileWriter, PdfFileReader
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
from datetime import datetime

if not os.path.exists("PDFs"):
    print("PDFs folder does not exist.  Please create a folder called 'PDFs and place that in the containing folder.'")
    quit()
else:
    print("PDFs folder exists")

if not os.path.exists("CroppedImages"):
    print("Creating CroppedImages Folder")
    os.mkdir("CroppedImages")
else:
    print("CroppedImages folder exists")

path = "PDFs"
onlyfiles = [f for f in listdir(path) if isfile(join(path, f))]

for file in onlyfiles:
    print(file)
    pathfile = "PDFs/" + file

    pdf_reader = PdfFileReader(str(pathfile), strict=False)
    numberPages = pdf_reader.numPages

    for x in range(numberPages):
        first_page = pdf_reader.getPage(x)

        first_page.mediaBox.upperLeft = (0, 480)
        first_page.mediaBox.lowerRight = (150, 400)

        fileExt = "_Cropped_Page" + str(x + 1) + ".pdf"
        croppedfilename = file.replace('.pdf', fileExt)
        outpath = "CroppedImages/" + croppedfilename
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(first_page)
        with Path(outpath).open(mode="wb") as output_file:
            pdf_writer.write(output_file)

        images = convert_from_path(outpath)
        for i in range(len(images)):
            images[i].save(outpath.replace('.pdf', '.jpg'))

        if os.path.exists(outpath):
            os.remove(outpath)
        else:
            print("This file does not exist")


document = Document()

path = "CroppedImages"
croppedImageFiles = [f for f in listdir(path) if isfile(join(path, f))]

for cropimages in croppedImageFiles:
    document.add_paragraph(cropimages)
    imgPath = "CroppedImages/" + cropimages
    document.add_picture(imgPath, width=Inches(5))

dateTimeObj = datetime.now()
documentName = dateTimeObj.strftime("%Y%m%d") + "_PDFImageDump.docx"
document.save(documentName)
