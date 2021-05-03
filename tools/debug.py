import os
from shutil import copyfile, rmtree
import json
import docx2txt
import openpyxl as oxl
from docxtpl import DocxTemplate
import win32com.client

data = docx2txt.process("t.docx")
print(data)

wordCerti = "temp.docx"
document = DocxTemplate("t.docx")
entry = {"participant_name":"Deekshant Wadhwa"}
document.render(entry)
document.save("temp.docx")
