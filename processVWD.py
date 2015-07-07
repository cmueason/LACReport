#!/Users/ernestchan/bin/python

from xlrd import open_workbook
from docx import Document
from docx.shared import Inches
import glob, os
import re


# this file has a function that processes one file, and returns a hash with all the values

def readFile(filename):
	book = open_workbook(filename,on_demand=True)
	for name in book.sheet_names():
    		if name.endswith('Lab'):
			sheet = book.sheet_by_name(name)
	d = dict()
	d["PT_INIT"] = sheet.cell_value(3, 6)
	d["PT_MRN"] = sheet.cell_value(4, 6)
	d["DATE"] = sheet.cell_value(5, 6)

	try:
		d["F8R"] = round(sheet.cell_value(11,1))
	except:
		d["F8R"] = -1 
 
	d["ATGR"] = round(sheet.cell_value(12,1))
	d["RCOR"] = round(sheet.cell_value(13,1))
	d["CBAR"] = round(sheet.cell_value(14,1))

	d["F8L"] = round(sheet.cell_value(11,2) * 100)
	d["ATGL"] = round(sheet.cell_value(12,2) * 100)
	d["RCOL"] = round(sheet.cell_value(13,2) * 100)
	d["CBAL"] = round(sheet.cell_value(14,2) * 100)

	d["F8U"] = round(sheet.cell_value(11,3) * 100)
	d["ATGU"]= round(sheet.cell_value(12,3) * 100)
	d["RCOU"] = round(sheet.cell_value(13,3) * 100)
	d["CBAU"] = round(sheet.cell_value(14,3) * 100)

	d["F8O"] = round(sheet.cell_value(11,5) * 100)
	d["ATGO"]= round(sheet.cell_value(12,5) * 100)
	d["RCOO"] = round(sheet.cell_value(13,5) * 100)
	d["CBAO"] = round(sheet.cell_value(14,5) * 100)

	return d

def writeFile(txt, filename):
	document = Document()
	document.add_paragraph(txt)	
	document.save(filename)


