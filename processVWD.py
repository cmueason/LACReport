#!/Users/ernestchan/bin/python

from xlrd import open_workbook
from docx import Document
from docx.shared import Inches
import glob, os
import re


# this file has a function that processes one file, and returns a hash with all the values


def myround100(x):
	try:
		return round(x*100)
	except:
		return (-1)

def myround(x):
	try:
		return round(x)
	except:
		return (-1)

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
 
	d["ATGR"] = myround(sheet.cell_value(12,1))
	d["RCOR"] = myround(sheet.cell_value(13,1))
	d["CBAR"] = myround(sheet.cell_value(14,1))

	d["F8L"] = myround100(sheet.cell_value(11,2))
	d["ATGL"] = myround100(sheet.cell_value(12,2))
	d["RCOL"] = myround100(sheet.cell_value(13,2))
	d["CBAL"] = myround100(sheet.cell_value(14,2))

	d["F8U"] = myround100(sheet.cell_value(11,3))
	d["ATGU"]= myround100(sheet.cell_value(12,3))
	d["RCOU"] = myround100(sheet.cell_value(13,3))
	d["CBAU"] = myround100(sheet.cell_value(14,3))

	d["F8O"] = myround100(sheet.cell_value(11,5))
	d["ATGO"]= myround100(sheet.cell_value(12,5))
	d["RCOO"] = myround100(sheet.cell_value(13,5))
	d["CBAO"] = myround100(sheet.cell_value(14,5))

	return d

def writeFile(txt, filename):
	document = Document()
	document.add_paragraph(txt)	
	document.save(filename)


