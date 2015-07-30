#!/Users/ernestchan/bin/python
import os
import xlrd
import datetime
import sys
import re
execfile('processVWD.py')

# helper function
def qualify(n):
	if n<0:
		return "NEGATIVE!!!"
	if n<4: 
		return "marginally"
	if n<9:
		return "minimally"
	if n<19:
		return "mildly"
	if n<40:
		return "moderately"
	if n<80: 
		return "quite"
	return "markedly"

def reduction(before, after, cutoff=0.35):
    if after==0:
        return False
    temp = (before-after)/before
    return (temp > cutoff)

def header(d,filename):
	try:
		s = []
		s.append(d["PT_INIT"])
		s.append("%0d"%d["PT_MRN"])
		(a,b,c,d,e,f) = xlrd.xldate_as_tuple(d["DATE"],0)
		s.append("%02d%02d%02d" % (a-2000,b,c))
		return " ".join(s)
	except:
		return os.path.basename(filename).replace(".xlsx","").replace(".xls","")
		
def conclusion(d, filename):
	s = []
	
	RCOU = d["RCOU"]
 	RCOR = d["RCOR"]
	CBAU = d["CBAU"]
	CBAL = d["CBAL"]
	F8U = d["F8U"]
	ATGL = d["ATGL"]
	F8R = d["F8R"]
	ATGR = d["ATGR"]
	F8L = d["F8L"]
	CBAR = d["CBAR"]
	ATGU = d["ATGU"]
	DATE = d["DATE"]
	RCOL = d["RCOL"]
	RCOO = d["RCOO"]
	ATGO = d["ATGO"]
	CBAO = d["CBAO"]
	F8O = d["F8O"]


	def helper(title, R, L, U, prev=""):
		def helper2(prev, now):
			if prev == now:
				return "again "
			else:
				return ""
		t = []
		result = ""
		if R / U > 1.3:
			result = "V HIGH ABNORMAL"
			t.append("The "+title+" is "+helper2(prev,result)+"significantly increased at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(R, L, U))
		elif R / U > 1.0:
			result = "V HIGH ABNORMAL"
			t.append("The "+title+" is "+helper2(prev,result)+qualify(R-U)+" increased at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(R, L, U))
		elif R < L:
			result = "LOW ABNORMAL"
			t.append("The "+title+" is "+helper2(prev,result)+"below the lower limit of the normal reference range at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(R, L, U))
		elif R - L <= 7:
			result = "LOW NORMAL"
			t.append("The "+title+" is "+helper2(prev,result)+"within the normal reference range, but close to the lower limit of normal at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(R, L, U))
		else:
			result = "NORMAL"
			t.append("The "+title+" is "+helper2(prev,result)+"normal at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(R, L, U))

		return [" ".join(t), result]

	# vwd antigen
	(sentence, result) = helper("von Willebrand factor (vWF) antigen level", ATGR, ATGL, ATGU)
	s.append(sentence)	

	# vwd ristocetin cofactor
	(sentence, result) = helper("vWF ristocetin cofactor activity", RCOR, RCOL, RCOU, result)	
	s.append(sentence)	

	# vwd cba
	(sentence, result) = helper("vWF collagen-binding activity", CBAR,CBAL,CBAU, result)	
	s.append(sentence)	

	if (RCOO > 60) and (CBAO > 60) and (ATGO > 60):
		s.append("Ratios of each of the von Willebrand factor functional activities to von Willebrand factor antigen are both normal.")

	# Factor 8
	if F8R > F8U:
		s.append("\n\nThe Factor VIII level is also "+qualify(F8R-F8U)+" increased at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(F8R,F8L,F8U))
	elif (F8R < F8U) and (F8R > F8L):
		s.append("The Factor VIII level is normal at {:0.0f}% (normal range {:0.0f}% to {:0.0f}%).".format(F8R, F8L, F8U))

	s.append("\n\nConclusion:")

	# conclusion
	if min(ATGO, CBAO, RCOO) >= 60:
		s.append("The current results provide no evidence for a diagnosis of von Willebrand disease.")
	else:
		a =  min(ATGR-ATGL, CBAR-CBAL, RCOR-RCOL)  
		b =  max(ATGR-ATGL, CBAR-CBAL, RCOR-RCOL)  
		# there might be disease!
		if (a >= 0) and (b <=15):
			s.append("The results in this study are insufficient to support a diagnosis of von Willebrand disease. That said, we note that the values are all in the lower region of the respective normal reference intervals. Since vWF and Factor VIII are well known to vary over time in their levels, and for example to be increased during times of stress, we would uggest that these assays be repeated once or possibly even twice more in the coming months, if the etiology of the patient's clinical bleeding remains unclear. In the event that the patient's basal levels of vWF/Factor VIII are actually lower, and we could possibly now be measuring them at their relatively high values, this would constitute important information.")
		elif (b<0):
			s.append("The current results provide support for a diagnosis of von Willebrand disease.")

	if ATGR > ATGU + 20:
		if F8R > F8U + 20:
			s.append("The increased levels of both vWF and of Factor VIII would appear likely to represent acute phase reactants.")
		else:
			s.append("The increased levels of vWF would appear likely to represent acute phase reactants.")
	else:
		if F8R > F8U + 20:
			s.append("The increased levels of Factor VIII would appear likely to represent acute phase reactants.")
		else:
			pass	



	return " ".join(s)

def footer(d,attending):
	today = datetime.date.today()
	formatdate = today.strftime('%B ')+str(today.day)+", "+str(today.year)
	if attending == "miller": 	return "Jonathan Miller, MD, PhD, Director of Coagulation Laboratory\t\t" + formatdate
	elif attending == "treml":	return "Angela Treml, MD, Attending, Coagulation Laboratory\t\t" + formatdate
	elif attending == "wool":	return "Geoffrey Wool, MD, PhD, Attending, Coagulation Laboratory\t\t" + formatdate
	elif attending == "sandeep":	return "Sandeep Gurbuxani, MD, Attending, Coagulation Laboratory\t\t" + formatdate
	else:				return "Jonathan Miller, MD, PhD, Director of Coagulation Laboratory\t\t" + formatdate

# takes in a hashtable, generates output
def report(d,filename, attending):
	s0 = header(d,filename)
	s4 = conclusion(d, filename)
	s5 = footer(d,attending)
	return s0+"\n\n"+s4+"\n\n\n\n\n"+s5

def processXLS(filename, attending):
	# read a file
	myhash = readFile(filename)
	s = report(myhash,filename, attending)
	print s

	# write to file
	writeFile(s,filename.replace(".xls", ".doc"))

# look at arguments
#print 'Number of arguments:', len(sys.argv), 'arguments.'
#print 'Argument List:', str(sys.argv
if __name__ == "__main__":
	if len(sys.argv)<3:
		print "Usage: "+sys.argv[0]+" filename.xlsx miller|treml|sandeep|wool"
		print "Output: filename.docx\n"
		exit(0)
	else:
		attending = "miller" if len(sys.argv) <= 2 else sys.argv[2]		
		filename = sys.argv[1]
		processXLS(filename, attending)


