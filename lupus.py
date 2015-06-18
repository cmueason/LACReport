import operator, os, xlrd, datetime, sys, re
execfile('global.py')
execfile('process.py')


# helper function
def qualify(n):
	if n<0: 	return "NEGATIVE!!!"
	if n<1: 	return "marginally"
	if n<2:		return "minimally"
	if n<10:	return "mildly"
	if n<17:	return "moderately"
	if n<30:	return "quite"
	return "markedly"

def reduction(before, after, cutoff=0.35):
    if isinstance(after,str):
	try:
		if float(after) == 0.0:
			return False
	except:
		return False
    if float(after)==0:
        return False
    temp = (before-after)/before
    return (temp > cutoff)

# generates the first portion of output
def first(d):
	s = []
	if d["LAPTT_R"]>d["LAPTT_U"]:	s.append("In the aPTT-based system, the initial clotting time is "+qualify(d["LAPTT_P"]) +" prolonged at {:0.1f} seconds.".format(d["LAPTT_R"]))
	else:				s.append("In the aPTT-based system, the initial clotting time is normal.")

	if d["LAPTT_R"]>d["LAPTT_U"]-1.5:
		prefix = "In the mixing phase of the aPTT-based system, the clotting time "
		if d["PTTMX_R"]>d["LAPTT_R"]: 	s.append(prefix + "actually prolongs slightly, and this prolongation is {:0.1f} seconds beyond the upper limit of the normal reference interval for this phase of testing.".format(d["PTTMX_P"]))
                elif d["PTTMX_P"] == 0:         s.append(prefix + "is now within the normal reference interval.")
		elif d["PTTMX_P"] <= 1.0:       s.append(prefix + "nearly corrects into the normal range, and remains prolonged by only {:0.1f} seconds beyond the upper reference limit for that phase of testing.".format(d["PTTMX_P"]))
                elif reduction(d["LAPTT_R"], d["PTTMX_R"]):   s.append(prefix + "shortens dramatically, but still remains prolonged by {:0.1f} seconds beyond the upper normal limit for that phase of testing.".format(d["PTTMX_P"]))
                else:  	s.append(prefix + "shortens, but still remains prolonged by {:0.1f} seconds beyond the upper normal limit for that phase of testing.".format(d["PTTMX_P"]))

		if d["PNP_R"]>0:
			if d["PNP_R"]<d["PNP_R_LYS"]:	s.append("In the confirmatory phase of the aPTT-based system (platelet neutralization procedure), the clotting time does not shorten in the presence of platelet lysate.")
			elif d["PNP_SD"]>=d["PNP_U"]:   s.append("In the confirmatory phase of the aPTT-based system (platelet neutralization procedure), there is significant shortening of the clotting time in the presence of platelet lysate ({:0.1f} SD, upper limit of normal {:0.1f} SD).".format( d["PNP_SD"],d["PNP_U"]))
			else:	s.append("In the confirmatory phase of the aPTT-based system (platelet neutralization procedure), there is shortening of the clotting time, but does not exceed the significant shortening of the clotting time in the presence of platelet lysate ({:0.1f} SD, upper limit of normal {:0.1f} SD).".format( d["PNP_SD"],d["PNP_U"]))
		
		if str(d["LTT_R"])!="":
			if d["LTT_R"]<d["LTT_L"]:	s.append("The thrombin time is short at {:0.1f} seconds (lower limit of normal is {:0.1f} seconds), a finding suggestive of possibly increased fibrinogen.".format(d["LTT_R"],d["LTT_L"]))
			elif d["LTT_R"]>d["LTT_U"]:	s.append("The thrombin time is "+qualify(d["LTT_P"])+" prolonged at {:0.1f} seconds.".format(d["LTT_R"]))
			else:				s.append("The thrombin time is normal.")
			
			if str(d["LTTHEP_R"]) != "":
				if d["LTT_R"] < d["LTTHEP_R"]:	s.append("Following incubation of the plasma with heparinase, the thrombin time does not shorten, and remains prolonged by {:0.1f} seconds.".format(d["LTTHEP_P"]))
				else:				s.append("Following incubation of the plasma with heparinase, the thrombin time shortens by {0.1f}.".format(d["LTT_R"] - d["LTTHEP_R"]))
	return " ".join(s) 

# generates the second portion of output
def MiddleSection(testname, initR, initU, initP, mixR, mixU, mixP, corrR, percentR, percentU):
	s = []
	if initR > initU:	s.append("In the " + testname + " system, the initial clotting time is "+qualify(initP)+" prolonged at {:0.1f} seconds.".format(initR))
	else:	
		s.append("In the " + testname + " system, the initial clotting time is normal.")
		return " ".join(s)

	if mixR>mixU:
		prefix = "In the mixing phase of the " + testname + " system, the clotting time "
		suffix = " seconds beyond the upper reference limit for that phase of testing."
		if mixP <= 1.0:                       s.append(prefix + "nearly corrects into the normal range, and remains prolonged by only {:0.1f}".format(mixP) + suffix)
                elif reduction(initR, mixR):          s.append(prefix + "shortens dramatically, but still remains prolonged by {:0.1f}".format(mixP) + suffix)
                elif abs(initR - mixR)<1.0:           s.append(prefix + "is nearly unchanged and remains prolonged by {:0.1f}".format(mixP) + suffix)
		else:		                      s.append(prefix + "shortens by {:0.1f} seconds, but still remains prolonged by {:0.1f}".format(initR - mixR,mixP) + suffix)

		prefix = "In the confirmatory phase, using high-concentration phospholipid, "
		suffix = " ({:0.0f}%, upper limit of normal {:0.0f}%).".format(percentR*100,percentU*100)
		if corrR > initR:			s.append(prefix + "there is no shortening of the clotting time.")
		elif percentR >= percentU+0.005:	s.append(prefix + "there is a significant shortening of the clotting time" + suffix)
	        elif abs(percentR - percentU) <0.005:   s.append(prefix + "the clotting time shortens and does come quite close to the 99th percentile upper limit of the normal reference interval" + suffix)
		elif percentR <= 0.01499:		s.append(prefix + "the clotting time shortens only minimally, but certainly does not approach anywhere close to the 99th percentile upper limit of the normal reference interval" + suffix)
		elif (percentR >0.01499) and (percentR <= 0.05):  s.append(prefix + "the clotting time shortens, but does not approach the 99th percentile upper limit of the normal reference interval" + suffix)
		else:			s.append("In the confirmatory phase, using high-concentration phospholipid, the clotting time shortens, but does not exceed the 99th percentile upper limit of the normal reference interval" + suffix)
	else:	s.append("In the mixing phase of the " + testname + " system, the clotting time is normal.")	
	return " ".join(s)

def header(d,filename):
	try:
		s = []
		s.append(d["PT_INIT"])
		s.append("%0d"%d["PT_MRN"])
		(a,b,c,d,e,f) = xlrd.xldate_as_tuple(d["DATE"],0)
		s.append("%02d%02d%02d" % (a-2000,b,c))
		return " ".join(s)
	except:
		return filename.replace(".xls","")
		
def conclusionAlgorithm(d):
    (drug, LAPTT, DRVVT, DPT) = (d["DRUG"], d["PNP_SD"], d["PCTCO_R"], d["DPTCOR_R"])
    # warfarin, heparin, enoxaparin,...
    mydrugs = ["heparin", "warfarin", "enoxaparin", "argatroban", "lmwh", "lovenox", "fondaparinux", "aspirin", "rivaroxaban"]
    drug = drug.lower()
    for i in mydrugs:  d[i] = (i in drug)
    s = []
    Case = "" 

    ### helper function ###
    def helper(C1, C2, C3):
        if C1 and C2 and C3:           return "all three testing systems"
        if C1 and C2 and not C3:       return "the aPTT- and DRVVT-based systems"
        if C1 and not C2 and C3:       return "the aPTT- and DPT-based systems"
        if C1 and not C2 and not C3:   return "the aPTT-based system"
        if not C1 and C2 and C3:       return "the DRVVT- and DPT-based systems"
        if not C1 and C2 and not C3:   return "the DRVVT-based system"
        if not C1 and not C2 and C3:   return "the DPT-based system"
	return ""

    def sumConditions(C1, C2, C3):
	return reduce(operator.add, map(lambda x: 1 if x else 0, [C1,C2,C3]), 0)

    (C1,C2,C3) = (LAPTT >= 6, DRVVT >= 0.20, DPT >= 0.20)
    if (sumConditions(C1,C2,C3)>=1) and (Case==""):
        Case="STRONG"
        s.append("The current testing, notably in " + helper(C1,C2,C3) + ", provides strong evidence supporting the presence of a functional lupus anticoagulant. Confirmatory repeat testing in 12 weeks is recommended. ")

    (C1,C2,C3) = (LAPTT >= 3.05, DRVVT >= 0.135, DPT >= 0.115)
    if (sumConditions(C1,C2,C3)>=2)  and (Case==""):
        Case = "POSITIVE"
        s.append("The current testing, notably in " + helper(C1,C2,C3) + ", provides evidence supporting the presence of a functional lupus anticoagulant. Confirmatory repeat testing in 12 weeks is recommended. ")
    
    (C1,C2,C3) = ((LAPTT >= 2.5) and (LAPTT < 3.05), (DRVVT >= 0.095) and (DRVVT < 0.135), (DPT >= 0.095) and (DPT < 0.115))
    if (sumConditions(C1,C2,C3)>=1)  and (Case == ""):
        Case = "BORDERLINE"
        s.append("The current testing, notably in "+helper(C1,C2,C3)+", reaches only to borderline significance for the identification of a functional lupus anticoagulant. ")

    # fondaparinux
    if d["fondaparinux"]:
        s.append("However, interpretation of subsequent studies to evaluate the continuing presence of a lupus anticoagulant were inconclusive due to the presence of fondaparinux, which can contribute to false positive results in the DRVVT-based system.")

    if d["rivaroxaban"]:
	s.append("Review of this patient's chart indicates that he is presently receiving the anti-Xa drug, rivaroxaban.  Recent literature demonstrates that this anticoagulant exerts a greater inhibitory effect upon the initial phase of the DRVVT as opposed to its effect upon the confirmatory phase, and accordingly rivaroxaban by itself is capable of producing a false positive pattern mimicking that of a lupus anticoagulant in the DRVVT-based testing system. Thus, we are unable based upon the present studies alone to determine how much of the abnormality seen in the functional testing is simply an artifact of the rivaroxaban anticoagulation, and how much might be resulting from the actual presence of a lupus anticoagulant.  Repeat DRVVT-based testing at such time that the patient is no longer receiving rivaroxaban may be considered.")

    (C1,C2,C3) = ((LAPTT == 0), (DRVVT == 0), (DPT == 0))
    if (sumConditions(C1,C2,C3)==3)  and Case=="":
        s.append("The results of this study provide no evidence to suggest the identification of a functional lupus anticoagulant.")
        Case = "STRONG NEGATIVE"    

    (C1, C2, C3) = ((LAPTT < 2.5), (DRVVT < 0.95), (DPT < 0.95))
    if (sumConditions(C1,C2,C3)==3)  and Case=="":
        s.append("This study does not provide sufficient evidence for the identification of a functional lupus anticoagulant.")
        Case  = "NEGATIVE"

    (C1,C2,C3) = (reduction(d["LAPTT_R"],d["PTTMX_R"], 0.18), reduction(d["DRVVS_R"],d["DRVVMX_R"], 0.18), reduction(d["DPTS_R"],d["DPTMX_R"],0.18))
    if d["warfarin"]:
        if (sumConditions(C1,C2,C3)>1):  s.append("\n\nNote: The marked prolongations of the initial clotting times seen in " + helper(C1, C2, C3) + ", together with substantial shortening in the mixing phases, appear consistent with the patient's history of anticoagulation with warfarin.") 
        else:  s.append("Note: The marked prolongation of the initial clotting time seen in " + helper(C1, C2, C3) + ", together with substantial shortening in the mixing phase, appears consistent with the patient's history of anticoagulation with warfarin.") 
    else:
	if (sumConditions(C1,C2,C3)>=1):
            s.append("\n\nNote: The prolonged initial clotting times with shortening upon mixing with normal plasma in " + helper(C1, C2, C3) + " suggests the deficiency of one or more clotting factors. This would appear consistent with the increased INR noted in the patient's chart. Diet, poor vitamin K absorption, or the effect of antibiotics upon guy bacteria all could be possible underlying causes. In addition, we have noted a slight prolongation of the patient's thrombin time that appears unrelated to any heparin effect. An abnormality in fibrinogen or perhaps an increase in fibrin split products could be among the possible explanations for this finding.") 
    return " ".join(s)
    
def conclusion(d, filename):
	s = ["Conclusion:"]
	s.append(conclusionAlgorithm(d))
	for i in findFiles(filename):
		concl = extractConclusion(i)
		concl = concl.replace("Conclusion:","")
		s.append("\n\nPrevious conclusion (" + i + "): " + str(concl).replace("u'","").replace("'",""))
	return " ".join(s)

def footer(attending):
	today = datetime.date.today()
	formatdate = today.strftime('%B ')+str(today.day)+", "+str(today.year)
	if attending == "miller": 	return "Jonathan Miller, MD, PhD, Director of Coagulation Laboratory\t\t" + formatdate
	elif attending == "treml":	return "Angela Treml, MD, Attending, Coagulation Laboratory\t\t" + formatdate
	elif attending == "wool":	return "Geoffrey Wool, MD, PhD, Attending, Coagulation Laboratory\t\t" + formatdate
	elif attending == "sandeep":	return "Sandeep Gurbuxani, MD, Attending, Coagulation Laboratory\t\t" + formatdate
	else:				return "Jonathan Miller, MD, PhD, Director of Coagulation Laboratory\t\t" + formatdate

def report(d,filename, attending):
	DRVVTSection = MiddleSection("DRVVT-based", d["DRVVS_R"], d["DRVVS_U"], d["DRVVS_P"], d["DRVVMX_R"],d["DRVVMX_U"],d["DRVVMX_P"], d["DRVVC_R"], d["PCTCO_R"], d["PCTCO_U"])
	DPTSection = MiddleSection("DPT-based", d["DPTS_R"], d["DPTS_U"], d["DPTS_P"], d["DPTMX_R"],d["DPTMX_U"],d["DPTMX_P"], d["DPTC_R"],d["DPTCOR_R"], d["DPTCOR_U"])
	return header(d, filename) + "\n\n" + first(d) + "\n\n" + DRVVTSection + "\n\n" + DPTSection + "\n\n" + conclusion(d, filename) + "\n\n\n\n\n" + footer(attending)

def processXLS(filename, attending):
	myhash = readFile(filename)
	s = report(myhash,filename, attending)
	print s
	writeFile(s,os.path.join(OUTPUTDIR, os.path.basename(filename.replace(".xls", ".docx"))))

if __name__ == "__main__":
	if len(sys.argv)<2:
		print "Usage: "+sys.argv[0]+" filename.xlsx + miller|treml|sandeep|wool"
		exit(0)
	attending = "miller" if len(sys.argv) <= 2 else sys.argv[2]		
	filename = sys.argv[1]
	processXLS(filename, attending)

