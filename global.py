#MAC
windows = False 
testing = False

if (not windows):
	REPORTDIR = "/Users/ernestchan/Desktop/reports/"
	RUNDIR    = "/Users/ernestchan/Desktop/Temp"
	OUTPUTDIR = "/Users/ernestchan/Desktop/Temp"

#WINDOWS
if (windows):
	if (not testing): 
		REPORTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\5. Final Report Archive\\"
		RUNDIR    = "N:\\Coag\\Lupus\\Lupus Reports\\3. Need Pathologist Review"	
		OUTPUTDIR = RUNDIR
	else:
		REPORTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\5. Final Report Archive\\"
		RUNDIR    = "N:\\Coag\\Lupus\\Lupus Reports\\3. Need Pathologist Review"	
		OUTPUTDIR = "N:\\Coag\\Lupus\\Lupus Reports\\6. Archived Auto Drafts\\test"	
	



