#!/Users/ernestchan/bin/python
import os
import lupus
import sys

execfile('global.py')

attending = "miller"
if len(sys.argv) >= 2:
	attending = sys.argv[1]

for filename in os.listdir(RUNDIR):
	if filename.endswith(".xls"):
		lupus.processXLS(os.path.join(RUNDIR, filename), attending);

