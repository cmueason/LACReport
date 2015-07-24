import os
import lupus
import sys
import time

execfile('global.py')

before = dict ([(f, None) for f in os.listdir (WATCHDIR)])
while 1:
  after = dict ([(f, None) for f in os.listdir (WATCHDIR)])
  added = [f for f in after if not f in before]
  removed = [f for f in before if not f in after]
  if added: 
        print "Added: ", ", ".join (added)
        attending = "miller"
        for filename in os.listdir(WATCHDIR):
	        if filename.endswith(".xls"):
		        lupus.processXLS(os.path.join(WATCHDIR, filename), attending); 

  for filename in os.listdir(RUNDIR):
        if filename.endswith(".xls"):
                lupus.processXLS(os.path.join(RUNDIR, filename), attending);i
  #if removed: print "Removed: ", ", ".join (removed)
  before = after
  time.sleep(10)



