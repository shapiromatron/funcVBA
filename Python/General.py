#!/usr/bin/env python
    
#Run Script:
#------------------------
import os
os.chdir(r'C:\Temp')
execfile('RunPy.py')


#===========
#| IPYTHON |
#===========
"""
To see available methods/attributes, type your variable, then "." and press TAB

For documentation, type method and then "?"
"""

with open('C:\Temp\dump.txt'), 'w') as outfile:
    outfile.write('me mi')
    np.savetxt(outfile, GSD_adjs.T, delimiter='\t')

    
#====================================================
#| ADD TO BOTTOM OF SCRIPT TO AUTOCALL FROM STARTUP |
#====================================================
def main(argv):
    if len(argv) == 1:
        RunPnCB()
    else:
        print "usage_of_exposure: %s SOME_FOLDER" % argv[0]

if __name__ == '__main__':
    main(sys.argv)

#|-------------|
#| DEBUG TIMER |
#|-------------|
class Timer:
    def __init__(self, TimerName):
        self.TimerName = TimerName
        self.time1 = time.clock()
    def StopTime(self):
        self.time2 = time.clock()
    def PrintTime(self):
        print "%s: %f" % (self.TimerName, self.time2 - self.time1)
        
# get location of running script
import os
print os.getcwd()

# get computer name
import os
comp_name = os.getenv('COMPUTERNAME')

# get current time
def now_string():
    import datetime
    now = datetime.datetime.now()
    print now.strftime("%Y-%m-%d %H:%M")

# print Tweet completion using Tweet EXE
import datetime
now = datetime.datetime.now()
time = now.strftime("%Y-%m-%d %H:%M")
tweet_exe = r'%s/tweet.exe' % os.getcwd()
tweet = 'TRIM.FaTE runs on %s complete at %s!' % (os.getenv('COMPUTERNAME'), time)
print tweet_exe
subprocess.call([tweet_exe, tweet])

