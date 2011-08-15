#Stop a script
#------------------------
break

#Run Script:
#------------------------
import os
os.chdir(r'C:\Temp')
execfile('RunPy.py')

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