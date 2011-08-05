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
        

#Run Script:
#------------------------
import os
os.chdir('C:\Temp')
execfile('RunPy.py')

# Install package using PIP (from command prompt):
#-------------------------------------------------
# cd C:\Python27\Scripts
# pip install scipy

# Run Script in DreamPie:
#-------------------------------------------------
import os
x = 'C:\Temp'
#r is raw, and it doesn't covert "\16" to some hex value
x = r'C:\Documents and Settings\16955\Desktop\Exterior RRP MC\Exterior RRP MC'
os.chdir(x)

y = 'MC_hammer.py'
execfile(y)
