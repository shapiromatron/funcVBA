% MATLAB/ACSL COMMANDS

% EXPORT CONTENTS FROM COMMAND WINDOW
set @diary		% turn diary on
set @nodiary	% turn diary off
diary "C:\Temp\out.out"		% export to this file

% Run DOS Command
! %add command here

% BEEP WHEN COMPLETE
!

%Batch Run Multiple Files
for runs=[1:10]
incmd="use "+"BatchMFile_"+ctos(num2str(runs))
eval(incmd)

%Run Model
start @nocallback

%Export Results
RESULT=[T MEANCLINGKG MEANCFNGKG MEANCBSNGKGLIADJ MEANBBNGKG MEANCBNDLINGKG MEANCBNGKG MAXCLINGKG MAXCFNGKG MAXCBSNGKGLIADJ MAXBBNGKG MAXCBNDLINGKG MAXCBNGKG]
end
filename=TextIntro+"ICF_"+ctos(num2str(runs))+".out"
outcmd="save RESULT @file="+filename+"@format=ascii @separator =comma"
eval(outcmd)

%Clear Command Window
cls