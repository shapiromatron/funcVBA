%STARTUP:
cls	%Clear Command Window

%##########################################
%##########################################
%##########################################
%##########################################
%##########################################
%##########################################
%######## START OPTION SETTING ############
%##########################################

% Global Options
%------------------------------------------
OutputFN ="C:\Temp\AJS_test_case_1830.out"
MaxNumberOfAttempts = 20
BeepWhenComplete = 1 		% 1 = beep, 0 = no beep

% Step 1 Options:
%------------------------------------------
Step1_StartingDose = 10.0 		% initial guess (ng/kg)
Step1_TargetConc = 1830.0 		% Target CBSNGKGLIADJ (ng/kg)

% Step 2 Options:
%------------------------------------------
% <none!>

% Step 3 Options:
%------------------------------------------
Step3_Dose = 0.001 % initial guess (NG/KG)

%##########################################
%########## END OPTION SETTING ############
%##########################################
%##########################################
%##########################################
%##########################################
%##########################################
%##########################################

%turn on diary
if exist(OutputFN) == 2
	delete(OutputFN)
end
set @diary
get(@diary)
diary(OutputFN)

%Print key inputs into diary
OutputFN
MaxNumberOfAttempts
BeepWhenComplete
Step1_StartingDose
Step1_TargetConc
Step3_Dose

%===================================================================================================
%===================================================================================================
%===================================================================================================
%|||||  STEP #1 ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
%===================================================================================================
%===================================================================================================
%===================================================================================================

% reset bounds
Higher_MSTOT = 0
Lower_MSTOT = 99999999
Higher_Value = 0
Lower_Value = 99999999

%initialize starting guess
CurrentGuess = Step1_StartingDose
	
for Run=1:MaxNumberOfAttempts
	
	%###################################
	%###################################
	%###################################
	%###################################
	%###################################
	%###################################
	% Paste m file below for Step #1
	%###################################
	% Change: MSTOT = CurrentGuess
	% Add: CurrentConc = CBSNGKGLIADJ (or whatever the output variable is)
	%###################################
	%###################################


	output @clear
	%prepare @clear year T CLINGKG CFNGKG CBSNGKGLIADJ BBNGKG  CBNDLINGKG CBNGKG
	prepare @clear T CBSNGKGLIADJ CBNGKG
	%output @all 

	% PARAMETERS FOR SIMULATION
	% Simulate a 4 year old boy exposed to the Seveso incident
	% A single pulse dose at time = 0
	% Target Concentration: 26,000 ppt (ng/kg)

	MAXT = 0.5
	CINT = 1. 
	EXP_TIME_ON  = 21900.    % Delay before begin exposure (HOUR) 6.2 years
	EXP_TIME_OFF = 21923.    %324120     % HOUR/YEAR !TIME EXPOSURE STOP (HOUR) 6.2 years + 23 hours
	DAY_CYCLE    = 24.       % TIME 
	BCK_TIME_ON  = 0.        % DELAY BEFORE BACKGROUND EXP (HOUR)
	BCK_TIME_OFF = 613200.   % TIME OF BACKGROUND EXP STOP (HOUR) 
	TIMELIMIT    = 26280.    % half a year (July 1976 until January 1977) past 6.2 years
	MSTOTBCKGR   = 0.002106    % ORAL BACKGROUND EXPOSURE DOSE (UG/KG)
			
	% oral dose oral dose oral dose 
	MSTOT        = CurrentGuess   % Serveso, ORAL DAILY EXPOSURE DOSE (NG/KG/day)
	DOSEIV       = 0         % 40  %50 %5 %0.5 %0.3 %0.2 %0.1%0.05%0.3 %NG/KG
	% oral dose oral dose oral dose 

	MEANLIPID    = 731       % 711 %664 %778 %468 %671 %730 %662 %592%615%730%
	PAS_INDUC    = 1         % NON INDUCTION (0) CONTROLE DE L'INDUCTION 

	%human variable parameter 
	MALE = 1.
	FEMALE = 0.
	Y0 = 0.                  % 0 years old at the beginning of the simulation

	start @nocallback
	%CBSNGKGLIADJ
	%CBSNGKGLIADJ_oneday=mean(_cbsngkgliadj(find(_t==58668):length(_t)))
	%CBSNGKGLIADJ_twoday=mean(_cbsngkgliadj(find(_t==58644):length(_t)))
	CBSNGKGLIADJ_oneday=mean(_cbsngkgliadj(find(_t==26112):length(_t)))
	%meanCBSNGKGLIADJ=mean(_cbsngkgliadj);
	%meanCBSNGKGLIADJ

	CurrentConc = CBSNGKGLIADJ_oneday

	%###################################
	% End m file paste for Step #1
	%###################################
	%###################################
	%###################################
	%###################################
	%###################################
	%###################################

	% SEE IF CURRENT VALUE IS REACHED, IF IT IS THEN EXIT
	% Note that this is based if the values are identical in scientific notation
	% using 3 digits (X.XXE-YY)
	if (strcmp(sprintf("%.2E",CurrentConc), sprintf("%.2E",Step1_TargetConc)) == 1) 
		break
	end 

	%IF ITS NOT REACHED, THEN BEGIN GUESSING

	%SET HIGHER AND LOWER WITH NEW GUESSES
	if CurrentConc > Step1_TargetConc
		Higher_MSTOT = CurrentGuess;
		Higher_Value = CurrentConc;
	end

	if CurrentConc < Step1_TargetConc
		Lower_MSTOT = CurrentGuess;
		Lower_Value = CurrentConc;
	end

	%Now, create a new guess, by a linear interpolation
	if (Higher_MSTOT ~= 0) & (Lower_MSTOT ~= 99999999) 
		CurrentGuess = Lower_MSTOT + (Step1_TargetConc - Lower_Value)*((Higher_MSTOT-Lower_MSTOT)/(Higher_Value-Lower_Value));
	end
	
	% no guess lower yet
	if Lower_MSTOT == 99999999
		CurrentGuess = CurrentGuess * 0.1;
	end
	
	% no guess higher yet
	if Higher_MSTOT == 0
		CurrentGuess = CurrentGuess * 10.0;
	end
end

Step1_StartingDose = CurrentGuess

%Save outputs for Step #1
Output_Step1 = "Step #1 Outputs:";
Output_Step1 = sprintf("\n%s\n%s",Output_Step1,"----------------");
Output_Step1 = sprintf("%s\n%s%g",Output_Step1,"Match found on run number = ",Run);
if MaxNumberOfAttempts == Run	
	Output_Step1 = sprintf("%s\n%s",Output_Step1,"Warning! Run Count Equal to MaxNumberOfAttempts!");
end 
Output_Step1 = sprintf("%s\n%s%g",Output_Step1,"Target Concentration to Match = ",Step1_TargetConc);
Output_Step1 = sprintf("%s\n%s%g",Output_Step1,"Target Concentration Reached = ",CurrentConc);
Output_Step1 = sprintf("%s\n%s%g\n",Output_Step1,"MSTOT to reach Concentration = ",Step1_StartingDose);


%===================================================================================================
%===================================================================================================
%===================================================================================================
%|||||  STEP #2 ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
%===================================================================================================
%===================================================================================================
%===================================================================================================

%###################################
%###################################
%###################################
%###################################
%###################################
%###################################
% Paste m file below for Step #2
%###################################
% Change: MSTOT = Step1_StartingDose
%###################################
%###################################

output @clear
%output @nciout=24*30 T CBSNGKGLIADJ
%prepare @clear year T CLINGKG CFNGKG CBSNGKGLIADJ BBNGKG  CBNDLINGKG CBNGKG
%output @nciout=24 T CBSNGKGLIADJ
prepare @clear T CBSNGKGLIADJ CBNGKG
%output @all 

% PARAMETERS FOR SIMULATION
% Simulate a 4 year old boy exposed to the Seveso incident
% A single pulse dose at time = 0
% Target Concentration: 26,000 ppt (ng/kg)

MAXT = 0.5
CINT = 1. %
EXP_TIME_ON  = 21900.     % Delay before begin exposure (HOUR) 6.2 years
EXP_TIME_OFF = 21923.     %324120     % HOUR/YEAR !TIME EXPOSURE STOP (HOUR) 6.2 years + 23 hours
DAY_CYCLE    = 24.        % TIME 
BCK_TIME_ON  = 0.         % DELAY BEFORE BACKGROUND EXP (HOUR)
BCK_TIME_OFF = 613200.    % TIME OF BACKGROUND EXP STOP (HOUR) 
TIMELIMIT    = 43800.     % 10 years
MSTOTBCKGR   = 0.002106     % ORAL BACKGROUND EXPOSURE DOSE (UG/KG)

% oral dose oral dose oral dose 
MSTOT        = Step1_StartingDose         % Serveso, ORAL DAILY EXPOSURE DOSE (NG/KG/day)
DOSEIV       = 0          % 40  %50 %5 %0.5 %0.3 %0.2 %0.1%0.05%0.3 %NG/KG
% oral dose oral dose oral dose 

MEANLIPID    = 730        % 711 %664 %778 %468 %671 %730 %662 %592%615%730%
PAS_INDUC    = 1          % NON INDUCTION (0) CONTROLE DE L'INDUCTION 

%human variable parameter 
MALE = 1.
FEMALE = 0.
Y0 = 0.                % 0 years old at the beginning of the simulation

start @nocallback
%CBSNGKGLIADJ
meanCBSNGKGLIADJ=mean(_cbsngkgliadj(find(_t==EXP_TIME_ON):length(_t)));
meanCBSNGKGLIADJ
maxCBSNGKGLIADJ=max(_cbsngkgliadj);
maxCBSNGKGLIADJ

%###################################
% End m file paste for Step #2
%###################################
%###################################
%###################################
%###################################
%###################################
%###################################

Step3_MaxConc = maxCBSNGKGLIADJ;
Step3_MeanConc = meanCBSNGKGLIADJ;

%Save outputs for Step #2
Output_Step2 = "Step #2 Outputs:";
Output_Step2 = sprintf("%s\n%s",Output_Step2,"----------------");
Output_Step2 = sprintf("%s\n%s%g",Output_Step2,"Starting Dose = ",Step1_StartingDose);
Output_Step2 = sprintf("%s\n%s%g",Output_Step2,"peak CBSNGKGLIADJ = ",maxCBSNGKGLIADJ);
Output_Step2 = sprintf("%s\n%s%g\n",Output_Step2,"mean CBSNGKGLIADJ = ",meanCBSNGKGLIADJ);

%===================================================================================================
%===================================================================================================
%===================================================================================================
%|||||  STEP #3 ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
%===================================================================================================
%===================================================================================================
%===================================================================================================

Output_Step3 = "Step #3 Outputs:";
Output_Step3 = sprintf("%s\n%s\n",Output_Step3,"----------------");

% Go through this two times; once to match the max, and once for the mean
for MatchTarget=1:2
	
	% define target and value to compare against
	if (MatchTarget == 1)
		ValueToCheck = "maxCBSNGKGLIADJ";
		TargetValue = Step3_MaxConc;
	else
		ValueToCheck = "meanCBSNGKGLIADJ";
		TargetValue = Step3_MeanConc;
	end
	Step3_Threshold = TargetValue * 0.005 % how close do we need to get to target	
	
	% reset bounds
	Higher_MSTOT = 0
	Lower_MSTOT = 99999999
	Higher_Value = 0
	Lower_Value = 99999999
	
	%initialize starting guess
	CurrentGuess = Step3_Dose
	
	%now, try to find a match, from 1 to the number of match attempts
	for Run=1:MaxNumberOfAttempts

		%###################################
		%###################################
		%###################################
		%###################################
		%###################################
		%###################################
		% Paste m file below for Step #3
		%###################################
		% Change: MSTOT = CurrentGuess		
		%###################################
		%###################################

output @clear
%output @nciout=24*30 T CBSNGKGLIADJ
%prepare @clear year T CLINGKG CFNGKG CBSNGKGLIADJ BBNGKG  CBNDLINGKG CBNGKG
prepare @clear T CBSNGKGLIADJ CBNGKG
%output @all 

% PARAMETERS FOR SIMULATION
% Simulate a 4 year old boy exposed to the Seveso incident
% A single pulse dose at time = 0
% Target Concentration: 26,000 ppt (ng/kg)

MAXT = 0.5
CINT = 1. %
EXP_TIME_ON  = 0.         % Delay before begin exposure (HOUR)
EXP_TIME_OFF = 43801.     % HOUR/YEAR !TIME EXPOSURE STOP (HOUR)
DAY_CYCLE    = 24.        % TIME 
BCK_TIME_ON  = 0.         %324120     % DELAY BEFORE BACKGROUND EXP (HOUR)
BCK_TIME_OFF = 613200     %324120     % TIME OF BACKGROUND EXP STOP (HOUR) 
TIMELIMIT    = 43800.     % 10 years
MSTOTBCKGR   = 0.         %3.35E-4       % ORAL BACKGROUND EXPOSURE DOSE (nG/KG/day)

% oral dose oral dose oral dose 
MSTOT        = CurrentGuess  % Serveso, ORAL DAILY EXPOSURE DOSE (NG/KG/day)
DOSEIV       = 0          % 40  %50 %5 %0.5 %0.3 %0.2 %0.1%0.05%0.3 %NG/KG
% oral dose oral dose oral dose 

MEANLIPID    = 730        % 711 %664 %778 %468 %671 %730 %662 %592%615%730%
PAS_INDUC    = 1          % NON INDUCTION (0) CONTROLE DE L'INDUCTION 

%human variable parameter 
MALE = 1.
FEMALE = 0.
Y0 = 0.                % 0 years old at the beginning of the simulation

start @nocallback
%CBSNGKGLIADJ
%CBNGKG
meanCBSNGKGLIADJ=mean(_cbsngkgliadj);
%meanCBSNGKGLIADJ
maxCBSNGKGLIADJ=max(_cbsngkgliadj);
%maxCBSNGKGLIADJ

		%###################################
		% End m file paste for Step #3
		%###################################
		%###################################
		%###################################
		%###################################
		%###################################
		%###################################

		% GET CURRENT VALUE
		if (ValueToCheck == "maxCBSNGKGLIADJ")
			CurrentValue = maxCBSNGKGLIADJ;
		else	% "meanCBSNGKGLIADJ"
			CurrentValue = meanCBSNGKGLIADJ;
		end

		% SEE IF CURRENT VALUE IS REACHED, IF IT IS THEN EXIT
		% Note that this is based if the values are identical in scientific notation
		% using 3 digits (X.XXE-YY)
		if (strcmp(sprintf("%.2E",CurrentValue), sprintf("%.2E",TargetValue)) == 1) 
			break
		end 

		%IF ITS NOT REACHED, THEN BEGIN GUESSING

		%SET HIGHER AND LOWER WITH NEW GUESSES
		if CurrentValue > TargetValue
			Higher_MSTOT = CurrentGuess;
			Higher_Value = CurrentValue;
		end

		if CurrentValue < TargetValue
			Lower_MSTOT = CurrentGuess;
			Lower_Value = CurrentValue;
		end

		%Now, create a new guess, by a linear interpolation
		if (Higher_MSTOT ~= 0) & (Lower_MSTOT ~= 99999999) 
			CurrentGuess = Lower_MSTOT + (TargetValue - Lower_Value)*((Higher_MSTOT-Lower_MSTOT)/(Higher_Value-Lower_Value));
		end
		
		% no guess lower yet
		if Lower_MSTOT == 99999999
			CurrentGuess = CurrentGuess * 0.1;
		end
		
		% no guess higher yet
		if Higher_MSTOT == 0
			CurrentGuess = CurrentGuess * 10.0;
		end
	end
	
	%Save outputs for Step #3
	Output_Step3 = sprintf("%s\n%s%s",Output_Step3,"Value to match = ",ValueToCheck);
	Output_Step3 = sprintf("%s\n%s%g",Output_Step3,"Match found on run number = ",Run);
	if MaxNumberOfAttempts == Run	
		Output_Step3 = sprintf("%s\n%s",Output_Step3,"Warning! Run Count Equal to MaxNumberOfAttempts!");
	end 
	Output_Step3 = sprintf("%s\n%s%g",Output_Step3,"Target Value = ",TargetValue);
	Output_Step3 = sprintf("%s\n%s%g",Output_Step3,"Current Value = ",CurrentValue);
	Output_Step3 = sprintf("%s\n%s%g\n",Output_Step3,"MSTOT of Current Value = ",CurrentGuess);
	
end 

% BEEP WHEN COMPLETE

if BeepWhenComplete == 1
	!
end 

% Print Outputs:
Output_Step1
Output_Step2
Output_Step3