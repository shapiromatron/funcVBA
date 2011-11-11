from win32com.client import Dispatch

def RunExcelMacro(FN, MacroName):
    xl = Dispatch('Excel.Application')
    xl.Visible = 1
    xl.DisplayAlerts = 0
    xl.Workbooks.Open(FN)
    xl.ActiveWorkbook.Sheets("IEUBK_in").Activate
    xl.ActiveWorkbook.Sheets("IEUBK_in").Range("F10:F93").Value = 100   #soil
    xl.ActiveWorkbook.Sheets("IEUBK_in").Range("G10:G93").Value = 50   #dust
    xl.ActiveWorkbook.Sheets("IEUBK_in").Range("I10:I93").Value = 5   #dust
    #xl.Run('RunShell')
    #xl.ActiveWorkbook.SaveAs(Filename="C:\Tools\Book2.xlsm")
    xl.Quit()
    

FN = r'C:\Tools\IEUBK Model 09-13-11.xlsm'
RunExcelMacro(FN,'blah')
AirData = r'C:\Tools\PnCB\OUTPUTS\MonteCarlo\THOROUGH_CHECK\ExposureProfiles_and_MonthlyData\08-30-11 Input File_real1_MonthlyAirConcs.txt'
DustData = r'C:\Tools\PnCB\OUTPUTS\MonteCarlo\THOROUGH_CHECK\ExposureProfiles_and_MonthlyData\08-30-11 Input File_real1_MonthlyAirConcs.txt'
    