#!/usr/bin/env python
import zipfile

def ZipIt(WriteFile, ZipFN):
    """ Create a new Zip file and add a file to it """
    newZip = zipfile.ZipFile
    with newZip(ZipFN, 'a', compression=zipfile.ZIP_DEFLATED) as myzip:
        myzip.write(WriteFile)
        
zipfn = r'C:\Documents and Settings\16955\Desktop\PnCB Unit Checks\Load to Conc\ziptest.zip'
fn = r'C:\Tools\PnCB\OUTPUTS\MonteCarlo\hr_air_conc_no_bg_real9.txt'
ZipIt(fn, zipfn)