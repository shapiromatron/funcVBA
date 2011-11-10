'''
Run this script to create a Win32 executable file
from a Python script, using PyInstaller.

PyInstaller site:           http://www.pyinstaller.org/
PyInstaller tar.bz2 file:   http://files.zibricky.org/pyinst/pyinstaller-1.5.1.tar.bz2
PyInstaller zip file:       http://files.zibricky.org/pyinst/pyinstaller-1.5.1.zip
Online docs:                http://www.pyinstaller.org/export/latest/tags/1.5.1/doc/Manual.html?format=raw
'''

import os
import sys
import subprocess


def pyinstaller_dir_func():
    try:
        print 'Specify the full path of the PyInstaller 1.5.1 directory:\n\n'
        pyinstaller_dir = os.path.abspath(raw_input('r'))
        listdir = os.listdir(pyinstaller_dir)
        success = os.path.isdir(pyinstaller_dir) is True and 'Makespec.py' in listdir and 'Build.py' in listdir
        return pyinstaller_dir, success
    except:
        return '', False

def pySource_fullpath_func():
    try:
        print '\n\nSpecify the full path of the \nPython file you want to compile:\n\n'
        pySource_fullpath = os.path.abspath(raw_input('r'))
        if os.path.isfile(pySource_fullpath) is True and len(pySource_fullpath) >= 8 and pySource_fullpath[-3:]=='.py':
            success = True
        else:
            success = False
        return pySource_fullpath, success
    except:
        return '', False

def EXE_outDir_func():
    try:
        print '\n\nSpecify the full path of the desired output directory:\n\n'
        EXE_outDir = os.path.abspath(raw_input('r'))
        return EXE_outDir, os.path.isdir(EXE_outDir)
    except:
        return '', False

def cmd_console_func():
    try:
        print ("\n\nDo you want the executable to open the Windows\n"
                "console (cmd prompt) when run? (y/n)\n\n")
        cmd_console = ''
        while cmd_console not in ('y', 'n'):
            cmd_console = raw_input('r')
        ret = '--noconsole' if cmd_console=='n' else ''
            
        return ret
    except:
        return False

def ico_fullpath_func():
    try:
        print ("\n\nIf you want the EXE file to have a custom icon (*.ico file),\n"
                "paste full file path here (otherwise leave blank):\n\n")
        ico_fullpath = os.path.abspath(raw_input('r'))
        
        if ico_fullpath == '':
            ico_fullpath = use_custom_ico = ''
        elif os.path.isfile(ico_fullpath) is False or len(ico_fullpath) < 8 or ico_fullpath[-3:] != 'ico':
            ico_fullpath = use_custom_ico = ''
        elif os.path.isfile(ico_fullpath) is True and len(ico_fullpath) >= 8 and ico_fullpath[-3:]=='ico':
            use_custom_ico = '--icon='
        else:
            ico_fullpath = use_custom_ico = ''
            
        return ico_fullpath, use_custom_ico
    except:
        return '', ''


def main(argv):
    '''
    Ask user to specify setup requirements, and compile
    Python script to an executable using PyInstaller.
    '''
    print '\n\nUSER INPUT REQUIRED\n\n'
    
    pyinstaller_dir_success = None
    pySource_fullpath_success = None
    EXE_outDir_success = None
    
    while pyinstaller_dir_success is not True:
        pyinstaller_dir, pyinstaller_dir_success = pyinstaller_dir_func()
        if pyinstaller_dir_success is not True:
            print ("\n\nERROR: the directory you specified for\n"
                    "PyInstaller is not a valid directory!\n\n")
        
    while pySource_fullpath_success is not True:
        pySource_fullpath, pySource_fullpath_success = pySource_fullpath_func()
        if pySource_fullpath_success is not True:
            print ("\n\nERROR: the full path you specified for the\n"
                    "Python script to compile is not valid!\n\n")
       
    while EXE_outDir_success is not True:
        EXE_outDir, EXE_outDir_success = EXE_outDir_func()
        if EXE_outDir_success is not True:
            print ("\n\nERROR: the directory you specified for\n"
                    "the EXE's output is not a valid directory!\n\n")
                    
    use_cmd_console = cmd_console_func()
    
    ico_fullpath, use_custom_ico = ico_fullpath_func()
    
    specfile_fullpath = EXE_outDir + '\\' + pySource_fullpath.split('\\')[-1][:-2] + 'spec'
    
    specfile_argv = ((r'python %s\Makespec.py --onefile --out=%s %s %s%s %s' 
                        % (pyinstaller_dir, EXE_outDir, use_cmd_console, use_custom_ico, ico_fullpath, pySource_fullpath)).split())
    build_argv = (r'python %s\\Build.py %s' % (pyinstaller_dir, specfile_fullpath)).split()

        
    specfile_process = subprocess.call(specfile_argv)
    build_process = subprocess.call(build_argv)

    
if __name__ == '__main__':
    argv = sys.argv
    main(argv)