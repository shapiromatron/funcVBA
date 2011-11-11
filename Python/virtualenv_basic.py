'''
Set up a bare-bones Python virtual environment with virtual env.

Note that virtualenv should be installed for the desired version
of Python before running this script for that version.

Install virtualenv for the desired version of Python this in Win cmd:
'C:\Python2x\Scripts>pip.exe install virtualenv'
    or:
'C:\Python2x\Scripts>easy_install.exe virtualenv'

Note that virtual environments built with virtualenv automatically
include pip. Using pip within the virtual environment only affects
the virtual environment's packages. See >virtualenv.exe for options
supported by virtualenv.

If you don't have a global install of pip (e.g. in C:\Python2x\Scripts), 
get it:  http://www.pip-installer.org/en/latest/installing.html
'''


import subprocess
import sys

def main(argv):
    '''Main entry point for this program.'''

    # Specify desired version of Python
    PY_EXE = "C:\\Python27\\python.exe"
    
    # Assume virtualenv.exe resides in 
    # the usual location relative to `PY_EXE`
    VIRTUAL_ENV_EXE = '\\'.join(PY_EXE.split('\\')[:-1]) + '\\Scripts\\virtualenv.exe'
    
    # Specify the name of our virtual environment
    ENV_NAME = 'virtualenv_testenv'
    
    # Specify desired location of the virtual 
    # environment that we will create. Change
    # 'C:\\' as needed.
    ENV_DEST = "C:\\" + ENV_NAME
    
    
    # Create virtual environment by calling
    # virtualenv.exe with given `ENV_DEST`
    retcode = subprocess.call([VIRTUAL_ENV_EXE, ENV_DEST])
    
    
    # A non-zero return code from the process
    # indicates an 'abnormal' execution.
    if retcode != 0:
        raise Warning("Error creating virtual environment at %s with %s!" % (ENV_DEST, VIRTUAL_ENV_EXE))
        return False # unsuccessful
    
    # You can test the virtual environment by
    # running 'activate.bat' in <ENV_DEST\Scripts\>
    return True # successful

if __name__ == '__main__':
    argv = sys.argv
    successful = main(argv)
    print '[virtualenv_basic.py] -- The install was successful: %s' % str(successful).upper()