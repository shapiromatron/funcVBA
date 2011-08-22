""" Steps to create an EXE: (http://www.py2exe.org/index.cgi/Tutorial)

1. Install py2exe on your computer (http://www.py2exe.org/)
2. Create the script that you want to compile as save it. 
3. Create a new setup script, as shown below
4. Run the script in MSDOS: "python setup.py py2exe"
5. Then, delete everything except stuff in the "dist" folder.  That's your executable

"""

# Filename: setup.py

from distutils.core import setup
import py2exe
setup(console=[r'C:\Temp\ScriptToCompile.py'])