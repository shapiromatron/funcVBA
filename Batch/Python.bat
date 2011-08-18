::#=====================#
::#   PYTHON COMMANDS   #
::#=====================#

:: PYTHON: PROFILE A SCRIPT
cd C:\ScriptLocation
python -m cProfile C:/ScriptLocation/ScriptName.py >> profile_summary.txt

:: Install package using PIP
cd C:\Python27\Scripts
pip install scipy
pip install cython