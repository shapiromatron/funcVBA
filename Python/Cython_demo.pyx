'''
Basic demo for learning Cython. This demo demonstrates using
Cython to write C code, not C++ code. Cython can only use one
of C or C++ code simultaneously, although this page says you
can write 'shim' modules to mix them (but mixing is not currently
supported out-of-the-box): 

    Notice this file's '.pyx' file extension. Cython code goes here.
This will be compiled to a 'Cython extension module.' This module
can then be called from a normal Python script.
    To compile this script, call the 'Cython_demo_setup.py' script:
	./>python Cython_demo_setup.py build_ext --inplace
Read Cython_demo_setup.py to see the basic elements of a Cython setup 
script. This template changes very little, especially when compiling
only one *.pyx file into an extension module. The script builds a C
source file from the *.pyx file(s), and this is then compiled to a
Python extension module (*.pyd). You can import the *.pyd file into
a pure Python environment as you would any module, and the C code
does not (in my experience) need to be present to use it:

	import Extension_module_name

ABOUT:

    Cython is a language that makes writing C extensions for the Python language
as easy as Python itself. Cython is based on the well-known Pyrex, but supports 
more cutting edge functionality and optimizations.
    The Cython language is very close to the Python language, but Cython 
additionally supports calling C functions and declaring C types on variables and 
class attributes. This allows the compiler to generate very efficient C code 
from Cython code. This makes Cython the ideal language for wrapping external C 
libraries, and for fast C modules that speed up the execution of Python code.
    Good Cython code achieves optimal speed by minimizing interaction with pure
Python types and maximizing interaction with C/C++ types. This means strong
typing via type declarations is essential for fast code. For a line-by-line
'annotation' of your Cython code's speed, use '\>Cython.py --annotate File.pyx.'
This generates an HTML file with each line highlighted, with brighter highlights
representing a larger number of Python calls (and thus slower execution).
    Optimizing Cython for complex procedures can be non-trivial. Consult the 
documentation and experiment frequently to find the best performance.

A good way to use this demo might be to 'annotate' it and find room for
improvement!
    
WEB SITES:
Cython home:            http://cython.org/
Documentation:          http://docs.cython.org/
Working with NumPy:     http://docs.cython.org/src/tutorial/numpy.html

INSTALLATION:
Binaries:               http://www.lfd.uci.edu/~gohlke/pythonlibs/#cython
MinGW on Windows:       http://docs.cython.org/src/tutorial/appendix.html
'''


#
#   IMPORTING
#
# import pure Python modules normally
import sys
import os
import numpy as np # import normally for Python calls

# import C-accessible modules using cimport
# this includes (to my knowledge) NumPy and
# appropriately wrapped C/C++ libraries,
# maybe SciPy as well
cimport numpy as np # import specially for C calls

# libc and libcpp come with Cython
# for me, here:
# C:\Python27\Scripts\Personal\Cython-0.15\Cython-0.15\Cython\Includes
from libc.math cimport remainder


#
#   CUSTOM C-TYPES
#
# ctypedef <type from known source> <new type alias>
# notice the '_t' for the numpy float type: this is
# the runtime-available type, and accesses C-code.
# The 'Working with NumPy' guide (link above) is
# critical for using NumPy with Cython. It's great!

ctypedef np.float64_t DTYPE_float_t


#
#   CONSTANTS
#
cdef int CONSTANT1 = 5      # cdef <type name> <variable name>
cdef float CONSTANT2 = 5.0
cdef float CONSTANT3        # not required to initialize a value
cdef list CONSTANT4 = []    # some benefit from cdef-ing Python built-ins


#
#   DEFINING FUNCTIONS NOT CALLED FROM THE OUTSIDE
#
#
# define functions called by this Cython
# extension module with 'cdef'. These must
# be defined outside of 'def' functions!

cdef float csum(np.ndarray[DTYPE_float_t, ndim=1] array):
    '''Fast sum of an array.'''
    
    # Declare some variables with known types.
    # Because all function variables are defined
    # as C-types, the for-loop below is about as
    # fast as a for-loop written in pure C!
    cdef DTYPE_float_t running_total = 0.
    cdef DTYPE_float_t value
    
    for value in array:
        running_total += value  # in-place addition
        
    return running_total
    

#
#   DEFINING FUNCTIONS CALLED BY PURE PYTHON
#
# you must use the normal 'def' or 'cpdef' 
# to accept and return pure Python types to
# pure Python code!

def main(DTYPE_float_t arg1, int arg2, float arg3):
    '''Compute stuff real fast!'''
    
    # Define some local variables
    # Use special 'buffer' syntax available for NumPy ndarrays
    # np.ndarray[<array data type>, <# of dim>]
    cdef np.ndarray[DTYPE_float_t, ndim=1] array = np.zeros(arg2, dtype=np.float64)
    cdef int counter
    cdef DTYPE_float_t ret = 0.
    
    for counter in range(0, 1000):
        if counter < arg2:
            ret += arg3
            
    return array + ret