#call from command line:
# python Cython_demo_setup.py build_ext --inplace

#!/usr/bin/env python
from distutils.core import setup
from distutils.extension import Extension
from Cython.Distutils import build_ext

module_name  = "Cython_demo"
source_files = ["Cython_demo.pyx"]

# ext_modules = [Extension(module_name, [list_of_files])]
ext_modules = [Extension(module_name, source_files)]

setup(
  name = module_name,
  cmdclass = {'build_ext': build_ext},
  ext_modules = ext_modules
)