#===============================================================================
#   STATEMENTS 
#===============================================================================
'''handy resources:   
        http://docs.python.org/tutorial/controlflow.html
    more advanced: http://docs.python.org/reference/simple_stmts.html
    more advanced: http://docs.python.org/reference/compound_stmts.html'''

=           # value assignment
and         # check multiple conditions being true
as          # name something 'as' something
assert      # check a condition - raise AssertionError exception if false
break       # break out of a loop (for loop, while loop)
class       # define a new class of objects
continue    # go to the next iteration of a loop
def         # define a function
del         # delete an object
elif        # else if; in an if block, conditionally trigger this when if is 
            # false
else        # in an if block, always trigger this when if is false
except      # in a try-except block, trigger this when some exception is hit
exec        # dynamically execute Python code
finally     # in a try-except-finally block, always do this at the end
for         # iterate over elements in a sequence
from        # from <module> import <object> [as <name>] --bracketed code is 
            # optional
global      # instantiate a global variable: global VarName
if          # check for a condition's true/false value
import      # import an external module
in          # access an iterable; e.g., 'if k in (1,2)'; 'for k in (1,2)'
is          # check for equality; e.g., 'if k is True'
lambda      # use a function ad-hoc without defining it
not         # opposite of 'is', check for absence of something, e.g. 
            # 'if not None' --> True
or          # check that at least one condition is true
pass        # do nothing, such as in a loop or when handling exceptions
print       # display value(s) to stdio: print 1, 2, 3 --> 1\n2\n3 
            # (\n is a new line char)
raise       # raise an exception
return      # return a value from a function, do not execute subsequent code 
            # in function
try         # in a try-except block, 'try' some code and handle some 
            # 'except'ions
with        # with(filename, mode) --> execute code block with 'filename' 
            # open then close 'filename'
while       # loop while a condition is true
yield       # return a value from a Python generator object

# CONDITIONAL SIGNS
EQUALITY:                   ==  # object == object, not necessarily numeric
INEQUALITY:                 !=  # object != object, not necessarily numeric

LESS THAN:                  <   # almost always numeric
LESS THAN OR EQUAL TO:      <=  # almost always numeric

GREATER THAN:               >   # almost always numeric
GREATER THAN OR EQUAL TO:   >=  # almost always numeric


#===============================================================================
#   STATEMENTS IN ACTION 
#===============================================================================

# =
# give k the value 5; 
# creates an integer (int) object with 
# value 5 in memory; k points to that object
k = 5 

#-------------------------------------------------------------------------------
# and
# check for multiple conditions
if k is 5 and k is < 10:
    print 'k is 5 and k is less than 10'
    
#-------------------------------------------------------------------------------
# assert
assert 1 + 1 == 2 # True, nothing happens
assert 1 + 1 == 11 # False, raises AssertionError

#-------------------------------------------------------------------------------
# break
for k in (1,2,3,4):
    if k > 2:
        break # for loop quits on 3rd iteration, when k equals 3
while k < 100:
    if k > 100: 
        break # while loop quits if k > 100
    else:
        k = k + 1 # make sure this isn't an infinite while loop!

#-------------------------------------------------------------------------------
# class
class NewObject():
    # self variable represents this object itself
    def __init__(self, arg1):
        # special __init__ function for all Python
        # classes sets properties of the object as
        # soon as you create it
        self.property1 = arg1
    def method(self):
        # defines a function belonging to self
        # that can operate on self and its properties
        # called using self.method() or objectname.method()
        print 'Do something with this NewObject.'
        print 'Methods are just functions!'

# to make an object from a class, simply call
# the class and set a variable to its value:
object = NewObject(5)
# `object` inherits the specific properties of
# the NewObject class of objects, and is called
# an 'instance' of the NewObject class.
print object.property1 # this prints 5
object.method() # calls `method` function

#-------------------------------------------------------------------------------
# continue
for k in [1,2,3,4,5]:
    if k == 3:
        continue # don't print 3, continue to k=4
    else:
        print k

#-------------------------------------------------------------------------------
# def
def Function(formalArgument, optionalArgument=5)
    return formalArgument * optionalArgument
    
k = Function(5) # sets k to 25
k = Function(5, 1) # sets k to 5 -- positions must be obeyed
k = Function(5, optionalArgument=0) # sets k to 0

#-------------------------------------------------------------------------------
# del
del k # deletes k from memory, will cause NameError if called without 
      # re-initializing!

#-------------------------------------------------------------------------------
# elif & else
# once one of the conditions is True,
# the others will not execute!
if 5 < 0:
    print '5 is less than 0'
elif 10 < 5:
    print '10 is less than 5'
elif 15 < 20:
    print '15 is less than 20'
else:
    print 'None of the above were true'

#-------------------------------------------------------------------------------
# except -- try/except/else/finally blocks
# a try statement cannot be alone!
try:
    print 5 + '5' # cant use '+' operator on int and str, raises TypeError
except:
    print 'Error!'  # catch any exception
    
try:
    print 5 + 5
except:
    print 'Error!'
else:
    print 'Success!' # 'else' triggered when no exception is found
    
try:
    print 5 + '5'
except TypeError, e: # only catch TypeError, save the error message to 'e'
    print e
    
try:
    print 5 + '5'
except TypeError, ValueError: # catch either of these errors, but cant save 
                              # error msg to 'e'
    print 'TypeError or ValueError!'
    
try:
    print 5 + '5'
except TypeError, e: # use 2 except lines to be able to save error msg to 'e'
    print e
except ValueError, e:
    print e
    
try:
    5 + '5'
except:
    print 'Error'
else:
    print 'Success!'
finally:    # regardless of errors, always execute this code
    print 'This ALWAYS happens!'

#-------------------------------------------------------------------------------
# finally statement -- see 'except' above

#-------------------------------------------------------------------------------
# for
iterable = [1,2,3]
for element in iterable:
    print element

#-------------------------------------------------------------------------------
# from
# from <module> import <object>
# from <module> import <object> as <new name>
from os import path
from os import path as p

#-------------------------------------------------------------------------------
# global
global newVar # cannot perform value assignment at the same time!
newVar = 5

#-------------------------------------------------------------------------------
# if
condition1 = (1 == 1) # True
condition2 = (1 == 0) # False
#   These print (they are True statements):
if condition1 is True:
    print '1 equals 1'
if condition1:
    print '1 equals 1'
if condition1 is not False:
    print '1 equals 1'
if condition2 is False:
    print '1 does not equal 0'
if not condition2:
    print '1 does not equal 0'
#   These do not print (they are False statements):
if (condition1 is False):
    print '1 does not equal 1'
if (not condition1):
    print '1 does not equal 1'
if (condition2 is True):
    print '1 equals 0'
if (condition2 is not False:
    print '1 equals 0'

#-------------------------------------------------------------------------------
# import
import os
import os as operatingSystem_module

#-------------------------------------------------------------------------------
# in
print 5 in (1,2,3) # prints False
for k in (1,2,3):  # iterates
    print (k in (1,2,3)) # prints True

#-------------------------------------------------------------------------------
# is -- basically a readable equals statement
5 is 5 # True
5 is not 5 # False
5 is 6 # False
5 is not 6 # True

#-------------------------------------------------------------------------------
# lambda
# take x as argument, return x + 5
Function = lambda x: x+5
print Function(5) # prints 5+5 = 10
# can be useful in list operations
map(lambda x: x*x, [1,2,3]) # returns new list [1,4,9]
map(lambda x,y: x*y, [1,2,3,4], [1,2,3,4]) # returns [1,4,9,16]

#-------------------------------------------------------------------------------
# not -- see 'is'

#-------------------------------------------------------------------------------
# or
if True or True:
    print 'At least 1 of the 2 was True' # prints
if True or False:
    print 'At least 1 of the 2 was True' # prints
if False or False:
    print 'At least 1 of the 2 was True' # does NOT print

#-------------------------------------------------------------------------------
# pass --- aka, do nothing on the line where 'pass' is located 
           # (NOT the same as break!)
try:
    5 + '5'
except:
    pass
    print 'Error!' # this still prints

if 5 == 5:
    pass
    print '5 is 5'

#-------------------------------------------------------------------------------
# print
print 5 # prints 5
print 1,2,3 # prints 1,2,3 on separate lines
# NB- there is also a print function in some version
# of Python.

#-------------------------------------------------------------------------------
# raise -- raise an error or warning
raise Warning('Warning!')
raise RuntimeWarning('Dubious runtime!')
raise RuntimeError('Unknown runtime error!')
raise TypeError # one at a time
raise ValueError 

#-------------------------------------------------------------------------------
# return -- more advanced than in 'def'
def ReturnExample(arg):
    if arg < 10:
        return arg * 10 # if triggered, 2nd return does not fire b/c function 
                        # returns and stops execution
    return arg # only occurs if arg is not less than 10

#-------------------------------------------------------------------------------
# try -- see 'except' position, after 'else' and before 'exec'

#-------------------------------------------------------------------------------
# with
# old version of Python required you to
# use the open() function to open a file
# and manually call its close() method.
# this automatically closes the file, saving
# a lot of headache.
with open('somefile.txt', 'r') as infile: # you must use the 'as' statement!
    lines = infile.readlines()
# upon restoring indentation, 'infile' object's 
# close() method is automatically called

#-------------------------------------------------------------------------------
# while
k = True:
while k is True: # this is an infinite loop- always code an exit condition
    print k
while k is True:
    print k
    k = False # this loop will only execute once if k is True when loop begins

#-------------------------------------------------------------------------------
# yield -- for making generators.
# Generators are very useful for accomplishing
# processor/memory intensive tasks with loops.
# normally when you loop over an iterable object,
# that iterable is stored in its entirety in memory.
# For huge data objects, this presents a problem.
# Generators work by only creating 1 element from
# the iterable at a time and destroying it as soon
# as possible.
#   first, create a function that returns a generator
def Generator(length=5, start=0):
    while start < length:
        yield i
        i += 1  # i = i + 1
#   now make an instance of the generator
gen = Generator()
for x in gen:
    print x   # prints 0, 1, 2, 3, 4