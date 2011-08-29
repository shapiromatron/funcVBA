' LOTS OF REALLY GOOD INFORMATION HERE:
' http://www.cpearson.com/excel/ErrorHandling.htm
Private Sub ErrorHandling()
    On Error GoTo 0                 'Default VBA case, Displays error message, ends execution
    On Error Resume Next            'Skips line with error and goes to next line
    On Error GoTo ErrHandler:       'Use error handler, once resolved go back to code
   
ErrHandler:                         'Must be in the sub/function/etc
    If Err.Number = 9 Then
        'do something to fix it
        Resume                      'Go back to same line that caused the problem
    End If
    
End Sub

'File Not Found Error
On Error GoTo FileNotFound
    Open (ExampleFile) For Input As #1
    FileCopy ExampleFile, ExamplePaste
On Error GoTo 0
FileNotFound:
    MsgBox "Error reading from file:" & vbNewLine & ExampleFile & _
                vbNewLine & "Make sure file exists and is readable and try again.", _
                vbCritical, "File error"
                
' return error in function:
'Return = CVErr(xlErrNA) xlErrDiv0, xlErrNA, xlErrName, xlErrNull, xlErrNum , xlErrRef, xlErrValue

'    POTENTIAL FIX?
'    MkDir (DirName)
'    Resume

'ERROR MESSAGES NUMBERS
' 3  Return without GoSub
' 5  Invalid procedure call
' 6  Overflow
' 7  Out of memory
' 9  Subscript out of range
' 10 This array is fixed or temporarily locked
' 11 Division by zero
' 13 Type mismatch
' 14 Out of string space
' 16 Expression too complex
' 17 Can't perform requested operation
' 18 User interrupt occurred
' 20 Resume without error
' 28 Out of stack space
' 35 Sub, Function, or Property not defined
' 47 Too many DLL application clients
' 48 Error in loading DLL
' 49 Bad DLL calling convention
' 51 Internal Error
' 52 Bad file name or number
' 53 file Not Found
' 54 Bad file mode
' 55 File already open
' 57 Device I/O error
' 58 File already exists
' 59 Bad record length
' 61 Disk full
' 62 Input past end of file
' 63 Bad record number
' 67 Too many files
' 68 Device unavailable
' 70 Permission denied
' 71 Disk Not Ready
' 74 Can't rename with different drive
' 75 Path/File access error
' 76 Path Not Found
' 91 Object variable or With block variable not set
' 92 For loop not initialized
' 93 Invalid pattern string
' 94 Invalid use of Null
' 97 Can't call Friend procedure on an object that is not an instance of the defining class
' 98 A property or method call cannot include a reference to a private object, either as an argument or
'    as a return value
' 298 System DLL could not be loaded
' 320 Can't use character device names in specified file names
' 321 Invalid file format
' 322 Can’t create necessary temporary file
' 325 Invalid format in resource file
' 327 Data value named not found
' 328 Illegal parameter; can't write arrays
' 335 Could not access system registry
' 336 ActiveX component not correctly registered
' 337 ActiveX Not component
' 338 ActiveX component did not run correctly
' 360 Object already loaded
' 361 Can't load or unload this object
' 363 ActiveX control specified not found
' 364 Object was unloaded
' 365 Unable to unload within this context
' 368 The specified file is out of date. This program requires a later version
' 371 The specified object can't be used as an owner form for Show
' 380 Invalid property value
' 381 Invalid property-array index
' 382 Property Set can't be executed at run time
' 383 Property Set can't be used with a read-only property
' 385 Need property-array index
' 387 Property Set not permitted
' 393 Property Get can't be executed at run time
' 394 Property Get can't be executed on write-only property
' 400 Form already displayed; can't show modally
' 402 Code must close topmost modal form first
' 419 Permission to use object denied
' 422 Property not found
' 423 Property or method not found
' 424 Object Required
' 425 Invalid object use
' 429 ActiveX component can't create object or return reference to this object
' 430 class doesn't support Automation
' 432 File name or class name not found during Automation operation
' 438 Object doesn't support this property or method
' 440 Automation Error
' 442 Connection to type library or object library for remote process has been lost
' 443 Automation object doesn't have a default value
' 445 Object doesn't support this action
' 446 Object doesn't support named arguments
' 447 Object doesn't support current locale setting
' 448 Named Not Argument
' 449 Argument not optional or invalid property assignment
' 450 Wrong number of arguments or invalid property assignment
' 451 Object not a collection
' 452 Invalid ordinal
' 453 Specified DLL function not found
' 454 Code Not resource
' 455 Code resource lock error
' 457 This key is already associated with an element of this collection
' 458 Variable uses a type not supported in Visual Basic
' 459 This component doesn't support the set pf events
' 460 Invalid Clipboard format
' 461 Specified format doesn't match format of data
' 480 Can't create AutoRedraw image
' 481 Invalid Picture
' 482 printer Error
' 483 Printer driver does not support specified property
' 484 Problem getting printer information from the system. Make sure the printer is set up correctly
' 485 Invalid picture type
' 486 Can't print form image to this type of printer
' 520 Can't empty Clipboard
' 521 Can't open Clipboard
' 735 Can't save file to TEMP directory
' 744 Search Not Text
' 746 Replacements too long
