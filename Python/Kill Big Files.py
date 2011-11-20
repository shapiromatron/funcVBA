#!/usr/bin/env python

""" 
Deletes all files greater than the specified threshold in the specified path.

Example function call from command line:
    python kill_big_files.py "M:\My Downloads\Installed Programs" 20
    (would delete all files > 20 mb in this path)
    
To pipe output to file:
    python kill_big_files.py "M:\My Downloads\Installed Programs" 20 >> C:\Temp\pipe.out
"""

import sys
import os

FAIL_MSG = """kill_big_files error- two inputs required: 
    1) Full path (i.e. "C:\Temp")
    2) minimum kill size (MB)"""

def kill_big_files(dir, min_kill_size_mb):
    """ If a file is greater than the specific limit, delete the file"""
    BYTES_IN_MB = 1048576
    min_kill_size = int(min_kill_size_mb) * BYTES_IN_MB # converts from MB to bytes
    if os.path.exists(dir) == True:
        files = os.listdir(dir)
        for file in [ os.path.join(dir, f) for f in files ]:
            size = os.path.getsize(file)
            if size >= min_kill_size:
                os.remove(file)
                print "Deleted: %s (%d MB)" % (file, size / BYTES_IN_MB)
    else:
        print "Path doesn't exist: %s" % dir
    
def main(argv):
    """ Attempt to kill big files if the number of arguments are correct,
        if any error is reached, spit fail message """
    if len(argv) == 3:
        dir =  argv[1]
        min_kill_size = argv[2]
        try:
            kill_big_files(dir, min_kill_size)
        except:
            print FAIL_MSG
    else:
        print FAIL_MSG

if __name__ == '__main__':
    main(sys.argv)