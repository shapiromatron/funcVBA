# move files from here to there
import sys
import os
import shutil

def CopyFiles(argv, tries=0):
    if tries >= 6:
        raise Exception('You entered %r incorrect directories. I QUIT!' % tries)
    # Ask user to specify source directory
    print "\nEnter source directory: (don't mind the r)\n"
    src = os.path.abspath(raw_input('r'))
    if os.path.isdir(src) is False: # if the user's supplied directory doesn't exist, restart function
        tries += 1
        print "\nThe specified source directory is not valid; try again."
        MoveFiles(argv, tries=tries) # recursive call
    # Ask user to specify destination directory
    print "\nEnter destination directory: (don't mind the r)\n\n"
    dest = os.path.abspath(raw_input('r'))
    if os.path.isdir(dest) is False: # if the user's supplied directory doesn't exist, restart function
        tries += 1
        print '\nThe specified destination directory is not valid; try again.'
        MoveFiles(argv, tries=tries) # recursive call
        
    def file_to_path(file, path):
        # returns a file's full path based on arguments
        return path+'\\'+file
    
    files_to_move = os.listdir(src) # get the names of the objects in the source directory
    src_paths = [file_to_path(file,src) for file in files_to_move if os.path.isfile(file_to_path(file,src))] # get the full paths for the source directory's files, exluding non-file items (e.g. folders)
    dest_paths= [file_to_path(file,dest) for file in files_to_move] # invent new paths for the destination files to be created based on the source directory's flat file names in src_paths
    src_dest_tups = zip(src_paths, dest_paths) # zip into list of tuples like [(sourcepath_0, destpath_0),...,(sourcepath_n-1,destpath_n-1)] to feed to shutil.copyfile
    
    for dest_path in dest_paths:
        if os.path.isfile(dest_path): # if the target destination filename exists, delete that file (currently doesn't seem to work if the destination file exists)
            os.remove(dest_path)
    
    shutil_copyfile = shutil.copyfile # for speed
    for tup in src_dest_tups: # e.g. tup = (sourcepath_0, destpath_0)
        source_path = tup[0] # e.g. C:\\here\\somefile.txt
        destin_path = tup[1] # e.g. C:\\there\\somefile.txt
        shutil_copyfile(source_path, destin_path)

if __name__ == '__main__':
    CopyFiles(sys.argv)