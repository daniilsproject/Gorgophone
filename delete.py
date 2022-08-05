import os
import config

path = config.path_to_save

def remove_one_page_files(path = path):
    counter = 0

    for filename in os.listdir(path):

        if ((filename[0] != '_') and (filename[1] != '_')):
            
            print ("File: "+filename)
            file_path = os.path.join(path, filename)
            os.remove(file_path)
            counter += 1

    print("All temporary files (%s) removes from %s " %(counter,path))


remove_one_page_files(path)