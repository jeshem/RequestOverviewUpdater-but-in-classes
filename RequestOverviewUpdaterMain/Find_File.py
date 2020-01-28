import shutil
import os
import datetime
import pywildcard

class Find_File:
    location = ""
    overviewfile =""
    today = ""

    def __init__(self, loc, ov_file):
        print ("Created Find_File object")
        self.location = loc
        self.overviewfile = ov_file
        self.today = str(datetime.datetime.now().date())


    #make a backup of the Total Request Overview excel sheet
    def make_copy(self):
        shutil.copy(self.location + "\\" + self.overviewfile + ".xlsx", self.location + 
                    "\\" + self.overviewfile + " " + self.today + " backup.xlsx")

    #find new files and insert them into a list
    def find_new_files(self, loc, list):
        print("fine_new_files called")
        dirs = os.listdir(loc)
        for file in dirs:
            #get the date the file was modified/created
            file_mod_time = os.stat(loc + "\\" + file).st_mtime
            file_create_time = os.stat(loc + "\\" + file).st_ctime

            #get the date the overview file was modified
            overview_mod_date = os.stat(self.location + "\\" + self.overviewfile + ".xlsx").st_mtime
            
            #only get the files that were modified/added today
            if file_mod_time > overview_mod_date or file_create_time > overview_mod_date:
                #if the file is the overview file, the backup of the overview file, or a master version of a file don't add it to the list
                if not file.startswith("~$"):
                    #if the file is a directory, go into it and check for new files
                    if os.path.isdir(loc + "\\" + file):
                        list = self.find_new_files(loc + "\\" + file, list)
                    #if the file is an excelsheet, add it to the list
                    if pywildcard.fnmatch(file, '*.xlsx'):
                        list.append(loc + "\\" + file)
        if list:
            self.make_copy()
        
        return list
    pass




