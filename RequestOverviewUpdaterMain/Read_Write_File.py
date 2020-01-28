import xlwings as xw


class Read_Write_File(object):

    def __init__(self, loc, ov_file):
        self.location = loc
        self.overviewfile = ov_file

        self.file_list=[]
        self.project_keys = []
        self.service_keys = []
        self.vm_keys = []

        self.project_data = {}
        self.service_data = {}
        self.vmcore_data = {}

    #append extracted information to the request overview fil
    @staticmethod
    def get_keys_from_init(self, location, proj_keys, serv_keys, vm_keys):
        section_name = ""
        with open(location + "/nsoci.ini") as fp:
            line = fp.readline()
            while line:
                s = line.strip()
           
                # skip blank line
                if not s == "":          
                    if section_name == "Projects":

                        if not '[' in s:
                            # extract keys under [Projects] and then continue
                            proj_keys.append(s)                    

                    elif section_name == "Services":
                        if not '[' in s:
                            # extract keys under [Services] and then continue
                            serv_keys.append(s)
                
                    elif section_name == "VM Cores":
                        if not '[' in s:
                            # extract keys under [VM Cores] and then continue
                            vm_keys.append(s)

                    if '[' in s:
                        section_name = self.check_section(s)
           
                line = fp.readline()
        return proj_keys, serv_keys, vm_keys

    @staticmethod
    def check_section(name):
        print ("\ncheck for section: " + name) 
    
        if name == "[Projects]":
            return "Projects"
    
        elif name == "[Services]":
            return "Services"
    
        elif name == "[VM Cores]":
            return "VM Cores"

    
    #using the keys found in the config file, pull data from the request forms
    @staticmethod
    def read_from_excel(file, proj_keys, serv_keys, vm_keys):
        print ("reading from " + str(file))
        wb = xw.Book(file)
        sht = wb.sheets[0]
        proj_data = {}
        serv_data = {}
        vm_data = {}
        dvm_data = {}
        #get data from the request forms
        for search in proj_keys:
            project_keys = sht.api.UsedRange.Find(search + ":")
            project_values = project_keys.offset(1, 7)
            data = project_values.value
            proj_data[search] = data
    
        for search in serv_keys:
            service_keys = sht.api.UsedRange.Find(search)
            service_values = service_keys.offset(1, 8)

            if service_values.value == "Not to be requested":
                data = None
            else:
                data = service_values.value
            serv_data[search] = data

            if search.startswith("VM") or search.startswith("BM"):
                num_of_cores = service_keys.offset(1, 2)
                print(num_of_cores.address + " " + str(int(num_of_cores.value)))
                for core_num in vm_keys:
                    if core_num == str(int(num_of_cores.value)):
                        print("Match found")
                        if data == None:
                            data = 0
                        if core_num in vm_data:
                            vm_data[core_num] = vm_data.get(core_num) + data
                        else:
                            vm_data[core_num] = data
    
        wb.close()
        return proj_data, serv_data, vm_data

    #using the data pulled from the request form, write to the overview file
    @staticmethod
    def write_to_excel(loc, overviewfile, proj_keys, serv_keys, proj_data, serv_data):
        wb = xw.Book(loc + "\\" + overviewfile + ".xlsx")
        sht = wb.sheets[0]
        xlShiftToDown = xw.constants.InsertShiftDirection.xlShiftDown
        
        #change these if the number of rows in the overview file start b
        max_rows = 500
        current_row = 6
        last_row = 0

        #insert a new row at the bottom of the list
        #currently, the max_rows is 500. This can be changed above if the list grows to exceed 500
        while current_row in range(max_rows):
            if sht.range((current_row, 1)).value == None:
                sht.range((current_row, 1)).api.EntireRow.Insert(Shift=xlShiftToDown)
                print("Row Inserted")
                last_row = current_row - 5
                current_row += max_rows
            current_row += 1
    
        #find the columns that the keys are in and insert the corresponding data into those columns
        for i in range(len(proj_keys)):

            #some of the keys taken from the request file do not match the column titles so I have to hardcode it
            #to look for the correct titles upon coming across those keys
        
            #I was unable to have the overview excel sheet automatically copy the formula to calculate the percentage
            #of 495k and 184k so I have to create it here. the value of four_nine_five_k and one_eight_four_k can be changed above
            if proj_keys[i] == "Total Monthly Cost":
                column_title = sht.api.UsedRange.Find("Budget")
                first_formula_cell = column_title.offset(last_row, 2)
                first_formula_cell.value = '=' + first_formula_cell.offset(1, 0).address + '/' + four_nine_five_k

                second_formula_cell = first_formula_cell.offset(1, 2)
                second_formula_cell.value = '=' + first_formula_cell.offset(1, 0).address + '/' + one_eight_four_k

            elif proj_keys[i] == "Project requestor":
                column_title = sht.api.UsedRange.Find("Resource requestor")

            else:
                column_title = sht.api.UsedRange.Find(proj_keys[i])
            insert_cell = column_title.offset(last_row, 1)
            insert_cell.value = proj_data[proj_keys[i]]

        for i in range (len(serv_keys)):
            column_title = sht.api.UsedRange.Find(serv_keys[i])
            insert_cell = column_title.offset(last_row, 1)
            insert_cell.value = serv_data[serv_keys[i]]

        wb.save()
        wb.close()

    def read_write(self, file_list):
        self.project_keys, self.service_keys, self.vm_keys = self.get_keys_from_init(self, self.location, self.project_keys, self.service_keys, self.vm_keys)
        
        for file in file_list:
            self.project_data, self.service_data, self.vmcore_data = self.read_from_excel(file, self.project_keys, self.service_keys, self.vm_keys)
            self.write_to_excel(self.location, self.overviewfile, self.project_keys, self.service_keys, self.project_data, self.service_data)