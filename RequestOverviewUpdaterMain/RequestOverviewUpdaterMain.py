import Find_File
import Read_Write_File

def main():
    
    #change loc to point at a local directory where the request forms and overview file are
    loc = r"C:\Users\shemchen\Desktop\excelPython"

    #make sure overviewfile matches the name of the overview file
    overviewfile = "NS-OCI_Resource Management-v2"
    
    file_list = []
    project_keys = []
    service_keys = []
    vmcore_keys = []

    project_data = {}
    service_data = {}
    vmcore_data = {}

    list_maker = Find_File.Find_File(loc, overviewfile)
    read_writer = Read_Write_File.Read_Write_File(loc, overviewfile)

    file_list = list_maker.find_new_files(loc, file_list)

    if file_list:
        read_writer.read_write(file_list)

    
    print(*file_list, sep = "\n")

if __name__ == "__main__":
    main()