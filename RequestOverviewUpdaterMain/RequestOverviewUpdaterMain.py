import Find_File

def main():
    
    loc = r"C:\Users\shemchen\Desktop\excelPython"
    overviewfile = "NS-OCI_Resource Management-v2"
    
    file_list = []
    project_keys = []
    service_keys = []
    vmcore_keys = []

    project_data = {}
    service_data = {}
    vmcore_data = {}

    list_maker = Find_File.Find_File(loc, overviewfile)

    file_list = list_maker.find_new_files(loc, file_list)
    
    print(*file_list, sep = "\n")

if __name__ == "__main__":
    main()