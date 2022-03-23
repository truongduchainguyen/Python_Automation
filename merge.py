import pandas as pd
import os

def read_file(location):
    cwd = os.getcwd()
    os.chdir(location)
    list_dir = os.listdir(os.getcwd())
    file_merge = pd.read_csv(list_dir[0])
    for i in list_dir:
        file = pd.read_csv(i)
        file_merge.append(file)

    os.chdir("D:/Python_Automation/")
    file_merge.to_excel('output_merge.xlsx', engine='xlsxwriter', index=False)
    
if __name__ == "__main__":
    directory = input()
    read_file(directory)