#################### Modules ####################

import psycopg2
import pandas as pd
# import numpy as np
import xlsxwriter
from config import config
import os
import psutil
import random
import time
import multiprocessing
import connectorx as cx


#################### Functions ####################

def run_query(query):
    return pd.read_sql_query(query, conn)

def multiprocess_query(string):
    with Pool(1) as p:
        data_load = p.map(run_query, string)
    return data_load
    

def query(string):
    """
    Reading a sql query and change it into a DataFrame
    
    Input: a string that is a query script in SQL
    Output: a DataFrame
    """
    #connect db
    params = config()
    conn = psycopg2.connect(**params)
    # postgres_url = "postgresql://postgres:tehamd1324@172.16.9.60:5432/zzz_nguyentdh_test"
    
    #read query
    #### Time decreases significantly (half of the old method) and so does space (approximately 5MB)
    #### but got problem with ";" so still need to think more
    
    df = pd.read_sql(string, con=conn)
    # df = cx.read_sql(postgres_url, string)

    return df

def read_file(location):
    """
    Usually for speed since we won't need to input the SQL query code
    Input: A file path of SQL script (Text file is accepted)
    Output: A string which is a SQL query code for query() func above
    """ 
    #configurate file location
    if "\\" in location:
        location = location.replace("\\", "//")    
    
    ## For closing file purpose
    with open(location, 'r', encoding='utf-8') as f:
        # string = file.read().replace('\n', '')
        string = [i.strip().replace("\n", "") for i in f]
        # string = [i.strip() for i in f.readlines()]

    if len(string) > 1:
        joined_string =  \
        """ 
        {}
        """.format("\n".join(string))
        return joined_string
    return string[0]
    
def cleaned_whitespace(content):
    """
    Removing all whitespace in a string 
    Input: An array of string
    Output: An array of no-whitespace string
    """
    cleaned = []
    for i in content:
        i = list(i)
        for j in range(len(i)):
            if isinstance(i[j], str):
                i[j] = str(i[j]).strip()
        # i[3] = i[3].strip()
        i = tuple(i)
        cleaned.append(i)
    return cleaned

def create_dir(dir_name):
    """
    Create a directory (folder) for saving the Excel file
    Input: a directory name
    Output: None
    """
    if dir_name not in os.listdir(os.getcwd()):
        os.mkdir(dir_name)

def make_excel(content, filename):
    """
    Create our Excel file output (with optional format in the below Format section)
    Input: a DataFrame (usually from query() function), a filename
    Output: None
    """
      
    
    #Variables
    folder = "Excel_Output"
    rand = str(random.randint(0, 99999999999))
    
    #Find Excel cols, rows
    
    df_rows = cleaned_whitespace(content.values.tolist())
    df_cols = content.columns.values
    
    
    #Create Excel file
    create_dir(folder)
    ##Check duplicate filename
    if filename + ".xlsx" not in os.listdir(os.path.join(os.getcwd(), folder)):
        filename = folder + "/" + filename + ".xlsx"
    else:
        filename = folder + "/" + filename + "_" + rand + ".xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()     
    
    
    
    #Create format
    
    ###Note:
    # Can use Hex code for color
    cell_format = workbook.add_format()
    cell_format.set_bold()
    cell_format.set_font_color('840036')

    row_format = workbook.add_format()
    row_format.set_bold()
    row_format.set_font_color('644B37')
    row_format.set_bg_color('33FFF6')
    
    #Write to Excel (format if needed)
    
    ##Write columns
    for col_num, value in enumerate(df_cols):
        worksheet.write(0, col_num, value, row_format)

    ##Write rows
    
    
    # for i in df_rows:
        # for j in i:
            # worksheet.write(row, col, j, cell_format)
            # col += 1
        # col = 0
        # row += 1
    
    #### cleaner version
    row = 1
    for i in df_rows:
        for col, j in enumerate(i):
            worksheet.write(row, col, j, cell_format)
        row += 1
    
    workbook.close()



#################### Main ####################

if __name__ == "__main__":

    location = input()
    filename = input()
    
    
    start = time.time() ########### Execute time (Optimization purpose) less time -> better
    input_query = read_file(location)   
    result = query(input_query)
    end_query = time.time()
    # result = multiprocess_query("select * from company;")
    
    start_excel = time.time()
    make_excel(result, filename)
    end_excel = time.time()
    
    cwd = os.path.join(os.getcwd(), "Excel_Output\\")
    if filename + ".xlsx" not in os.listdir(cwd):
        print("Failed")
    else:
        print("Success!! - Check your Python_Automation/Excel_Output/ folder")
    
    
    end = time.time()
    print("Execution time:", end-start) ########### Execute time (Optimization purpose) less -> better
    print("Query time:", end_query-start) ########### Query time less -> better
    print("Excel Creation time:", end_excel-start_excel) ########### Create Excel file time less -> better
    process = psutil.Process(os.getpid()).memory_info().rss/1024.0/1024.0
    # print("{}MB".format(round(process.memory_info().rss/1024.0/1024.0, 3)))   
    print("{}MB".format(round(process, 3))) ###### Memory Usage (Optimization purpose) in megabytes less -> better
        
    