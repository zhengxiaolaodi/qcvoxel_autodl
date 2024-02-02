import xlrd
import pandas as pd
import os
import xlwt
import time
from datetime import datetime
import os
 
def get_files_in_folder(folder_path):
    file_list = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            file_list.append(file_path)
    return file_list
