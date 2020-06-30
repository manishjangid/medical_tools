import pandas as pd
import numpy as np
from os.path import basename
import os, collections, csv

data_frame_last0 = []
data_frame_last1 = []
data_frame_last3 = []
data_frame_lastminus1 = []
numberOfSheets = 4

cwd = os.path.abspath('')
print cwd
files = os.listdir(cwd) 
print files
for filen in files:
    print filen
    if filen.endswith('.XLS'):
        data = pd.read_excel(filen, Sheet1 = 'd-', header=0)
        print data.head(2)
        data_frame_lastminus1.append(data)
        data = pd.read_excel(filen, Sheet1 = 'd0', header=0)
        data_frame_last0.append(data)
        data = pd.read_excel(filen, Sheet1 = 'd1', header=0)
        data_frame_last1.append(data)
        data = pd.read_excel(filen, Sheet1 = 'd3', header=0)
        data_frame_last3.append(data)
        
final_list = ["last-1.xls", "last0.xls","last1.xls","last3.xls"]
data_frame_lastminus1 = pd.concat(data_frame_lastminus1)
data_frame_lastminus1.to_excel(final_list[0], index=False)
data_frame_last0 = pd.concat(data_frame_last0)
data_frame_last0.to_excel(final_list[1], index = False)
data_frame_last1 = pd.concat(data_frame_last1)
data_frame_last1.to_excel(final_list[2])
data_frame_last3 = pd.concat(data_frame_last3)
data_frame_last3.to_excel(final_list[3])
