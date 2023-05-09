import pandas as pd
import openpyxl 
from json2html import *
import json
from pprint import pprint
from jsonmerge import merge
from jsonmerge import Merger

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/graph.xlsx")
read_excel.index=read_excel['Unnamed: 0']
read_excel=read_excel.drop(['Unnamed: 0'],axis=1)

#column_list
column_list=list(read_excel.columns)

#index_list
index_list=list(read_excel.index)

#value_list
value_list=read_excel.values.round(1).tolist()

#json file format
json_object={"columns":column_list, "index":index_list, "values":value_list}
jj=json.dumps(json_object)
print(jj)



