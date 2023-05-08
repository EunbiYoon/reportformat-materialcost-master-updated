import pandas as pd
import openpyxl 
from json2html import *
import json
from pprint import pprint
from jsonmerge import merge
from jsonmerge import Merger

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/graph.xlsx")

#column_list
column_list=list(read_excel.columns)
column_json=json.dumps(column_list)

#index_list
index_list=list(read_excel['Unnamed: 0'])
index_json=json.dumps(index_list)

#value_list
value_list=str(list(read_excel.values)).replace('array(', '').replace(", dtype=object)","")

#json file format
columns=column_list
rows=index_list
values=value_list
json_object={"columns":columns, "index":rows, "values":value_list}
jj=json.dumps(json_object)
print(jj)



