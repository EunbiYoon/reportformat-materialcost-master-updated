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
value_list=list(read_excel.values)

#json file format
merger=Merger(column_json)
result=merger.merge(column_json)
pprint(result)



