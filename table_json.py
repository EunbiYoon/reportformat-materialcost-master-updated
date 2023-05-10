import pandas as pd
import openpyxl 
from json2html import *
import json
from pprint import pprint
from jsonmerge import merge
from jsonmerge import Merger

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/0130/result_0130.xlsx", sheet_name="TL")
read_excel.index=read_excel['Unnamed: 0']
read_excel=read_excel.drop(['Unnamed: 0'],axis=1)
extract_data=read_excel.T
extract_data=extract_data[["vs BOM","PO Price Change","Substitute Change"]]
extract_data["PO + Substitute"]=extract_data["PO Price Change"]+extract_data["Substitute Change"]
extract_data=extract_data.T

#column_list
column_list=list(extract_data.columns)

#index_list
index_list=list(extract_data.index)

#value_list
value_list=extract_data.values.round(1).tolist()

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'')
print("")
print(column_json)

print(extract_data.index)
# data,index json file format
data_json=str({"vs BOM":extract_data["vs BOM"],
               "PO Price Change":extract_data["PO Price Change"], 
               "Substitute Change":extract_data["Substitute Change"],
               "PO + Substitute":extract_data["PO + Substitute"]}).replace("{",'').replace("}",'')
print("")
print(data_json)


json_object={"columns":column_list, "index":index_list, "values":value_list}
jj=json.dumps(json_object)
print(jj)



