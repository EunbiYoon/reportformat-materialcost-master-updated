import pandas as pd
import openpyxl 
from json2html import *
import json
from pprint import pprint
from jsonmerge import merge
from jsonmerge import Merger

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/NPTGERP/0320/result_0320.xlsx", sheet_name="RV13D1AMAZU.ABWEUUS_result")
read_excel.index=read_excel['Unnamed: 0']
read_excel=read_excel.drop(['Unnamed: 0'],axis=1)
extract_data=read_excel.T
extract_data=extract_data[["vs BOM","PO Price Change","Substitute Change"]]
extract_data["PO + Substitute"]=extract_data["PO Price Change"]+extract_data["Substitute Change"]

#column_list
column_list=list(extract_data.index)

#index_list
value1=list(extract_data["vs BOM"].round(1))
value2=list(extract_data["PO Price Change"].round(1))
value3=list(extract_data["Substitute Change"].round(1))
value4=list(extract_data["PO + Substitute"].round(1))

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print("")
print(column_json)

# data,index json file format
data_json=str({"vs BOM":value1,
               "PO Price Change":value2, 
               "Substitute Change":value3,
               "PO + Substitute":value4}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','"nan"')
print(","+data_json)



