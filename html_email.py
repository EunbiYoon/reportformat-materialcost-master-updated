import pandas as pd
import matplotlib.pyplot as plt
from itertools import repeat
import smtplib
import email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
from email.mime.application import MIMEApplication 
import os 


#####지난번에 했던 결과 소환
FL_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="F3P2CYUBW.ABWEUUS_result")
TL_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="T1889EFHUW.ABWEUUS_result")
DR_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="RV13D1AMAZU.ABWEUUS_result")

FL_item_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="F3P2CYUBW.ABWEUUS_worst item")
TL_item_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="T1889EFHUW.ABWEUUS_worst item")
DR_item_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', sheet_name="RV13D1AMAZU.ABWEUUS_worst item")

######데이터 정리
#index
FL_last_result.index=FL_last_result["Unnamed: 0"].values
FL_last_result=FL_last_result.drop(["Unnamed: 0"],axis=1)
TL_last_result.index=TL_last_result["Unnamed: 0"].values
TL_last_result=TL_last_result.drop(["Unnamed: 0"],axis=1)
DR_last_result.index=DR_last_result["Unnamed: 0"].values
DR_last_result=DR_last_result.drop(["Unnamed: 0"],axis=1)


#소숫점 1자리 맞춰주기
FL_round1=FL_last_result.round(1)
TL_round1=TL_last_result.round(1)
DR_round1=DR_last_result.round(1)

FL_item_result=FL_item_result.round(1)
TL_item_result=TL_item_result.round(1)
DR_item_result=DR_item_result.round(1)

#np.nan -> blank
FL_blank=FL_round1.fillna('')
TL_blank=TL_round1.fillna('')
DR_blank=DR_round1.fillna('')

FL_item_result=FL_item_result.fillna('')
TL_item_result=TL_item_result.fillna('')
DR_item_result=DR_item_result.fillna('')

################## Trend Table ##################
############ FL ############
FL_html=FL_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="padding:7px; border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
FL_html=FL_html.replace("text-align: right;","text-align: center;")
#html - th,td
FL_html=FL_html.replace('<td>','<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">')
FL_html=FL_html.replace('<th>','<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Production Qty</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Production Qty</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>','<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Material Cost</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">vs BOM</th>','<th style="color:black;background-color:rgb(242,242,242); border:1px solid grey; border-collapse: collapse;">vs BOM</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Price Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>')

#merge cell - column
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;"></th>','<th colspan="2" style="color:navy; background-color:#EFE6FF">Index</th>')

#merge cell - row1
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NaN</th>','')
FL_html=FL_html.replace('<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>','<th rowspan="4" style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>')

#merge cell - row2
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NaN</th>','')
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th rowspan="3" style="color:navy;background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')

# added index
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Result</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Result</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Net</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Total</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Total</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Overhead</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Overhead</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Defect</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Defect</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Price Change</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>')
FL_html=FL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>','<td style="font-weight:550; background-color:#C1FF00; color:navy; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>')

# key index color -> the other 3 changed upper line
FL_html=FL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>','<th style="color:white;background-color:#5C00FE; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>')
FL_html=FL_html.replace('<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#5C00FE; color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')

############ TL ############
TL_html=TL_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="padding:7px; border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
TL_html=TL_html.replace("text-align: right;","text-align: center;")
#html - th,td
TL_html=TL_html.replace('<td>','<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">')
TL_html=TL_html.replace('<th>','<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Production Qty</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Production Qty</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>','<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Material Cost</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">vs BOM</th>','<th style="color:black;background-color:rgb(242,242,242); border:1px solid grey; border-collapse: collapse;">vs BOM</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Price Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>')

#merge cell - column
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;"></th>','<th colspan="2" style="color:navy; background-color:#EFE6FF">Index</th>')

#merge cell - row1
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NaN</th>','')
TL_html=TL_html.replace('<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>','<th rowspan="4" style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>')

#merge cell - row2
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NaN</th>','')
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th rowspan="3" style="color:navy;background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')

# added index
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Result</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Result</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Net</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Total</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Total</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Overhead</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Overhead</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Defect</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Defect</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Price Change</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>')
TL_html=TL_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>','<td style="font-weight:550; background-color:#C1FF00; color:navy; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>')

# key index color -> the other 3 changed upper line
TL_html=TL_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>','<th style="color:white; background-color:#5C00FE; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>')
TL_html=TL_html.replace('<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#5C00FE; color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')


############ DR ############
#html - table
DR_html=DR_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="padding:7px; border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
DR_html=DR_html.replace("text-align: right;","text-align: center;")
#html - th,td
DR_html=DR_html.replace('<td>','<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">')
DR_html=DR_html.replace('<th>','<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Production Qty</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Production Qty</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">BOM Material Cost</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>','<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Material Cost</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">vs BOM</th>','<th style="color:black;background-color:rgb(242,242,242); border:1px solid grey; border-collapse: collapse;">vs BOM</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Price Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>','<th style="color:black;background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Overhead Material Cost</th>')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>','<th style="color:black;background-color:rgb(217,217,217); border:1px solid grey; border-collapse: collapse;">Defect Material Cost</th>')

#merge cell - column
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;"></th>','<th colspan="2" style="color:navy; background-color:#EFE6FF">Index</th>')

#merge cell - row1
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NaN</th>','')
DR_html=DR_html.replace('<th style="color:black;background-color:rgb(191,191,191); border:1px solid grey; border-collapse: collapse;">PAC</th>','<th rowspan="4" style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">PAC</th>')

#merge cell - row2
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapsecollapse;">NaN</th>','')
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th rowspan="3" style="color:navy; background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')

# added index
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Result</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Result</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Net</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Total</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Total</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Overhead</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Overhead</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Defect</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Defect</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Price Change</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>','<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">Substitute Change</td>')
DR_html=DR_html.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>','<td style="font-weight:550; background-color:#C1FF00; color:navy; border:1px solid grey; border-collapse: collapse;">Price + Substitute</td>')

# key index color -> the other 3 changed upper line
DR_html=DR_html.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>','<th style="color:white; background-color:#5C00FE; border:1px solid grey; border-collapse: collapse;">BOM vs PAC</th>')
DR_html=DR_html.replace('<td style="font-weight:550; background-color:#EFE6FF; color:navy; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>','<td style="font-weight:550; background-color:#5C00FE; color:white; border:1px solid grey; border-collapse: collapse;">PAC Net - BOM Net</td>')


################## Item Table ##################
############ FL ############
#html - table
FL_item=FL_item_result.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
FL_item=FL_item.replace("text-align: right;","text-align: center;")
FL_item=FL_item.replace('<td>','<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">')
FL_item=FL_item.replace('<th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">')

#remove unamed for colspan
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 2</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 3</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 4</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 5</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 6</th>','')

FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 8</th>','')

FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 10</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 11</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 12</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 13</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 14</th>','')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 15</th>','')

#remove unamed for colspan
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT</th>','<th colspan="7" style="color:white; background-color:#FF009B; border:1px solid grey; border-collapse: collapse;">NPT</th>')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th colspan="2" style="color:navy;background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">GERP</th>','<th colspan="7" style="color:white;background-color:#00C5FF; border:1px solid grey; border-collapse: collapse;">GERP</th>')

#remove column
FL_item=FL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NaN</th>','<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;"></th>')

#color in the npt vs gerp
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Seq</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Seq</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Level</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Level</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Parent Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Parent Part</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Child Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Child Part</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Description</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Description</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Qty</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Qty</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">Price</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Price</td>')

FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">match</td>','<td style="background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">match</td>')
FL_item=FL_item.replace('<td style= "background-color:white; border:1px solid grey; border-collapse: collapse;">price match</td>','<td style="background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">price match</td>')

#index 0 delete
FL_item=FL_item.replace('<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;">0</th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;"></th>')


############ TL ############
#html - table
TL_item=TL_item_result.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
TL_item=TL_item.replace("text-align: right;","text-align: center;")
TL_item=TL_item.replace('<td>','<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">')
TL_item=TL_item.replace('<th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">')

#remove unamed for colspan
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 2</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 3</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 4</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 5</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 6</th>','')

TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 8</th>','')

TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 10</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 11</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 12</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 13</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 14</th>','')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 15</th>','')

#remove unamed for colspan
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT</th>','<th colspan="7" style="color:white; background-color:#FF009B; border:1px solid grey; border-collapse: collapse;">NPT</th>')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th colspan="2" style="color:navy;background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">GERP</th>','<th colspan="7" style="color:white; background-color:#00C5FF; border:1px solid grey; border-collapse: collapse;">GERP</th>')

#remove column
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NaN</th>','<th style="color:navy;background-color:#EFE6FF; border:1px solid grey; border-collapse: collapse;"></th>')

#color in the npt vs gerp
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Seq</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Seq</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Level</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Level</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Parent Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Parent Part</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Child Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Child Part</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Description</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Description</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Qty</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Qty</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Price</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">match</td>','<td style= "background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">match</td>')
TL_item=TL_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">price match</td>','<td style= "background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">price match</td>')

#index 0 delete
TL_item=TL_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">0</th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;"></th>')


############ DR ############
#html - table
DR_item=DR_item_result.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid grey; border-collapse:collapse; font-family:Arial Narrow;">')
#column align center
DR_item=DR_item.replace("text-align: right;","text-align: center;")
DR_item=DR_item.replace('<td>','<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">')
DR_item=DR_item.replace('<th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">')

#remove unamed for colspan
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 1</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 2</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 3</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 4</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 5</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 6</th>','')

DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 8</th>','')

DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 10</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 11</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 12</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 13</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 14</th>','')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">Unnamed: 15</th>','')

#remove unamed for colspan
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT</th>','<th colspan="7" style="color:white;background-color:#FF009B; border:1px solid grey; border-collapse: collapse;">NPT</th>')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>','<th colspan="2" style="color:navy;background-color:#C1FF00; border:1px solid grey; border-collapse: collapse;">NPT vs GERP</th>')
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">GERP</th>','<th colspan="7" style="color:white; background-color:#00C5FF; border:1px solid grey; border-collapse: collapse;">GERP</th>')

#remove column
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">NaN</th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;"></th>')

#color in the npt vs gerp
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Seq</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Seq</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Level</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Level</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Parent Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Parent Part</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Child Part</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Child Part</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Description</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Description</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Qty</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Qty</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">Price</td>','<td style= "background-color:#ECFFAF; color:navy; font-weight:550; border:1px solid grey; border-collapse: collapse;">Price</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">match</td>','<td style= "background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">match</td>')
DR_item=DR_item.replace('<td style="background-color:white; border:1px solid grey; border-collapse: collapse;">price match</td>','<td style= "background-color:#929292; color:white; font-weight:550; border:1px solid grey; border-collapse: collapse;">price match</td>')

#index 0 delete
DR_item=DR_item.replace('<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;">0</th>','<th style="color:navy;background-color:#ECFFAF; border:1px solid grey; border-collapse: collapse;"></th>')

################################## Send email ################################## 
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
# msg['To']='iggeun.kwon@lge.com, incheol.kang@lge.com, sehee.aiello@lge.com, jacey.jung@lge.com, gilnam.lee@lge.com, steven.yang@lge.com, jajoon1.koo@lge.com, wolyong.ha@lge.com, dowan.han@lge.com'
# msg['Cc']='ethan.son@lge.com, jongseop.kim@lge.com, richard.song@lge.com, minhyoung.sun@lge.com, kitae3.park@lge.com, tg.kim@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#Subject 꾸미기
msg['Subject']='[테네시 재료비 관리 Task] 5월 4주차 BOM과 실제 생산 투입 재료비 차이 분석'

# html table attach
FL_attach = MIMEText(FL_html, "html")
TL_attach = MIMEText(TL_html, "html")
DR_attach = MIMEText(DR_html, "html")
FL_attach_item = MIMEText(FL_item, "html")
TL_attach_item = MIMEText(TL_item, "html")
DR_attach_item = MIMEText(DR_item, "html")

msg.attach(MIMEText('<h4 style="font-weight:300;font-family:Arial Narrow; color:black">Dear All, <br/><br/>I would like to share TN Production Site 3 Main Model Material Cost Trend.<br/>Please refer to the attachment and below information.<br/>Thank you,<br/><br/></h4>','html'))

msg.attach(MIMEText('<h3 style="font-family:Arial Narrow; color:grey">Front Loader - F3P2CYUBW.ABWEUUS</h3>','html'))
msg.attach(FL_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(FL_attach_item)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:Arial Narrow; color:grey">Top Loader - T1889EFHUW.ABWEUUS</h3>','html'))
msg.attach(TL_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(TL_attach_item)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:Arial Narrow; color:grey">Dryer - RV13D1AMAZU.ABWEUUS</h3>','html'))
msg.attach(DR_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(DR_attach_item)



#첨부 파일1
etcFileName='FL_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_FL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-DisPricesition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일2
etcFileName='TL_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_TL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-DisPricesition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일3
etcFileName='DR_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_DR.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-DisPricesition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#첨부 파일4
etcFileName='result_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-DisPricesition','attachment', filename=etcFileName)
    msg.attach(etcPart)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
