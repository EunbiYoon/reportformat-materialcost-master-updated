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
FL_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/result_0306.xlsx', sheet_name="F3U8CNU3W.ABWEUUS_result")
TL_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/result_0306.xlsx', sheet_name="T1889EFHUW.ABWEUUS_result")
DR_last_result=pd.read_excel('C:/Users/RnD Workstation/Documents/NPTGERP/0306/result_0306.xlsx', sheet_name="RV13D1AMAZU.ABWEUUS_result")


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

#np.nan -> blank
FL_blank=FL_round1.fillna('')
TL_blank=TL_round1.fillna('')
DR_blank=DR_round1.fillna('')


######html style
###FL
#html - table
FL_html=FL_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid black; border-collapse:collapse; font-family:sans-serif;">')
#column align center
FL_html=FL_html.replace("text-align: right;","text-align: center;")
#html - th,td
FL_html=FL_html.replace('<td>','<td style="width:63px; background-color:white; border:1px solid black; border-collapse: collapse;">')
FL_html=FL_html.replace('<th>','<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Production Qty</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Production Qty</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PAC</th>','<th style="padding:7px;color:black;background-color:rgb(191,191,191); border:1px solid black; border-collapse: collapse;">PAC</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Material Cost</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">vs BOM</th>','<th style="padding:7px;color:black;background-color:rgb(242,242,242); border:1px solid black; border-collapse: collapse;">vs BOM</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PO Price Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">PO Price Change</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Substitute Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Substitute Change</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>')
FL_html=FL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>')


###TL
#html - table
TL_html=TL_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid black; border-collapse:collapse; font-family:sans-serif;">')
#column align center
TL_html=TL_html.replace("text-align: right;","text-align: center;")
#html - th,td
TL_html=TL_html.replace('<td>','<td style="width:63px; background-color:white; border:1px solid black; border-collapse: collapse;">')
TL_html=TL_html.replace('<th>','<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Production Qty</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Production Qty</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PAC</th>','<th style="padding:7px;color:black;background-color:rgb(191,191,191); border:1px solid black; border-collapse: collapse;">PAC</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Material Cost</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">vs BOM</th>','<th style="padding:7px;color:black;background-color:rgb(242,242,242); border:1px solid black; border-collapse: collapse;">vs BOM</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PO Price Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">PO Price Change</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Substitute Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Substitute Change</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>')
TL_html=TL_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>')


###DR
#html - table
DR_html=DR_blank.to_html().replace('<table border="1" class="dataframe">','<table class="dataframe" style="border:1px solid black; border-collapse:collapse; font-family:sans-serif;">')
#column align center
DR_html=DR_html.replace("text-align: right;","text-align: center;")
#html - th,td
DR_html=DR_html.replace('<td>','<td style="width:63px; background-color:white; border:1px solid black; border-collapse: collapse;">')
DR_html=DR_html.replace('<th>','<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">')
#white row : production qty / bom material cost/ 
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Production Qty</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Production Qty</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">BOM Material Cost</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PAC</th>','<th style="padding:7px;color:black;background-color:rgb(191,191,191); border:1px solid black; border-collapse: collapse;">PAC</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Material Cost</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">vs BOM</th>','<th style="padding:7px;color:black;background-color:rgb(242,242,242); border:1px solid black; border-collapse: collapse;">vs BOM</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">PO Price Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">PO Price Change</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Substitute Change</th>','<th style="padding:7px;color:black;background-color:white; border:1px solid black; border-collapse: collapse;">Substitute Change</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Overhead Material Cost</th>')
DR_html=DR_html.replace('<th style="padding:7px;color:white;background-color:rgb(128,128,128); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>','<th style="padding:7px;color:black;background-color:rgb(217,217,217); border:1px solid black; border-collapse: collapse;">Defect Material Cost</th>')


#html - table
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
# msg['To']='iggeun.kwon@lge.com, sehee.aiello@lge.com, jacey.jung@lge.com, gilnam.lee@lge.com, steven.yang@lge.com, jajoon1.koo@lge.com, wolyong.ha@lge.com, dowan.han@lge.com'
# msg['Cc']='ethan.son@lge.com, jongseop.kim@lge.com, richard.song@lge.com, minhyoung.sun@lge.com, kitae3.park@lge.com, tg.kim@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#Subject 꾸미기
msg['Subject']='[테네시 재료비 관리 Task] 3월 2주차 BOM과 실제 생산 투입 재료비 차이 분석'

# html table attach
FL_attach = MIMEText(FL_html, "html")
TL_attach = MIMEText(TL_html, "html")
DR_attach = MIMEText(DR_html, "html")

msg.attach(MIMEText('<h4 style="font-weight:300;font-family:sans-serif; color:black">Dear All, <br/><br/>I would like to share TN Production Site 3 Main Model Material Cost Trend.<br/>Please refer to the attachment and below information.<br/>Thank you,<br/><br/></h4>','html'))
msg.attach(MIMEText('<h3 style="font-family:sans-serif; color:grey">Front Loader - F3U8CNU3W.ABWEUUS</h3>','html'))
msg.attach(FL_attach)
msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif; color:grey">Top Loader - T1889EFHUW.ABWEUUS</h3>','html'))
msg.attach(TL_attach)
msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif; color:grey">Dryer - RV13D1AMAZU.ABWEUUS</h3>','html'))
msg.attach(DR_attach)


#첨부 파일1
etcFileName='FL_BOM_Comparison_0306.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0306/BOM Comparison_FL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일2
etcFileName='TL_BOM_Comparison_0306.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0306/BOM Comparison_TL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일3
etcFileName='DR_BOM_Comparison_0306.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0306/BOM Comparison_DR.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#첨부 파일4
etcFileName='result_0306.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0306/result_0306.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")

