import smtplib
import datetime as dt
from email import encoders
from email.mime.base import MIMEBase
from socket import gaierror
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from function_def import load_workbook,search_column,search_row,get_maximum_row
import os

def get_key_value_from_excel(ws, keyColumn, valueColumn, item):
    c1 = search_column(ws,keyColumn)
    c2 = search_column(ws,valueColumn)
    r = search_row(ws,c1,item)
    val = ws.cell(row=r,column=c2).value
    return val

def get_count_with_status(ws,forColumn, status):
    c = search_column(ws,forColumn)
    r = get_maximum_row(ws)
    counter = 0
    for row in range(1,r+1):
        item = ws.cell(row=row, column=c)

        if (status in item.value):
            counter+=1
    return counter

def loadExcelFromDir(FileName):
    absFilePath = os.path.abspath(__file__)
    fileDir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(fileDir)
    for file in os.listdir(fileDir):
        if file == FileName:
            cpath = os.path.join(fileDir,file)
    wb = load_workbook(cpath)
    return wb

def send_mail(Recipients,wsResult):
    sender = "abdul-kadir.khan@capgemini.com"




