import pyodbc
import pandas as pd
import sys
from datetime import datetime
import warnings
import os

today = datetime.now()
warnings.filterwarnings('ignore')

#Database Credentials
server = 'your server name'
database = 'your database name'
username = 'your username'
password = 'your password'

#Establish Connection to SQL Database
print("Connecting to OBI Database...")
connection = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
print("Connection Successful")

#Setup SQL Queries
serialnum = "Enter your first query"
ewo = "Enter your second query"

#SERIAL NUMBER SEARCH
def serial_query(connection, serialnum):
    print("Executing Query...")
    cursor = connection.cursor()
    data = pd.read_sql(serialnum, connection, params = [SN])
    writer = pd.ExcelWriter('Line_Data_SN'+ SN + '_' + today.strftime('%Y_%m_%d')+'.xlsx', engine = 'xlsxwriter')
    data.to_excel(writer, sheet_name = 'Line 1')
    writer.save()
    print("\nQuery Successful, file has been generated.\nFile is stored here", os.getcwd())

#BATCH NUMBER SEARCH
def ewo_query(connection, ewo):
    print("\nExecuting Query...")
    cursor = connection.cursor()
    data = pd.read_sql(serialnum, connection, params = [EWO])
    writer = pd.ExcelWriter('Line_Data_SN'+ EWO + '_' + today.strftime('%Y_%m_%d')+'.xlsx', engine = 'xlsxwriter')
    data.to_excel(writer, sheet_name = 'Line 1')
    writer.save()
    print("\nQuery Successful, file has been generated.\nFile is stored here", os.getcwd())

#Select Query
select_query = input("\nPlease select a search option \nType '1' to search by  Device Serial Number \nType '2' to search by Batch Number \n\nSelect Option and Press Enter:")
if select_query ==("1"):
    SN = input("Enter Serial Number:")
    serial_query(connection, serialnum)
elif select_query ==("2"):
    EWO = input("Enter EWO Number:")
    ewo_query(connection, ewo)

input()
