# coding:utf-8  #must be line 1
from openpyxl import Workbook
from openpyxl import load_workbook
import sys

reload(sys) 
sys.setdefaultencoding('gb18030') #for UnicodeDecodeError: 'ascii' codec can't decode byte 0xa1 in position 0: ordinal not in range(128)

orderType = ["新订","续订","复活"]


 

         




wb = load_workbook("order.xlsx")
ws_ttl = wb.get_sheet_by_name("ttl")
ws_cmb = wb.get_sheet_by_name("cmb")







#get  value
payerNameList = []
for row in tuple(ws_ttl.rows):
  payerName = row[4].value    
  if (payerName.encode('utf8') in orderType ):
     payerName = row[5].value  
  print("get payerName:%s "% payerName)
  payerNameList.append(payerName)

lineBegin  = 0
for row in tuple(ws_cmb.rows):
   if not paymentDesc :
      lineBegin = lineBegin + 1
   paymentDesc = row[6].value  
   paymentPrice = row[3].value  
   print("get payment price %s, desc %s "% (paymentPrice, paymentDesc))

print("cmb lineBegin %s "% (lineBegin))   

print orderType








 
