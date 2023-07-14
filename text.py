import requests
import time
import xml.etree.ElementTree as ET
from openpyxl import Workbook
with open("shiniecute.xml", "r") as file:
    xml_data=file.read()

root = ET.fromstring(xml_data)
paramss = {}
for element in root:
    paramss[element.tag] = element.text
for key, value in paramss.items():
    print(f"{key}: {value}")
def fillSheet(sheet,data,row):
    for column, value in enumerate(data,1):
        sheet.cell(row=row,column=column,value=value)
def returnStrDayList(sY,sM,eY,eM,d="01"):
    result=[]
    if sY==eY:
        for month in range(sM,eM+1):
            month=str(month)
            if len(month)==1:
                month="0"+month
            result.append(str(sY)+month+d)
        return result
    for year in range(sY,eY+1):
        if year==sY:
            for month in range(sM,13):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+d)
        elif year==eY:
            for month in range(1,eM+1):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+d)
        else:
            for month in range(1,13):
                month=str(month)
                if len(month)==1:
                    month="0"+month
                result.append(str(year)+month+d)
    return result
fields=["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"]
wb=Workbook()
sheet=wb.active
sheet.title="fields"
fillSheet(sheet,fields,1)
sY,sM=int(paramss["startYear"]),int(paramss["startMonth"])
eY,eM=int(paramss["endYear"]),int(paramss["endMonth"])
yearList=returnStrDayList(sY,sM,eY,eM)
row=2
for yearmonth in yearList:
    rq=requests.get(paramss["url"],params={
        "response":"json",
        "date": yearmonth,
        "stockNo": paramss["stockNo"]
    })
    jsonData=rq.json()
    dailyPriceList=jsonData.get("data",[])
    for dailyPrice in dailyPriceList:
        fillSheet(sheet,dailyPrice,row)
        row+=1
    time.sleep(3)
name=paramss["excelName"]
wb.save(name+".xlsx")
print("done :>")