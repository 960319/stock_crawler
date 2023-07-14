import xml.etree.ElementTree as ET
import requests
import time
from openpyxl import Workbook #匯入函式庫
def xml_to_dict(xml_file):
    # 解析XML檔案
    tree = ET.parse(xml_file)
    root = tree.getroot()

    data_dict = {}
    
    # 遍歷XML元素
    for child in root:
        tag = child.tag
        text = child.text
        data_dict[tag] = text
    
    return data_dict

def print_data(data_dict):
    # 顯示字典內容
    for key, value in data_dict.items():
        print(f"{key}: {value}")

def fillsheet(sheet,data,row): #兼立一個function名稱裡面放置三種參數sheet,data,row
    for column, value in enumerate(data,1): #讀取資料
        sheet.cell(row = row,column = column,value = value)
        #將資料放置在row行column列上，其格子裡填寫value資料
def  returnStrDAyList(startYear,startMonth,endYear,endMonth,day = "01"):
    result = []
    if startYear == endYear:
        for month in range(startMonth,endMonth+1):
            month = str(month)
            if len(month) == 1: #從1~9變01~09
                month = "0" + month
            result.append(str(startYear)+month+day)
        return result
    for year in range(startYear,endYear+1):
        if year == startYear:
            for month in range(startMonth,13):
                month = str(month)
                if len(month) == 1: #從1~9變01~09
                    month = "0" + month
                result.append(str(year)+month+day)
        elif year == endYear:
            for month in range(1,endMonth+1):
                month = str(month)
                if len(month) == 1: #從1~9變01~09
                    month = "0" + month
                result.append(str(year)+month+day)
        else:
            for month in range(1,13):
                month = str(month)
                if len(month) == 1: #從1~9變01~09
                    month = "0" + month
                result.append(str(year)+month+day)
    return result 

# 讀取data.xml並轉換為字典
data_dict = xml_to_dict('data.xml') #原來data_di+ct是字典
fields = ["日期","成交股數","成交量","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差"] #串列
wb = Workbook() #建立excel檔案
sheet = wb.active #讓excel啟動，建立第一個工作表格
sheet.title = "fields"
fillsheet(sheet, fields,1) #執行函式，注意參數要放對
startYear,startMonth = int(data_dict["startYear"]),int(data_dict["startMonth"])
endYear,endMonth = int(data_dict["endYear"]),int(data_dict["endMonth"])
#上面兩行為讀取字典裡的內容
yearList = returnStrDAyList(startYear,startMonth,endYear,endMonth) #執行函式
# print(yearList)
row = 2
for YearMonth in yearList:
    rq = requests.get(data_dict["url"],params={
        "response":"json",
        "date":YearMonth,
        "stockNo":data_dict["stockNo"]
    })
    jsonData = rq.json()
    dailyPriceList = jsonData.get("data",[])
    for dailyPrice in dailyPriceList:
        print(dailyPrice)
        fillsheet(sheet,dailyPrice,row)
        row+=1
    time.sleep(3)
name = data_dict["excelName"]
wb.save(name+".xlsx") #存檔


# 印出字典內容
print_data(data_dict)
