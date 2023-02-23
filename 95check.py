import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
print("***按日95%分位点数据计算程序启动***")
time.sleep(1)

#确定时间间隔标准
oneday=datetime(2000,1,2)-datetime(2000,1,1)
datetype=type(datetime(2000,1,1))

#统一获取城市名称
cityname=input("请输入你要计算95%分位点数据的城市名：")
dataname=input("请输入要计算的数据列序号（C/D/E/F/G）:")

#打开Excel
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

#打开文件并获取s对应城市sheet
wb=app.books.open('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
sheet=wb.sheets[str(cityname)]

#确定新增列标
addnamef=sheet.range('A2').expand("right").columns.count
addname=chr(ord('A')+addnamef)

#新增/取消隐藏指定列
sheet.api.Columns(addname).EntireColumn.Hidden = False
sheet.api.Columns(chr(ord(addname)+1)).EntireColumn.Hidden = False

#写入列首行
sheet.range(addname+'1').value='0.95分位点'
sheet.range(chr(ord(addname)+1)+'1').value=sheet.range(dataname+'1').value

#确定日期格式和间隔
listyear=datetime(2000,1,1)
oneday=datetime(2000,1,2)-datetime(2000,1,1)
#构造可供计算的数组并初始化
valuelist = [[] for i in range(366)]

#遍历日期并排序
check=2
pbar=tqdm()
while True:
    checkday=sheet.range('B'+str(check)).value
    if type(checkday)!=datetype:
        break
    checkvalue=sheet.range(dataname+str(check)).value
    checktime=datetime(2000,checkday.month,checkday.day)
    days=(checktime-datetime(2000,1,1)).days
    if checkvalue!=999999:
       valuelist[days].append(checkvalue)
    check+=1
    pbar.update(1)
pbar.close

#覆写

timenow=listyear
pbar=tqdm()
for i in range(0,366):
    valuelist[i].sort()
    listlen=len(valuelist[i])
    result=listlen*0.95
    if type(result)==int:
        sheet.range(addname+str(i+2)).value=str(timenow.month)+'/'+str(timenow.day)
        sheet.range(chr(ord(addname)+1)+str(i+2)).value=valuelist[i][result-1]
    else:
        result=int(result)
        sheet.range(addname+str(i+2)).value=str(timenow.month)+'/'+str(timenow.day)
        sheet.range(chr(ord(addname)+1)+str(i+2)).value=(valuelist[i][result-1]+valuelist[i][result])/2   
    timenow+=oneday
    pbar.update(1)
pbar.close
print(cityname+'95%分位点数据已写入'+addname+'列及其后列')
#关闭并保存
wb.save('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
wb.close()
app.quit()