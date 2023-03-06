#本程序为计算有效年平均量程序
#注意：需完成主程序后并确认无空白行后使用
import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time

print("***按年平均气象数据计算程序启动***")
time.sleep(1)

#确定时间间隔标准
oneday=datetime(2000,1,2)-datetime(2000,1,1)
datetype=type(datetime(2000,1,1))

#统一获取城市名称
cityname=input("请输入你要计算数据的城市名：")
dataname=input("请输入要计算平均值的数据列序号（C/D/E/F/G）:")

#打开Excel
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

#打开文件并获取sheet1
wb=app.books.open('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
sheet=wb.sheets[str(cityname)]#处理的时候记得改城市名称

#确定新增列标
addnamef=sheet.range('A2').expand("right").columns.count
addname=chr(ord('A')+addnamef)

#新增/取消隐藏指定列
sheet.api.Columns(addname).EntireColumn.Hidden = False
sheet.api.Columns(chr(ord(addname)+1)).EntireColumn.Hidden = False

#写入列标题
sheet.range(addname+'1').value='年平均'
sheet.range(chr(ord(addname)+1)+'1').value=sheet.range(dataname+'1').value

#读取有效行数
starttime=sheet.range('B2').value#实际开始时间
startday=datetime(starttime.year,starttime.month,starttime.day)#时刻归零化开始时间
print('该城市气象数据记录开始时间：'+str(starttime))
rows1=sheet.used_range.last_cell.row#源文件总行数
for i in range(2,rows1+10):#为抵消空格所需多余循环次数，暂定为十次
   if type(sheet.range(str('B'+str(i))).value)!=datetype:
      break
i-=1
endtime=sheet.range(str('B'+str(i))).value#实际结束时间
endday=datetime(endtime.year,endtime.month,endtime.day)#时刻归零化结束时间
print('该城市气象数据记录结束时间：'+str(endtime))
rows=i#源文件有效行数
timegap=endday-startday
timeall=timegap+oneday#计算一次应有总时间，防备简阳循环爆表

#计算气温年平均
sheet.range(addname+'1').value='年平均'+sheet.range(dataname+'1').value
check=2
fortimes=int(str(timeall.days))+2
pbar=tqdm(total=int(int(timeall.days)/365))
number=2
while check < fortimes: 
  yearstart=sheet.range(str('B'+str(check))).value
  ystartday=datetime(yearstart.year,yearstart.month,yearstart.day)
  yr=ystartday.year
  list=[]
  flag=1
  while check<fortimes and sheet.range(str('B'+str(check))).value.year==yr:
    if sheet.range(dataname+str(check)).value!=999999:
       list.append(sheet.range(dataname+str(check)).value)
       check+=1
    else:
       list.append(sheet.range(dataname+str(check)).value)
       check+=1
       flag=0
  if flag==1:
    sheet.range(addname+str(number)).value=yr
    sheet.range(chr(ord(addname)+1)+str(number)).value=sum(list)/len(list)
    number+=1
  else:
    sheet.range(addname+str(number)).value=yr
    sheet.range(chr(ord(addname)+1)+str(number)).value=999999
    number+=1
  pbar.update(1)
pbar.close()
wb.save('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
wb.close()
app.quit()
print('***按年平均气象数据计算程序结束 数据保存在'+addname+'列中***')
print("\n")
