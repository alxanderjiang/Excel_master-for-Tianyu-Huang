import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
print("***全自动数据写入程序启动***")
time.sleep(1)
#确定时间间隔标准
oneday=datetime(2000,1,2)-datetime(2000,1,1)
datetype=type(datetime(2000,1,1))

#统一获取城市名称
cityname=input("请输入你要计算数据的城市名：")
#打开Excel
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
#打开文件并获取s对应城市sheet
wb=app.books.open('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
sheet=wb.sheets[str(cityname)]

#新增/取消隐藏指定列
sheet.api.Columns('H').EntireColumn.Hidden = False
sheet.api.Columns('I').EntireColumn.Hidden = False
sheet.api.Columns('J').EntireColumn.Hidden = False
sheet.api.Columns('K').EntireColumn.Hidden = False
sheet.api.Columns('L').EntireColumn.Hidden = False
sheet.api.Columns('M').EntireColumn.Hidden = False

sheet.api.Columns('N').EntireColumn.Hidden = False
sheet.api.Columns('O').EntireColumn.Hidden = False
sheet.api.Columns('P').EntireColumn.Hidden = False
sheet.api.Columns('Q').EntireColumn.Hidden = False
sheet.api.Columns('R').EntireColumn.Hidden = False
sheet.api.Columns('S').EntireColumn.Hidden = False

#以下为按月平均计算部分
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

#计算气温月平均气温
sheet.range('H1').value='月平均'
sheet.range('I1').value=sheet.range('C1').value
sheet.range('J1').value=sheet.range('D1').value
sheet.range('K1').value=sheet.range('E1').value
sheet.range('L1').value=sheet.range('F1').value
sheet.range('M1').value=sheet.range('G1').value
check=2
fortimes=int(str(timeall.days))+2
pbar=tqdm(total=int(int(timeall.days)/30))
number=2
while check < fortimes: 
  monthstart=sheet.range(str('B'+str(check))).value
  monstartday=datetime(monthstart.year,monthstart.month,monthstart.day)
  mon=monstartday.month
  list1=[]
  list2=[]
  list3=[]
  list4=[]
  list5=[]
  flag1=1
  flag2=1
  flag3=1
  flag4=1
  flag5=1
  
  while check<fortimes and sheet.range(str('B'+str(check))).value.month==mon:
    if sheet.range('C'+str(check)).value!=999999:
       list1.append(sheet.range('C'+str(check)).value)
    else:
       list1.append(sheet.range('C'+str(check)).value)
       flag1=0
    
    if sheet.range('D'+str(check)).value!=999999:
       list2.append(sheet.range('D'+str(check)).value)
    else:
       list2.append(sheet.range('D'+str(check)).value)
       flag2=0
    
    if sheet.range('E'+str(check)).value!=999999:
       list3.append(sheet.range('E'+str(check)).value)
    else:
       list3.append(sheet.range('E'+str(check)).value)
       flag3=0
    
    if sheet.range('F'+str(check)).value!=999999:
       list4.append(sheet.range('F'+str(check)).value)
    else:
       list4.append(sheet.range('F'+str(check)).value)
       flag4=0

    if sheet.range('G'+str(check)).value!=999999:
       list5.append(sheet.range('G'+str(check)).value)
    else:
       list5.append(sheet.range('G'+str(check)).value)
       flag5=0      
    check+=1       
  
  sheet.range('H'+str(number)).value=str(monstartday.year)+'/'+str(monstartday.month)
  
  if flag1==1:
    sheet.range('I'+str(number)).value=sum(list1)/len(list1)
    
  else:
    sheet.range('I'+str(number)).value=999999
    
  
  if flag2==1:
    sheet.range('J'+str(number)).value=sum(list2)/len(list2)
    
  else:
    sheet.range('J'+str(number)).value=999999
    
  
  if flag3==1:
    sheet.range('K'+str(number)).value=sum(list3)/len(list3)
    
  else:
    sheet.range('K'+str(number)).value=999999
      
  
  if flag4==1:
    sheet.range('L'+str(number)).value=sum(list4)/len(list4)
    
  else:
    sheet.range('L'+str(number)).value=999999

  if flag5==1:
    sheet.range('M'+str(number)).value=sum(list5)/len(list5)   
  else:
    sheet.range('M'+str(number)).value=999999
  
  number+=1
  pbar.update(1)
pbar.close()
print("所有按月平均数据计算完成，下面计算按年平均数据")
time.sleep(3)

#以下为年平均计算
#计算气温年平均气温
sheet.range('N1').value='年平均'
sheet.range('O1').value=sheet.range('C1').value
sheet.range('P1').value=sheet.range('D1').value
sheet.range('Q1').value=sheet.range('E1').value
sheet.range('R1').value=sheet.range('F1').value
sheet.range('S1').value=sheet.range('G1').value
check=2
fortimes=int(str(timeall.days))+2
pbar=tqdm(total=int(int(timeall.days)/365))
number=2
while check < fortimes: 
  yearstart=sheet.range(str('B'+str(check))).value
  yrstartday=datetime(yearstart.year,yearstart.month,yearstart.day)
  yr=yrstartday.year
  list1=[]
  list2=[]
  list3=[]
  list4=[]
  list5=[]
  flag1=1
  flag2=1
  flag3=1
  flag4=1
  flag5=1
  while check<fortimes and sheet.range(str('B'+str(check))).value.year==yr:
    if sheet.range('C'+str(check)).value!=999999:
       list1.append(sheet.range('C'+str(check)).value)
    else:
       list1.append(sheet.range('C'+str(check)).value)
       flag1=0
    
    if sheet.range('D'+str(check)).value!=999999:
       list2.append(sheet.range('D'+str(check)).value)
    else:
       list2.append(sheet.range('D'+str(check)).value)
       flag2=0
    
    if sheet.range('E'+str(check)).value!=999999:
       list3.append(sheet.range('E'+str(check)).value)
    else:
       list3.append(sheet.range('E'+str(check)).value)
       flag3=0
    
    if sheet.range('F'+str(check)).value!=999999:
       list4.append(sheet.range('F'+str(check)).value)
    else:
       list4.append(sheet.range('F'+str(check)).value)
       flag4=0    
    
    if sheet.range('G'+str(check)).value!=999999:
       list5.append(sheet.range('G'+str(check)).value)
    else:
       list5.append(sheet.range('G'+str(check)).value)
       flag5=0  
    
    check+=1 
  
  sheet.range('N'+str(number)).value=str(yrstartday.year)
  
  if flag1==1:
    sheet.range('O'+str(number)).value=sum(list1)/len(list1)
    
  else:
    sheet.range('O'+str(number)).value=999999
    
  
  if flag2==1:
    sheet.range('P'+str(number)).value=sum(list2)/len(list2)
    
  else:
    sheet.range('P'+str(number)).value=999999
    
  
  if flag3==1:
    sheet.range('Q'+str(number)).value=sum(list3)/len(list3)
    
  else:
    sheet.range('Q'+str(number)).value=999999
     
  
  if flag4==1:
    sheet.range('R'+str(number)).value=sum(list4)/len(list4)
    
  else:
    sheet.range('R'+str(number)).value=999999

  if flag5==1:
    sheet.range('S'+str(number)).value=sum(list5)/len(list5)

  else:
    sheet.range('s'+str(number)).vaule=999999   
  
  number+=1
  pbar.update(1)
pbar.close()
print("所有按年平均数据计算完成")

wb.save('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
wb.close()
app.quit()