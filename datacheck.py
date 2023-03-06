#本分程序仅用于单独确定并记录气象数据异常的序号
#不具有修复功能
#与主程序功能有重叠，慎重使用

import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
print("***气象数据检验(单独检验)程序启动***")
time.sleep(1)

#确定时间间隔标准
oneday=datetime(2000,1,2)-datetime(2000,1,1)
datetype=type(datetime(2000,1,1))

#统一获取城市名称
cityname=input("请输入你要检测数据的城市名：")

#打开Excel
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
#打开文件并获取sheet1
wb=app.books.open('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
sheet=wb.sheets[str(cityname)]#处理的时候记得改城市名称


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


check=3
fortimes=int(str(timeall.days))+2#确定序号上限，防备简阳爆表

#不合理气象数据检测并记录
print('***正在向记事本写入异常天气数据***')
filename=cityname+'气象异常数据序号合集（单独检验）.txt'
file=open(filename,"w+")
pbar=tqdm(total=fortimes-1)
for i in range(2,fortimes):
   if sheet.range(str('C'+str(i))).value==999999 or sheet.range(str('D'+str(i))).value==999999 or sheet.range(str('E'+str(i))).value==999999 or sheet.range(str('F'+str(i))).value==999999 or sheet.range(str('G')+str(i)).value==999999:
      file.write(str(i)+' '+str(sheet.range(str('B'+str(i))).value)+'\n')
   pbar.update(1)
pbar.close()
file.close()
print('异常天气数据写入完成')
#关闭Excel   
wb.save('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
wb.close()
app.quit()
print('***气象数据检验(单独检验)程序结束 本次检验'+cityname+'异常天气数据***')
print(cityname+'异常天气已写入该目录下文本文件')
print('\n')