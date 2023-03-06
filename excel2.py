import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
print("***气象数据修复程序启动***")
time.sleep(1)
#确定时间间隔标准
oneday=datetime(2000,1,2)-datetime(2000,1,1)
datetype=type(datetime(2000,1,1))

#统一获取城市名称
# cityname=input("请输入你要修复数据的城市名：")

#获取文件路径
from tkinter import filedialog
from tkinter import *
root = Tk()
root.filename = filedialog.askopenfilename(initialdir="C:/", title="Select a File", filetypes=(("Xlsx files", "*.xlsx"), ("all files", "*.*")))
xlfilename=root.filename
root.destroy()
print("数据文件名：",xlfilename)

#改表
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
#打开文件并获取sheet1
wb=app.books.open(xlfilename)#文件路径
for sheet in wb.sheets:
   cityname=sheet.name

#读取有效行数
   starttime=datetime(int(sheet.range('B2').value),int(sheet.range('C2').value),int(sheet.range('D2').value))#实际开始时间
   startday=datetime(starttime.year,starttime.month,starttime.day)#时刻归零化开始时间
   print(cityname+'气象数据记录开始时间：'+str(starttime))
   rows1=sheet.used_range.last_cell.row#源文件总行数
   for i in range(2,rows1+10):#为抵消空格所需多余循环次数，暂定为十次
      if sheet.range(str('B'+str(i))).value==None:
         break
   i-=1
   endtime=datetime(int(sheet.range('B'+str(i)).value),int(sheet.range('C'+str(i)).value),int(sheet.range('D'+str(i)).value))#实际结束时间
   endday=datetime(endtime.year,endtime.month,endtime.day)#时刻归零化结束时间
   print(cityname+'气象数据记录结束时间：'+str(endtime))
   rows=i#源文件有效行数
   timegap=endday-startday
   timeall=timegap+oneday#计算一次应有总时间，防备简阳循环爆表
   
   if timeall.days==rows-1:
      wb.save(xlfilename)
      print(cityname+'气象数据无误')
      continue
   
   input("点击任意键开始修复气象数据")

#以下开始修复气象数据
   check=3
   fortimes=int(str(timeall.days))+2#确定序号上限，防备简阳爆表
#for i in range(3,fortimes):
   pbar=tqdm(total=fortimes)
   while check < fortimes: 
      cellnow=datetime(int(sheet.range('B'+str(check)).value),int(sheet.range('C'+str(check)).value),int(sheet.range('D'+str(check)).value))
      cellup=datetime(int(sheet.range('B'+str(check-1)).value),int(sheet.range('C'+str(check-1)).value),int(sheet.range('D'+str(check-1)).value))
      cellnowdate=datetime(cellnow.year,cellnow.month,cellnow.day)
      cellupdate=datetime(cellup.year,cellup.month,cellup.day)#时刻归零化当前位置时间
#   print(check)#为了粗暴地显示进度
      if cellup==endday:
         break #到达最终时间则结束循环   
      if cellnowdate-cellupdate==oneday:
         check=check+1
      elif cellnowdate==cellupdate:
         sheet.range('B'+str(check)).api.EntireRow.Delete()
         rows=rows-1
      else:
         sheet.api.Rows(check).Insert()
         sheet.range(str('A'+str(check))).value=sheet.range(str('A'+str(check-1))).value
         insertdate=cellup+oneday
         sheet.range('B'+str(check)).value=insertdate.year
         sheet.range('C'+str(check)).value=insertdate.month
         sheet.range('D'+str(check)).value=insertdate.day
         sheet.range('E'+str(check)).value=cityname
         sheet.range(str('F'+str(check))).value=[999999,999999,999999,999999,999999]
         check=check+1
         rows=rows+1 
      pbar.update(1)
   pbar.close()
   print("修复后表格中有天数:")
   print(rows-1)
   print('实际应有天数:')
   print(timeall.days)

#隐藏多余空白行
   print("***正在隐藏多余空白行***")
   rows2=sheet.used_range.last_cell.row
   print("现有空白行：")
   print(rows2-fortimes+1)
   sheet.api.Rows(str(fortimes)+':'+str(rows2)).EntireRow.Hidden=True
   print("隐藏完成，请注意重复检查是否仍有多余空白行")

#不合理气象数据检测并记录
   print('***正在向记事本写入异常天气数据***')
   filename=cityname+'气象异常数据序号合集.txt'
   file=open(filename,"w+")
   pbar2=tqdm(total=fortimes-1)
   for i in range(2,fortimes):
      if sheet.range(str('F'+str(i))).value==999999 or sheet.range(str('G'+str(i))).value==999999 or sheet.range(str('H'+str(i))).value==999999 or sheet.range(str('I'+str(i))).value==999999 or sheet.range('J'+str(i)).value==999999:
         file.write(str(i)+' '+str(sheet.range(str('B'+str(i))).value)+str(sheet.range('C'+str(i)).value)+str(sheet.range('D'+str(i)).value)+'\n')
      pbar2.update(1)   
   pbar2.close()
   file.close()
   print('异常天气数据写入完成')
#保存Excel   
   wb.save(xlfilename)#文件路径

wb.close()
app.quit()
print('***气象修复程序结束 本次修复'+cityname+'气象数据***')
print(cityname+"异常天气已写入该目录下文本文件")
print("\n")
     
