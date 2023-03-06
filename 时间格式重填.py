import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
print("***气象数据修复程序启动***")
time.sleep(1)

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
#以下开始更改时间格式
    print('现在更改'+cityname+'时间格式')
    check=2
    pbar=tqdm()
    while sheet.range('B'+str(check)).value!=None:
        checktime=datetime(int(sheet.range('B'+str(check)).value),int(sheet.range('C'+str(check)).value),int(sheet.range('D'+str(check)).value))
        sheet.range('B'+str(check)).value=checktime
        sheet.range('A'+str(check)).value=cityname
        check+=1
        pbar.update(1)
    pbar.close()
    sheet.range('B1').column_width=13
    sheet.range('C1').api.EntireColumn.Delete()
    sheet.range('C1').api.EntireColumn.Delete()
    sheet.range('C1').api.EntireColumn.Delete()

    sheet.range('B1').value='Datetime'
    sheet.range('A1').value='Station_name'
#保存Excel   
    wb.save(xlfilename)#文件路径

wb.close()
app.quit()
print('***程序结束***')