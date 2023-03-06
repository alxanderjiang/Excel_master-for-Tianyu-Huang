import xlwings as xw
from datetime import datetime
from os import system
from tqdm import tqdm
import time
import math

#统一获取城市名称
cityname=input("请输入你要计算WGBT数据的城市名：")
#打开Excel
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
#打开文件并获取s对应城市sheet
wb=app.books.open('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
sheet=wb.sheets[str(cityname)]
#获取温度，相对湿度数据位置
Taflag=0
Uflag=0
for i in 'ABCDEFGHIJKLMN':
    if sheet.range(i+'1').value=='TEM':
        Taname=i
        Taflag=1
        print(i)
    if sheet.range(i+'1').value=='RHU':
        Uname=i
        Uflag=1
        print(i)
if Taflag==0 or Uflag==0:
    print("未找到干球温度和相对湿度，请检查数据表标题是否完整后重新运行本程序")
    exit()            

#获取数据添加位置
addnamef=sheet.range('A2').expand("right").columns.count
addname=chr(ord('A')+addnamef)


##新增/取消隐藏指定列
sheet.api.Columns(addname).EntireColumn.Hidden = False

#计算并写入WGBT数据    
sheet.range(addname+'1').value='WGBT'
Pa=101325#大气压
i=2
pbar=tqdm()
while True:
    if sheet.range(Taname+str(i)).value==999999 or sheet.range(Uname+str(i)).value==999999:
        i+=1
    else:
        Ta=sheet.range(Taname+str(i)).value
        U=sheet.range(Uname+str(i)).value*0.01
        ETa=611.2*math.exp(17.67*Ta/(243.5+Ta))
        hvTa=2501+1.85*Ta
        Tw=Ta
        while True:
             ETw=611.2*math.exp(17.67*Tw/(243.5+Tw))
             hvTw=2501+1.85*Tw
             U1=Pa/ETa*(1-0.622*hvTa/(1.01*Tw-1.01*Ta+0.622*ETw/(Pa-ETw)*hvTw+0.622*hvTa))
             if abs(U1-U)<0.001:
                break
             Tw-=0.01
        Tw=round(Tw,1)
        WGBT=0.7*Tw+0.3*Ta
        sheet.range(addname+str(i)).value=round(WGBT,1)
        i+=1
    
    if type(sheet.range('B'+str(i)).value)!=type(datetime(2000,1,1)):
        sheet.range(addname+str(i)).value=999999
        break 
    # if sheet.range('B'+str(i)).value==None:
    #     break另一种方式，可以尝试     
    pbar.update(1)
pbar.close()

wb.save('test\副本北京14时温压湿水汽压19510101-20141130.xlsx')#文件路径
wb.close()
app.quit()