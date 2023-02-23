#v0.4.0更新日志
#1.新增了主控程序，用户不再需要单独执行各.py文件
#2.完善了文本文档写入和平均值写入功能
#3.新增了分位点功能
#4.修复了若干已知问题
#5.                         开发者   2023.2.21    
import os
import time
print('***欢迎使用本软件***')
time.sleep(1)
while True:
  print('***软件菜单***\n1.气象数据修复\n2.异常气象数据记录\n3.按月有效平均气温计算\n4.按年有效平均气温计算\n5.分位点计算\n6.退出')
  a=input('请选择将要使用的功能：')
  if a=='1':
     os.system("python exceltest1.py")
  elif a=='2':
     os.system("python datacheck.py")
  elif a=='3':
     os.system("python monthaverage.py")
  elif a=='4':
     os.system("python yearaverage.py")
  elif a=='5':
     os.system("python 95check.py")
  else: break      
print("***软件退出 感谢您的使用***")