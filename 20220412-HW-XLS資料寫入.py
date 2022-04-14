"""
matplotlib
"""
import tkinter as tk
import matplotlib.pyplot as plt                            # 匯入matplotlib 的pyplot 類別，並設定為plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.image as mpimg
import time, threading
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']   # 步驟一（替換sans-serif字型）
plt.rcParams['axes.unicode_minus'] = False                 # 步驟二（解決座標軸負數的負號顯示問題）
import xlrd
import xlwt
"""
####繼續昨天作業####
透過讀取 covid19.xls 檔，
把資料讀取進來，並現在在6個 matplotlib
"""
read=xlrd.open_workbook('covid19.xls')    # 打開excel
sheet=read.sheets()[0]                    # 讀取[0]
print("多少筆：",sheet.nrows)               # 多少筆(列)
print("多少欄位：",sheet.ncols)             #  多少欄位

#取得日期資料
def getDate(sheet,n):
    date=[]
    for i in range(1,sheet.nrows):
        i=sheet.cell(i,n-1).value
        date.append(i)
    return date

date1=getDate(sheet,1)
print("date1=",date1)

#取得台北資料
def getTaipei(sheet,n):
    taipei=[]
    for i in range(1,sheet.nrows):
        i=sheet.cell(i,n-1).value
        taipei.append(i)
    return taipei

case3=getTaipei(sheet,2)
print("case3=",case3)

#取得新北市資料
def getNTaipei(sheet,n):
    nTaipei=[]
    for i in range(1,sheet.nrows):
        i=sheet.cell(i,n-1).value
        nTaipei.append(i)
    return nTaipei

case1=getNTaipei(sheet,3)
print("case1=",case1)

#取得桃園資料
def getToayuan(sheet,n):
    taoyuan=[]
    for i in range(1,sheet.nrows):
        i=sheet.cell(i,n-1).value
        taoyuan.append(i)
    return taoyuan

case2=getDate(sheet,4)
print("case2=",case2)

#取得城市名稱
label1=sheet.cell(0,2).value
print("label1：",label1)
label2=sheet.cell(0,3).value
print("label2：",label2)
label3=sheet.cell(0,1).value
print("label3：",label3)
label1="新北市全區"
label2='桃園市全區'
label3="台北市全區"
listDates1=[date1,date1,date1]
listCases1=[case1,case2,case3]
listlabels1=[label1,label2,label3]

win = tk.Tk()
win.title("Product Description")
win.resizable(True,True)
w=1200
h=800
win.minsize(300,280)
win.maxsize(w,h)

#第一張圖表
fig1 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig1, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=0,y=0) # 加到視窗中
ax=fig1.add_subplot(111) # 注意！一定要添加，用來指定位置

# 圖表
i=0
while i<len(listDates1):
    ax.plot(listDates1[i],listCases1[i],label=listlabels1[i],linewidth=2.0)
    i=i+1
ax.legend(loc ="upper left")
ax.set_title("New Covid Case")
ax.set_ylabel("Number Of New Case")                     # 顯示Y 座標的文字
ax.set_xlabel("Date of April")                                   # 顯示X 座標的文字

#第二張圖表
fig2 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig2, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=400,y=0) # 加到視窗中
ax=fig2.add_subplot(111) # 注意！一定要添加，用來指定位置

#bar chart
i=0
while i<len(listDates1):
    ax.bar(listDates1[i],listCases1[i], width=1, edgecolor="white", linewidth=0.7)
    i=i+1

ax.set_title("New Covid Case")
ax.set_ylabel("Number Of New Case")                     # 顯示Y 座標的文字
ax.set_xlabel("Date of April")                                   # 顯示X 座標的文字

#第三張圖表
fig3 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig3, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=800,y=0) # 加到視窗中
ax=fig3.add_subplot(111) # 注意！一定要添加，用來指定位置

#pie chart
#計算3三個行政區4/1-9新增確診總數
a = 0
totalCase1 = 0
totalCase2 = 0
totalCase3 = 0
while a <= 8:
    totalCase1 = totalCase1 + case1[a]
    totalCase2 = totalCase2 + case2[a]
    totalCase3 = totalCase3 + case3[a]
    a = a + 1
print("totalCase1=",totalCase1)
print("totalCase2=",totalCase2)
print("totalCase3=",totalCase3)
totalCase1=899
totalCase2=230
totalCase3=410

listTotalCases1=[totalCase1,totalCase2,totalCase3]
colors=["lightgreen","darkgreen","gold" ]

ax.pie(listTotalCases1,colors=colors,labels=listlabels1, radius=3, center=(4, 4),
       wedgeprops={"linewidth": 1, "edgecolor": "white"}, frame=True)
ax.legend(loc ="upper left")
ax.set_title("New Covid Case")
ax.set_ylabel("Number Of New Case")                     # 顯示Y 座標的文字
ax.set_xlabel("Date of April")

#第四張圖表
fig4 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig4, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=0,y=400) # 加到視窗中
ax=fig4.add_subplot(111) # 注意！一定要添加，用來指定位置

#scatter
plt.subplot(2,3,4)
sizes=[10,20,30,40,50,80,100,150,200]
colors=[30,60,90,120,150,180,200,220,255]  # 0~255

i=0
while i<len(listDates1):
    ax.scatter(listDates1[i],listCases1[i],s=sizes, c=colors)
    i=i+1

#第五張圖表
fig5 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig5, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=400,y=400) # 加到視窗中
ax=fig5.add_subplot(111) # 注意！一定要添加，用來指定位置

#grouped bar
labels=date1
menTotal=[41,40,42,55,71,92,90,82,85]
womenTotal= [55,42,47,64,74,92,75,87,81]

w=0.35
x1=[0,1,2,3,4,5,6,7,8]
x2=[0+w,1+w,2+w,3+w,4+w,5+w,6+w,7+w,8+w]

ax.bar(x1, menTotal, w, label='Men')
ax.bar(x2, womenTotal, w, label='Women')

ax.legend(loc ="upper left")
ax.set_title("New Covid Case")
ax.set_ylabel("Number Of New Case")                     # 顯示Y 座標的文字

#第六張圖表
fig6 = plt.Figure(figsize=(5, 5), dpi=80)# 指定繪製元件  with=80*5  高度=80*5
canvas = FigureCanvasTkAgg(fig6, win) # 使用FigureCanvasTkAgg 元件
canvas.get_tk_widget().place(x=800,y=400) # 加到視窗中
ax=fig6.add_subplot(111) # 注意！一定要添加，用來指定位置

i=0
while i<len(listDates1):
    ax.step(listDates1[i], listCases1[i], label=listlabels1[i])
    ax.plot(listDates1[i],listCases1[i],'o--', color='grey', alpha=0.3)
    i=i+1

ax.legend(loc ="upper left")
win.mainloop()

