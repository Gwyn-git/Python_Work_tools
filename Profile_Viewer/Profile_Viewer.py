import pandas as pd
import os
from tkinter import messagebox
from tkinter import *
import tkinter.font
import tkinter.ttk
import matplotlib.pyplot as plt
import matplotlib.style as mplstyle
import time
import datetime
import sys


#한글사용 세팅
from matplotlib import font_manager, rc
#font_path = "C:/Windows/Fonts/gulim.ttc"
#font = font_manager.FontProperties(fname=font_path).get_name()
#rc('font', family=font)


##### 함수 선언 #####


# 폴더 내 파일 리스트 확인
#path = os.path.abspath('')[0:-7]
path = os.path.abspath(__file__)
path = path.replace('Viewer/Profile_Viewer.py','')
list_l = os.listdir(path)                                         # 폴더 내 파일 리스트
list_xl = [file for file in list_l if file.endswith(".xlsx")]     # 폴더 내 파일리스트 중 xlsx 파일 리스트
list_n = []
for c in range(len(list_xl)):
    list_n.append(list_xl[c][:-5])                              # 확장자 없는 파일명 리스트


# 파일 로딩
def fileopen(event):
    global df, sd, ed                                           # 모든 함수에서 접근 가능한 df(불러온 로그 전체), sd,ed(로그 내 시작,종료일) 선언
    slabel_v.set('                             Loading...')
    window.update()
    df = pd.read_excel(path + '/' + chamberbox.get() + '.xlsx') # 선택한 파일 열기
    sd = df.iloc[0, 0]                                      
    ed = df.iloc[-1, 0]
    slabel_v.set('Period : ' + str(sd) + ' ~ ' + str(ed))                 # 파일 저장 기간 표기
    if sd.year == ed.year:                                      # 파일 내부 시작 연도 & 끝 연도 비교
        yearsbox['values'] = (sd.year)                          # 같으면 하나만
        yearebox['values'] = (sd.year) 
    else:
        yearsbox['values'] = (sd.year,ed.year)                 # 다르면 두개 다
        yearebox['values'] = (sd.year,ed.year)
        
    if sd.month == ed.month:                                      # 파일 내부 시작 월 & 끝 월 비교
        monthsbox['value'] = (sd.month)                          # 같으면 하나만
        monthebox['value'] = (sd.month)
    else:
        monthsbox['value'] = (sd.month,ed.month)                  # 다르면 두개 다
        monthebox['value'] = (sd.month,ed.month)
    ystext.set(sd.year)
    mstext.set(sd.month)
    dstext.set(sd.day)
    hstext.set(sd.hour)
    if str(int(str(sd.minute)[0:1])+1)+'0' != '60':
        bstext.set(str(int(str(sd.minute)[0:1])+1)+'0')
    else:
        hstext.set(int(sd.hour)+1)
        bstext.set('00')
    yetext.set(ed.year)
    metext.set(ed.month)
    detext.set(ed.day)
    hetext.set(ed.hour)
    betext.set(str(ed.minute)[0:1]+'0')


#조건 만족 확인
def check():
    start = 0                                                                                                                                                                                   #start값 0으로 세팅
    s_stamp = datetime.datetime(int(yearsbox.get()), int(monthsbox.get()), int(sdaybox.get()), int(stimebox.get()), int(sminbox.get()), 0)
    e_stamp = datetime.datetime(int(yearebox.get()), int(monthebox.get()), int(edaybox.get()), int(etimebox.get()), int(eminbox.get()), 0)
    
    if chamberbox.get() == '' or yearsbox.get() == '' or monthsbox.get() == '' or sdaybox.get() == '' or stimebox.get() == '' or sminbox.get() == '' or yearebox.get() == '' or monthebox.get() == '' or edaybox.get() == '' or etimebox.get() == '' or eminbox.get() == '':                      #조건 중 비어있는 값이 있는지 확인
        messagebox.showwarning('Error','필요 조건을 모두 입력하세요')
        sys.exit()
    elif s_stamp > e_stamp:                                                                                 #년+월+일을 합쳐서 숫자로 바꾼 다음 시작,종료일 비교
        messagebox.showwarning('Error', '종료일이 시작일보다 빠를 수 없습니다')
        sys.exit()
    elif s_stamp == e_stamp or s_stamp >= e_stamp:     #년+월+일이 같은 경우 시간비교
        messagebox.showwarning('Error', '종료시간이 시작시간보다 같거나 빠를 수 없습니다')
        sys.exit()
    else:
        global start_i, end_i, sdf
        #start_i = int(df.index[df['Date_Time'] == s_stamp])
        #start_i = df[df['Date_Time'] == s_stamp].index[0]

        try:
            start_i = df[df['Date_Time'] == s_stamp].index[0]
        except:
            messagebox.showwarning('Error', '시작일 Data가 존재하지 않습니다.')
            sys.exit()
        try:
            end_i = df[df['Date_Time'] == e_stamp].index[0]

            sdf = df[(df['Date_Time'] >= s_stamp) & (df['Date_Time'] <= e_stamp)]

            
        except:
            messagebox.showwarning('Error', '종료일 Data가 존재하지 않습니다.')
            sys.exit()

        
        start = 1
        return start                                                                                                                                                                            #모든 에러가 없을 경우에 start를 1로 변경하고 반환



#엑셀 추출
def export():
    start = check()                                                                                                                                                                     #check함수 실행
    if start == 1:
        try:
            sdf.to_excel(path + '/Viewer/export/' + chamberbox.get() + ' export.xlsx', index=False)                                                                                         #선택된 데이터 excel로 저장
            messagebox.showinfo('Complete', 'export 폴더에 [' + chamberbox.get() + ' export.xlsx] 파일을 생성했습니다.')                                                                    #저장 후 폴더명, 파일이름 알림
        except:
            messagebox.showwarning('Error', '프로그램이 있는 폴더에 export 폴더를 생성해주세요.')                                                                                           #저장 폴더 명이 없을 경우 에러



#그래프 출력
def graph():
    #print(sdf)
    start = check()                                                                                                                                                                     #check함수 실행
    if start == 1:
        plt_title = chamberbox.get() + ' ' + str(sdf.iloc[0,0]) + ' ~ ' + str(sdf.iloc[-1,0])
        plt.figure(figsize=(15, 8))
        plt.title(plt_title)

        if len(df.columns) == 2:
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 1], color='deepskyblue', label=sdf.columns[1])
            plt.legend(ncols=1, loc=(1,0.5), fontsize=9)
            plt.show()
        elif len(df.columns) == 3:
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 1], color='deepskyblue', label=sdf.columns[1])
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 2], color='salmon', label=sdf.columns[2])
            plt.legend(ncols=1, loc=(1,0.5), fontsize=9)
            plt.show()
        elif len(df.columns) == 4:
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 1], color='deepskyblue', label=sdf.columns[1])
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 2], color='salmon', label=sdf.columns[2])
            plt.plot(sdf.iloc[:, 0], sdf.iloc[:, 3], color='limegreen', label=sdf.columns[3])
            plt.legend(ncols=1, loc=(1,0.5), fontsize=9)
            plt.show()
    
    
                                                                                                                                                     



# 창 생성
window = tkinter.Tk()

window.title("REL_Chamber_Profile")
window.geometry("600x290+200+200")
window.resizable(False, False)

titlefont = tkinter.font.Font(family="맑은 고딕",size=15)
menufont = tkinter.font.Font(family="맑은 고딕",size=12)
minifont = tkinter.font.Font(family="맑은 고딕",size=10)
periodfont = tkinter.font.Font(family="맑은 고딕",size=10, weight='bold')
alramfont = tkinter.font.Font(family="맑은 고딕",size=8, weight='bold')


values1=[str(i) for i in list_n]
values2=[]
for i in range(1,32):
    if len(str(i)) == 1:
        values2.append('0'+str(i))
    else:
        values2.append(str(i))
values3=[]
for i in range(0,24):
    if len(str(i)) == 1:
        values3.append('0'+str(i))
    else:
        values3.append(str(i))
values4=[]
for i in range(0,51,10):
    if str(i) == '0':
        values4.append('00')
    else:
        values4.append(str(i))
values5 = [1,5,10,15,20,30,60]

ystext = StringVar()
mstext = StringVar()
dstext = StringVar()
hstext = StringVar()
bstext = StringVar()
yetext = StringVar()
metext = StringVar()
detext = StringVar()
hetext = StringVar()
betext = StringVar()






#창구성
label = tkinter.Label(window, text=' REL Chamber Profile Viewer', width=30, anchor='w', font = titlefont)
label.place(x=10, y=10)
label = tkinter.Label(window, text='Chamber : ', width=10, anchor='w', font = menufont)
label.place(x=10, y=60)
chamberbox = tkinter.ttk.Combobox(window, height=20, values=values1, font=menufont)
chamberbox.place(x=100, y=60)
chamberbox.bind("<<ComboboxSelected>>", fileopen)
slabel_v = StringVar(window)
slabel = tkinter.Label(window, textvariable = slabel_v, width=50, anchor='w', font = periodfont)
slabel.place(x=30, y=100)
label = tkinter.Label(window, text='Start : ', width=7, anchor='center', font = menufont)
label.place(x=10, y=160)
label = tkinter.Label(window, text='End : ', width=7, anchor='center', font = menufont)
label.place(x=10, y=220)
label = tkinter.Label(window, text='End : ', width=7, anchor='center', font = menufont)
label.place(x=10, y=220)
label = tkinter.Label(window, text='개선 사항 건의 - kihun717@lginnotek.com', width=50, anchor='e', font = alramfont)
label.place(x=240, y=270)
label = tkinter.Label(window, text='Ver 221212', width=50, anchor='w', font = alramfont)
label.place(x=10, y=270)

#조건 이름
label = tkinter.Label(window, text='Year', width=7, anchor='center', font = menufont)
label.place(x=77, y=130)
label = tkinter.Label(window, text='Month', width=7, anchor='center', font = menufont)
label.place(x=157, y=130)
label = tkinter.Label(window, text='Day', width=7, anchor='center', font = menufont)
label.place(x=223, y=130)
label = tkinter.Label(window, text='hour', width=7, anchor='center', font = menufont)
label.place(x=308, y=130)
label = tkinter.Label(window, text='min', width=7, anchor='center', font = menufont)
label.place(x=369, y=130)



#조건 콤보박스(Start)
yearsbox = tkinter.ttk.Combobox(window, width=4, font = menufont, textvariable=ystext)
yearsbox.place(x=82, y=160)
monthsbox = tkinter.ttk.Combobox(window, width=3, font = menufont, textvariable=mstext)
monthsbox.place(x=167, y=160)
label = tkinter.Label(window, text='/', width=3, anchor='w', font = menufont)
label.place(x=218, y=160)
sdaybox = tkinter.ttk.Combobox(window, width=3, height=20, values=values2, font=menufont, textvariable=dstext)
sdaybox.place(x=234, y=160)
stimebox = tkinter.ttk.Combobox(window, width=3, height=20, values=values3, font=menufont, textvariable=hstext)
stimebox.place(x=316, y=160)
label = tkinter.Label(window, text=':', width=3, anchor='w', font = menufont)
label.place(x=368, y=160)
sminbox = tkinter.ttk.Combobox(window, width=3, height=20, values=values4, font=menufont, textvariable=bstext)
sminbox.place(x=380, y=160)


#조건 콤보박스(End)
yearebox = tkinter.ttk.Combobox(window,  width=4, font = menufont, textvariable=yetext)
yearebox.place(x=82, y=220)
monthebox = tkinter.ttk.Combobox(window, width=3, font = menufont, textvariable=metext)
monthebox.place(x=167, y=220)
label = tkinter.Label(window, text='/', width=3, anchor='w', font = menufont)
label.place(x=218, y=220)
edaybox = tkinter.ttk.Combobox(window, width=3, height=20, values=values2, font=menufont, textvariable=detext)
edaybox.place(x=234, y=220)
etimebox = tkinter.ttk.Combobox(window, width=3, height=20, values=values3, font=menufont, textvariable=hetext)
etimebox.place(x=316, y=220)
label = tkinter.Label(window, text=':', width=3, anchor='w', font = menufont)
label.place(x=368, y=220)
eminbox = tkinter.ttk.Combobox(window, width=3, height=20, values=values4, font=menufont, textvariable=betext)
eminbox.place(x=380, y=220)

#실행버튼
#label = tkinter.Label(window, text='그래프 단위(min)', width=20, anchor='w', font = alramfont, wraplength=80)
#label.place(x=458, y=105)
#graphbox = tkinter.ttk.Combobox(window,  width=4, font = minifont, values=values5)
#graphbox.set(10)
#graphbox.place(x=523, y=109)
graph_button = tkinter.Button(window, text = '그래프 보기', overrelief='solid', width=15, height=3, repeatdelay=1000, repeatinterval=100, command=graph)
graph_button.place(x=460, y=140)
export_button = tkinter.Button(window, text = '엑셀 추출', overrelief='solid', width=15, height=3, repeatdelay=1000, repeatinterval=100, command=export)
export_button.place(x=460, y=210)


window.mainloop()
