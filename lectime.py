import os
import schedule
import time
import sys
import webbrowser
import pyautogui as pt
import requests
from tkinter.messagebox import showwarning,askyesno,showerror
import subprocess
import tkinter as tk
import datetime
from tkinter import *
from threading import Thread
try:
    import getpass4
except ImportError:
    os.system('python -m pip install getpass4')
from win32api import GetSystemMetrics as sm
from PIL import ImageTk, Image

USER_NAME = getpass.getuser()
h=sm(1)
w=sm(0)
time1 = ''

class HoverButton(tk.Button):
    def __init__(self, master, **kw):
        tk.Button.__init__(self,master=master,**kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self['background'] = self['activebackground']

    def on_leave(self, e):
        self['background'] = self.defaultBackground
        
def tick(clock,day,date,title,event):
    global time1
    if s_time_day =='Saturday' or s_time_day =='Sunday':
        title.config(text='No Upcoming Lectures \n Today')
        event.config(text='')
    # get the current local time from the PC
    time2 = time.strftime('%H:%M')
    day1 = datetime.datetime.today().strftime('%A')
    date1 = datetime.date.today()
    # if time string has changed, update it
    if time2 != time1:
        time1 = time2
        clock.config(text=time2)
        day.config(text=day1)
        date.config(text=date1)
        # calls itself every 200 milliseconds
        # to update the time display as needed
        # could use >200 ms, but display gets jerky
    clock.after(500, lambda:tick(clock,day,date,title,event))


def internet_on():
    try:
        requests.get('http://www.zoom.us',timeout=5)
        return 
    except(requests.ConnectionError,requests.Timeout):
        #showwarning('ZOOM Lectuers','Please check your internet connection')
        return
         
def add_to_startup():
    USER_NAME = getpass.getuser()
    try:
        bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
        path=os.listdir(bat_path)
        path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    except:
        bat_path=r"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        path=os.path.expanduser('~')
        path=path+bat_path
        
    count=1
    for file in os.listdir(path):
        if file=="ZOOM lectures.lnk":
            continue
        elif int(len(os.listdir(path)))==count:
            if askyesno('ZOOM Lectuers','do you want to enable autostart?')==True:
                import sys
                from swinlnk.swinlnk import SWinLnk
                swl = SWinLnk()
                swl.create_lnk(sys.argv[0],path + '\\' + "ZOOM lectures.lnk")
        count=count+1
    
def launchzoom():
    time.sleep(10)
    position1=pt.locateOnScreen("launch.png",confidence=.6)
    pt.moveTo(position1)
    pt.click()

def audioenable():
    time.sleep(10)
    position=pt.locateOnScreen("enableaudio.png",confidence=.6)
    pt.moveTo(position)
    pt.click()
    
def click():
    try:
        launchzoom()
    except OSError:
        pass
    try:
        audioenable()
    except OSError:
        pass
    
def monday_lec10():
    webbrowser.open("https://learn.zoom.us/j/69560728079?pwd=TjhTaXV3dW9qam15bFl0UzlyZXBKUT09")
    click()
    
def monday_lec1():
    webbrowser.open("https://learn.zoom.us/j/64478166940?pwd=NDhEL0l5b3hEK05oZTQzcUZoWW9rZz09")
    click()
    
def monday_lec3():
    webbrowser.open("https://learn.zoom.us/j/62244574317?pwd=YkdCY0RNN21FWnU3TklCTmFzQlhOdz09")
    click()
    
def tuesday_lec3():
    webbrowser.open("https://learn.zoom.us/j/64171334570?pwd=TUlWY0JJeHk0QlhIRGN6d2lITUtnQT09")
    click()
    
def wednesday_lec1():
    webbrowser.open("https://learn.zoom.us/j/66839673147?pwd=ZUtBeFVFL1EwM00xVG1kdVRqY29RQT09")
    click()
    
def wednesday_lec3():
    webbrowser.open("https://learn.zoom.us/j/65371646713?pwd=VW1BbXkzekdIaGVjbjNqOEpjaloydz09")
    click()
    
def thursday_lec8():
    webbrowser.open("https://learn.zoom.us/j/67904481619?pwd=Yy94NjFwOUJ5UWVFaXZ6OVhLNm9mUT09")
    click()
    
def thursday_lec1():
    webbrowser.open("https://learn.zoom.us/j/69186141934?pwd=dzFXa294bFJBRm03Y21OaGxhdlB1UT09")
    click()

def thursday_lec3():
    webbrowser.open("https://learn.zoom.us/j/66176598233?pwd=L2hlMjd5bUxLTDdIOG1ZN2wvMVpIdz09")
    click()
    
def friday_lec8():
    webbrowser.open("https://learn.zoom.us/j/68603941277?pwd=eEcxRWp5OFVUaVFaNHBsbHVZMW01dz09")
    click()
    
def friday_lec10():
    webbrowser.open("https://learn.zoom.us/j/64478166940?pwd=NDhEL0l5b3hEK05oZTQzcUZoWW9rZz09")
    click()
    
def friday_lec1():
    webbrowser.open("https://learn.zoom.us/j/69260551295?pwd=alpQNHAxWDYvTGhsUnpSUFd0clZ3Zz09")
    click()

def friday_lec3():
    webbrowser.open("https://learn.zoom.us/j/69475957342?pwd=SlF2YWVTMno3bldQdEhhNXRrcklkUT09")
    click()
    
def friday_lec5():
    webbrowser.open("https://learn.zoom.us/j/67933801349?pwd=TTJ2WHVLdmhUTEZURDJJM0Z6eTk2Zz09")
    click()
    
def run():
    add_to_startup()
    #Testing------
    #webbrowser.open('https://learn.zoom.us/j/67904481619?pwd=Yy94NjFwOUJ5UWVFaXZ6OVhLNm9mUT09')
    #launchzoom()
    #audioenable()
    #-------------
    schedule.every().monday.at('9:55').do(monday_lec10)
    schedule.every().monday.at('12:55').do(monday_lec1)
    schedule.every().monday.at('14:55').do(monday_lec3)
    schedule.every().tuesday.at('14:55').do(tuesday_lec3)
    schedule.every().wednesday.at('12:55').do(wednesday_lec1)
    schedule.every().wednesday.at('14:55').do(wednesday_lec3)
    schedule.every().thursday.at('7:55').do(thursday_lec8)
    schedule.every().thursday.at('12:55').do(thursday_lec1)
    schedule.every().thursday.at('2:55').do(thursday_lec3)
    schedule.every().friday.at('7:55').do(friday_lec8)
    schedule.every().friday.at('9:55').do(friday_lec10)
    schedule.every().friday.at('12:55').do(friday_lec1)
    schedule.every().friday.at('14:55').do(friday_lec3)
    schedule.every().friday.at('16:55').do(friday_lec5)
    schedule.every().saturday.at('15:14').do(friday_lec5)
    schedule.every(1).minutes.do(internet_on)
    #Thread(target=mainwindow).start()
    lap=1
    while True:
        #print('go')
        if lap==1:
            internet_on()
            lap=0
        schedule.run_pending()
        time.sleep(1)
        
def window():
    window = tk.Tk()
    window.title('Zoom Lectures')
    window.attributes('-topmost', True)
    window.iconbitmap(path+'leczoom.ico')
    ws=window.winfo_screenwidth()
    hs=window.winfo_screenheight()
    window.geometry('%dx%d+%d+%d'% (w*26/100,h*26/100,w*72/100,h*2/100))
    window.resizable(0,0)
    image = Image.open('pic.png')
    #image = Image.open(requests.get('https://www.reviewgeek.com/p/uploads/2019/12/afaf4fd9.png',stream=True).raw)
    #time.sleep(3)
    image = image.resize((round(w*26/100),round(h*26/100)), Image.ANTIALIAS)
    photo = ImageTk.PhotoImage(image)
    pic = Label(window,image=photo)
    pic.pack()
    clock = Label(pic,font=('helvetica', 30),bg='#e7e9eb')
    clock.place(x=w*17/100,y=h*0.5/100)
    day = Label(pic,text=datetime.datetime.today().strftime('%A'),font=('helvetica', 10),bg='#e7e9eb')
    day.place(x=w*16/100,y=h*6/100)
    date = Label(pic,text=datetime.date.today(),font=('helvetica', 10),bg='#e7e9eb')
    date.place(x=w*20/100,y=h*6/100)
    title=Label(pic,text='Upcoming Lectures',font=('helvetica', 12,'bold'),bg='#e7e9eb')
    title.place(x=w*5/100,y=h*11/100)
    event=Label(pic,text=DIS_LABEL,font=('helvetica', 10),bg='#e7e9eb',justify=LEFT)
    event.place(x=w*8/100,y=h*14/100)
    tick(clock,day,date,title,event)
    #refresh=HoverButton(window, text = ":", font = "Bahnschrift", background='#e7e9eb', activebackground = "white", relief='flat',  width = "3", height = "1", highlightthickness=0, bd=0, command =window)
    #refresh.place(x=450, y=10)
    #window.after(3000,refresh.invoke())
    window.mainloop()

subj_table={'Monday':[['BUS1340',2],['BCC1340',2],['DSC1340',4]],'Tuesday':[['COM1340',2]],'Wednesday':[['BUS1340',2],['DSC1340',2]],'Thursday':[['PUB1240',2],['BUS1340',2],['ENGLISH',2]],'Friday':[['PUB1240',2],['BCC1340',2],['COM1240',2],['ICT1340',2],['ICT13400',2]],'Saturday':[],'Sunday':[]}
time_table={'Monday':[10,13,15],'Tuesday':[15],'Wednesday':[13,15],'Thursday':[8,13,15],'Friday':[8,10,13,15,17],'Saturday':[],'Sunday':[]}
s_time=datetime.datetime.now()
s_time_date = "{:%H.%M}".format(s_time)
s_time_day = datetime.datetime.today().strftime('%A')
#print(s_time_date, s_time_day)

time1=0
  
for time2 in time_table[s_time_day]:
    if len(time_table[s_time_day])==1:
        time1=(time_table[s_time_day])[0]
        time2=time1+0.30
        
    if time1==0:
        time1=time2
        continue
    if float(s_time_date)>float(time1) and float(s_time_date)<float(time2):
        definition=s_time_day+'_lec'+str(time1)
        print(definition)
        break
    time1=time2
    
DIS_LABEL=''
for l in (subj_table[s_time_day]):
    if DIS_LABEL=='':
        DIS_LABEL=l[0]
        continue
    DIS_LABEL=DIS_LABEL+'\n'+l[0]
    
Thread(target=window).start()
time.sleep(3)

if s_time_day =='Saturday' or s_time_day =='Sunday':
    #print('There are no lectures today')
    sys.exit()

Thread(target=run).start()
webbrowser.open("http://lms.mgt.sjp.ac.lk/my")
subprocess.Popen(r'C:\\Users\\%s\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'% USER_NAME,shell=True)

