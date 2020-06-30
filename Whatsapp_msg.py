# -*- coding: utf-8 -*-
"""
Created on Sun Apr 26 06:50:07 2020

@author: AYUSH CHATPALLIWAR
"""

# -*- coding: utf-8 -*-


from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import urllib.parse
from tkinter import *
from tkinter import scrolledtext,messagebox
import tkinter.font
from tkinter import filedialog
import pandas as pd

driver = None
Link = "https://web.whatsapp.com/"
wait = None
scount=0



def whatsapp_login(nums,msg,stbrow,stmsg):
    global wait, driver, Link, scount
    if(nums != '' and msg != ''):
        chrome_options = Options()
        chrome_options.add_argument('--user-data-dir=./User_Data')
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 20)
        print("SCAN YOUR QR CODE FOR WHATSAPP WEB IF DISPLAYED")
        driver.get(Link)
        driver.maximize_window()
        print("QR CODE SCANNED")
        print(msg)
        numlst=nums.split('\n')
        
        #print(numlst)
        if(numlst[-1]==''):
            numlst.pop()
        for x in numlst:
            send_message('91'+x,msg, 1)
            sleep(10)
        driver.close() # Close the Open tab
        driver.quit()
        stbrow.delete(1.0, END)
        stmsg.delete(1.0, END)
        notif=str(scount)+" "+"Messages Sent Sucessfully"
        messagebox.showinfo("Whatsapp Notification", notif)
    else:
        messagebox.showinfo("Whatsapp Notification", "Empty Field!")
    
    
    
def send_message(number,msg,count):
    global scount
    try:
        # Reference : https://faq.whatsapp.com/en/android/26000030/
        print("In send_message_to_unsavaed_contact method")
        params = {'phone': str(number), 'text': str(msg)}
        end = urllib.parse.urlencode(params)
        final_url = Link + 'send?' + end
        print(final_url)
        driver.set_page_load_timeout(2)
        driver.get(final_url)
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, '//div[@title = "Menu"]')))
        print("Page loaded successfully.")
        sleep(5)
        WebDriverWait(driver, 0).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[3]/button'))).click()
        print("Message sent successfully.\n")
        scount=scount+1
    except Exception as e:
        print(number+' is not a valid Whatsapp Number\n')
        
    
    
def browse(stbrow):
    
    global data
    
    import_file_path = filedialog.askopenfilename()
    data = pd.read_excel (import_file_path)
    phlist = data["Phone No"].tolist()
    for i in phlist:
        stbrow.insert(END,i)
        stbrow.insert(END,'\n')

#GUI    
    
win =Tk()
win.configure(background='#25D366')
win.title("Whatsapp")
win.geometry('1000x780+250+0')

lblhead=Label(win,text="Whatsapp Messenger In Python ",font=("Times New Roman",30,'bold'),bg="#25D366",fg='white')
lblhead.place(x=250)
    
FontOfEntryList=tkinter.font.Font(family="TimesNewRoman",size=15)
FontOfDropList=tkinter.font.Font(family="TimesNewRoman",size=10)
    
btnbrow=Button(win,text="Browse File",width=15,font=("Aerial",12,"bold"),command=lambda: browse(stbrow))
btnbrow.place(x=150,y=100)
    
stbrow=scrolledtext.ScrolledText(win,font=FontOfEntryList,width='60',height='5')
stbrow.place(x=150,y=150)

lblmsg=Label(win,text="Enter Message : ",font=('Aerial',25,'bold'),bg='#25D366',fg='white')
lblmsg.place(x=150, y=300)

stmsg=scrolledtext.ScrolledText(win,font=FontOfEntryList,width='60',height='5')
stmsg.place(x=150,y=350)

btnsend=Button(win,text="Send",width=15,font=("Aerial",15,"bold"),command=lambda: whatsapp_login(stbrow.get('1.0', 'end-1c'),stmsg.get('1.0', 'end-1c'),stbrow,stmsg))
btnsend.place(x=400,y=520)

win.mainloop()
