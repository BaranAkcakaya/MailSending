# -*- coding: utf-8 -*-
"""
Created on Wed Feb  5 16:58:35 2020

@author: lenovoz
"""

import pandas as pd
import  smtplib
import sys


veriler=pd.read_excel('MailInfo.xlsx')
data=veriler.values
send_list=[]
listNot=[]
listComp=[]
delList=[]

from email.mime.multipart import  MIMEMultipart
from email.mime.text import MIMEText
from tkinter import StringVar,Button,Label,Entry,Tk
from tkinter import  messagebox

def mailSend():
    say=0
    for i in range(len(data)):
        message=MIMEMultipart()
        message['From']=''
        message['To']=data[i][3]
        message['Subject']='Proje Tutarı'
        if(data[i][3] in send_list):
            continue
        if(str(data[i][6]) != "nan"):
            listNot.append(data[i][6])
        html='<html><head>Merhaba</head>'
        html=html+'<html><p><b><u>'+str(data[i][2]).upper()+'</u></b></p><body>'
        html=html+'<p>'+str(data[i][5])+' yılına ait '
        html=html+str(data[i][0]).upper()+' '
        html=html+str(data[i][1]).upper()+' bedeliniz: '
        html=html+'<span style="color: red;"><b>'+str(data[i][4])+' TL</b></span>'
        for j in range(i+1,len(data)):
            if(data[i][3]==data[j][3]):
                if(str(data[j][6]) != "nan"):
                    listNot.append(data[j][6])
                if(data[i][2]==data[j][2]):
                    html=html+'<br>'+str(data[j][5])+' yılına ait '
                    html=html+str(data[j][0]).upper()+' '
                    html=html+str(data[j][1]).upper()+' bedeliniz: '
                    html=html+'<span style="color: red;"><b>'+str(data[j][4])+' TL</b></span>'
                else:
                    listComp.append(data[j])
        while(len(listComp)!=0):
            temp=listComp[0][2]
            html=html+'<html><p><b><u>'+temp.upper()+'</u></b></p><body>'
            for t in range(0,len(listComp)):
                if(temp==listComp[t][2]):
                    html=html+'<br>'+str(listComp[t][5])+' yılına ait '
                    html=html+str(listComp[t][0]).upper()+' '
                    html=html+str(listComp[t][1]).upper()+' bedeliniz: '
                    html=html+'<span style="color: red;"><b>'+str(listComp[t][4])+' TL</b></span>'
                else:
                    delList.append(listComp[t])
            if(len(delList)==0):
                listComp.clear()
                delList.clear()
            else:
                listComp.clear()
                for m in delList:
                    listComp.append(m)
                delList.clear()
                
        html=html+'</br></p>'
        if(len(listNot)!=0):
            html=html+'<p>'
            for k in listNot:
                html=html+'<b><u><br>NOT:</u>'+str(k).upper()+'</b>'
            html=html+'</br></p>'
        send_list.append(data[i][3])
        listNot.clear()
        html=html+'<head> İyi Günler</head>'+'</body></html>'
        html_text=MIMEText(html,'html')
        message.attach(html_text)
        try:
            try:
                mail.sendmail(message['From'],message['To'],message.as_string())
            except:
                print("basarısız")
            print('Mail Gönderildi.')
            cikti1=Label(wind)
            cikti1["text"]='Mail gönderildi: '+str(data[i][3])
            cikti1.place(x=50,y=350+(say*30))
            say=say+1
        except:
            print("Mail Gönderilemedi:",data[i][3])
            print('Hata:',sys.exc_info()[0])
            cikti2=Label(wind)
            cikti2["text"]='Mail gönderilemedi: '+str(data[i][3])
            cikti2.place(x=50,y=350+(say*30))
            say=say+1

            
def giris_onay1():
    mail_adresi=ment1.get()
    sifre=ment2.get()
    mail=smtplib.SMTP('smtp.gmail.com',587)
    mail.ehlo()
    mail.starttls()
    try:
        mail.login(mail_adresi,sifre)
        messagebox.showinfo("Giriş","Giriş Başarılı.")
        mailSend()
        try:
            mail.close()
        except:
            print("kapatılamadı")
    except:
        messagebox.showinfo("Giriş","Giriş Başarısız."+"Hata:"+str(sys.exc_info()[0]))
        print('Hata:',sys.exc_info()[0])
        
def giris_onay2():
    global mail
    mail_adresi=ment1.get()
    sifre=ment2.get()
    mail=smtplib.SMTP('smtp.live.com',587)
    mail.ehlo()
    mail.starttls()
    try:
        mail.login(mail_adresi,sifre)
        messagebox.showinfo("Giriş","Giriş Başarılı.")
        mailSend()
        try:
            mail.close()
        except:
            print("kapatılamadı")
    except:
        messagebox.showinfo("Giriş","Giriş Başarısız."+"Hata:"+str(sys.exc_info()[0]))
        print('Hata:',sys.exc_info()[0])


wind=Tk()
ment1=StringVar()
ment2=StringVar()
wind.title("Kullanıcı Giriş Ekrenı")
wind.geometry("800x500")
wind.configure(background="#2a52be")
inp1=Label(wind,text="Kullanıcı Adı",width=20,bg='#ffffff').place(x=50,y=50)
ent1=Entry(wind,textvariable=ment1,width=50).place(x=220,y=50)
inp2=Label(wind,text="Şifre",width=20,bg='#ffffff').place(x=50,y=100)
ent2=Entry(wind,textvariable=ment2,width=50).place(x=220,y=100)
btn3=Button(wind,text="Gmail Giriş",width=20,height=1,bg='#ffffff',font='bold',command=giris_onay1).place(x=125,y=175)
btn4=Button(wind,text="Hotmail Giriş",width=20,height=1,bg='#ffffff',font='bold',command=giris_onay2).place(x=375,y=175)
wind.mainloop()
