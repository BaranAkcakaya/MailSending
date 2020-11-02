# -*- coding: utf-8 -*-
"""
Created on Thu Feb  6 15:35:53 2020

@author: lenovoz
"""
'''
mylist = [[1, "blah", "blah2", "blah3", "blah4", "blah5"], [2, "blah33", "blah44", "blah55", "blah66", "blah77"], [3, "blah111", "blah222", "blah333", "blah444", "blah555"]]
mylistitem = (["""<a href=>""" + str(i) + """</a>""" for i in mylist])
merged = [item for sublist in zip(mylistitem) for item in sublist]
htmlline = '\n'.join(merged)
#print htmlline
html = """\
        <html>
          <head></head>
          <body>
            <p><br>
                Python Script Errors & Logging Reports<br> <br> <br>
                | Log Items | <br>
                """ + htmlline + """<br> <br>
            </p>
          </body>
        </html>
        """
print( html)
'''

import pandas as pd

veriler=pd.read_excel('MailInfo.xlsx')
data=veriler.values
send_list=[]
listNot=[]
listComp=[]
delList=[]
htmlL=[]

for i in range(len(data)):
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
    #body="Merhaba"+"\n"+str(data[i][5])+" yılına ait "+str(data[i][0])+" "+str(data[0][1])+" bedeliniz: "+str(data[i][4])+" TL"+"\n"
    for j in range(i+1,len(data)):
        if(data[i][3]==data[j][3]):
            if(str(data[j][6]) != "nan"):
                listNot.append(data[j][6])
            #body=body+str(data[j][5])+" yılına ait "+str(data[j][0])+" "+str(data[j][1])+" bedeliniz: "+str(data[j][4])+" TL"+"\n"
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
    htmlL.append(html)


from tkinter import *
from tkinter import  messagebox
import  smtplib
import sys
    
def giris_onay1():
    mail_adresi=ment1.get()
    sifre=ment2.get()
    mail=smtplib.SMTP('smtp.gmail.com',587)
    mail.ehlo()
    mail.starttls()
    try:
        mail.login(mail_adresi,sifre)
        messagebox.showinfo("Giriş","Giriş Başarılı.")
    except:
        messagebox.showinfo("Giriş","Giriş Başarısız.")
        print('Hata:',sys.exc_info()[0])
        
def giris_onay2():
    mail_adresi=ment1.get()
    sifre=ment2.get()
    mail=smtplib.SMTP('smtp.live.com',587)
    mail.ehlo()
    mail.starttls()
    try:
        #mail.login(mail_adresi,sifre)
        messagebox.showinfo("Giriş","Giriş Başarılı.")
        for s in range(0,10):
            cikti1=Label(wind)
            cikti1["text"]='Mail gönderildi: '+str(s)
            cikti1.place(x=50,y=350+(s*30))
    except:
        messagebox.showinfo("Giriş","Giriş Başarısız.")
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

    
    
    
    
    
    
    
    