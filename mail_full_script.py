# -*- coding: utf-8 -*-
"""
Created on Wed Mar 18 09:52:20 2020

@author: josea.luna
"""
import pandas as pd
import datetime
import unidecode
import win32com.client as win32 

df = pd.read_csv('users.csv') #Nombre del archivo 

name = df['name'].tolist()
jdate = df['Last seen in Jira Software'].tolist()
cdate = df['Last seen in Confluence'].tolist()
em = df['email'].tolist()
cr = df['created'].tolist()
ji = df['Jira Software'].tolist()
con = df['Confluence'].tolist()
gname = df['groupname'].tolist()

maxim = len(df.index)#Cantidad de filas

mailistj = []
mailistjn = []
namelistj=[]
namelistjn=[]
mailistc = []
mailistcn = []
namelistc=[]
namelistcn=[]

for i in range (0,maxim):
        groupn= gname[i]
        if groupn == 'DevOps (billing)':
            newl = jdate[i]
            n = name[i]
            email = em[i]
            created = cr[i]
            jira = ji[i]
            if jira == 'Yes':  
                if newl == 'Never logged in':
                    crea = datetime.datetime.strptime(created, '%d-%b-%y')
                    today = datetime.datetime.now() #Se toma el dia de hoy
                    tn_days = (today - crea).days #Se resta el dia de hoy con el del campo en el excel
                    if tn_days >= 60:
                        if n=="Vanessa Lopez":
                            print("No enviar correo")
                        elif n=="Gabriel Fuentes":
                            print("No enviar correo")
                        elif n=="Jacqueline Romero":
                            print("No enviar correo")
                        else:
                            mailistjn.append(email)
                            namelistjn.append(n)
                else:
                    jdatef = datetime.datetime.strptime(newl, '%d-%b-%y')
                    today = datetime.datetime.now() #Se toma el dia de hoy 
                    t_days = (today - jdatef).days #Se resta el dia de hoy con el del campo en el excel

                    if t_days >= 60:
                        if n=="Vanessa Lopez":
                            print("No enviar correo")
                        elif n=="Gabriel Fuentes":
                            print("No enviar correo")
                        elif n=="Jacqueline Romero":
                            print("No enviar correo")
                        else:
                            mailistj.append(email)
                            namelistj.append(n)

for i in range (0,maxim):
        groupn = gname[i]
        if groupn == 'DevOps (billing)':
            newl = cdate[i]
            n = name[i]
            email = em[i]
            created = cr[i]
            conflu = con[i]
            if conflu == 'Yes': 
                if newl == 'Never logged in':
                    crea = datetime.datetime.strptime(created, '%d-%b-%y')
                    today = datetime.datetime.now()
                    tn_days = (today - crea).days 

                    if tn_days >= 60:
                        if n=="Vanessa Lopez":
                            print("No enviar correo")
                        elif n=="Gabriel Fuentes":
                            print("No enviar correo")
                        elif n=="Jacqueline Romero":
                            print("No enviar correo")
                        else:
                            mailistcn.append(email)
                            namelistcn.append(n)
                else:
                    cdatef = datetime.datetime.strptime(newl, '%d-%b-%y')
                    today = datetime.datetime.now()
                    t_days = (today - cdatef).days 
                    if t_days >= 60:
                        if n=="Vanessa Lopez":
                            print("No enviar correo")
                        elif n=="Gabriel Fuentes":
                            print("No enviar correo")
                        elif n=="Jacqueline Romero":
                            print("No enviar correo")
                        else:
                            mailistc.append(email)
                            namelistc.append(n)

listc = mailistc+mailistcn
listj = mailistj+mailistjn

list1 = []
list3 = []
for i in listj:
    if i not in listc:
        list1.append(i)                      

list2 = []
for i in listc:
    if i not in listj:
        list2.append(i)
    else:
        if i not in list3:
            list3.append(i)
print(list1)
print(list2)
print(list3)
nlistc = namelistc+namelistcn
nlistj = namelistj+namelistjn

nlist1 = []
nlist3 = []
for i in nlistj:
    if i not in nlistc:
        nlist1.append(i)                

nlist2 = []
for i in nlistc:
    if i not in nlistj:
        nlist2.append(i)
    else:
        if i not in nlist3:
            nlist3.append(i)       

#///////////////////////////////////Se genera el archivo de texto
# 

hoy = today.strftime('%B') 
f = open('email_list_'+hoy+'.txt','+w')
f.write('Jira users: '+ '\n')
for a, b in zip(nlist1,list1):   
    f.write(a + ' - ' + b + '\n')
f.write('\n')   
f.write('Confluence users: ' + '\n')
for a, b in zip(nlist2,list2):   
    f.write(a + ' - ' + b + '\n')
f.write('\n')
f.write('Jira & Confluence users: ' + '\n')
for a, b in zip(nlist3,list3):   
    f.write(a + ' - ' + b + '\n')
f.write('\n')
f.close()


#//////////////////////////Se generan los emails para cada usuario y se envían

mailsj = []
if len(list1)!=0:
    for j in range(0,len(list1)):
        strj=list1[j]
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = strj
        #mail.CC = 'stefany.hernandez@softtek.com; jacqueline.romero@softtek.com'
        mail.Subject = 'Acceso Jira DevOps COE'
        mail.HTMLBody = '<p>Buen día</p> <p>El motivo de este correo es para informar que tu cuenta de Jira en el COE de DevOps será desactivada debido a inactividad. Para reactivar tu cuenta favor de enviarme un correo con copia a Jacqueline Romero.</p> <p>Quedo atento a tus comentarios.</p> <p>Saludos<p/>' #this field is optional
        mail.Send()

mailsc=[]
if len(list2)!=0:
    for j in range(0,len(list2)):
        strc=list2[j]
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = strc
        #mail.CC = 'stefany.hernandez@softtek.com; jacqueline.romero@softtek.com'
        mail.Subject = 'Acceso Jira DevOps COE'
        mail.HTMLBody = '<p>Buen día</p> <p>El motivo de este correo es para informar que tu cuenta de Confluence en el COE de DevOps será desactivada debido a inactividad. Para reactivar tu cuenta favor de enviarme un correo con copia a Jacqueline Romero.</p> <p>Quedo atento a tus comentarios.</p> <p>Saludos<p/>' #this field is optional
        mail.Send()

mailsjc=[]
if len(list3)!=0:
    for j in range(0,len(list3)):
        strjc=list3[j]
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = strjc
        #mail.CC = 'stefany.hernandez@softtek.com; jacqueline.romero@softtek.com'
        mail.Subject = 'Acceso Jira DevOps COE'
        mail.HTMLBody = '<p>Buen día</p> <p>El motivo de este correo es para informar que tu cuenta de Jira y Confluence en el COE de DevOps será desactivada debido a inactividad. Para reactivar tu cuenta favor de enviarme un correo con copia a Jacqueline Romero.</p> <p>Quedo atento a tus comentarios.</p> <p>Saludos<p/>' #this field is optional
        mail.Send()

#//////////////////////////////////////Email correspondiente a Jacki y Stef
with open('email_list_'+hoy+'.txt', 'r') as myfile:
    data=myfile.read()
    outlook = win32.Dispatch('outlook.application')        
    mail = outlook.CreateItem(0)
    mail.To = 'josea.luna@softtek.com' #'stefany.hernandez@softtek.com; jacqueline.romero@softtek.com'Correos de Stef y Jackie
    mail.Subject = 'Accesos correspondientes a'
    mail.body = 'La lista de usuarios se da acontinuación: \n' + data      
    mail.Send()        
