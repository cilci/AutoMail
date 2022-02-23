import win32com.client
import tkinter as tk
from tkinter import ttk
import datetime
import calendar
import re
week = ["月","火","水","木","金","土","日"]

#Outlookを使える状態に
outlook = win32com.client.Dispatch("Outlook.Application")


#ポップアップ作成
root = tk.Tk()
root.title("メール作成")
#root.geometry('500x700')

##########関数作成#################################################
#コンボボックスの選択値をゲット
def get_combo(event):
    read_to("宛先/"+str(combobox.current()+1)+".txt")
    read_subject("件名/"+str(combobox.current()+1)+".txt")
    read_body("本文/"+str(combobox.current()+1)+".txt")

def read_to(txt):
    with open(txt, encoding="utf-8") as f:
        to_cc_bcc = f.readlines() #宛先読み込み
    mail_to.delete('1.0', 'end')
    mail_to.insert('1.0',to_cc_bcc[1])
    mail_cc.delete('1.0', 'end')
    mail_cc.insert('1.0',to_cc_bcc[4])
    mail_bcc.delete('1.0', 'end')
    mail_bcc.insert('1.0',to_cc_bcc[7])

def read_subject(txt):
    with open(txt, encoding="utf-8") as f:
        subject = f.read() #宛先読み込み
    dt_now = datetime.datetime.now()
    if("TODAY" in subject):
        subject = re.sub(r'TODAY', dt_now.strftime(' %Y/%m/%d')+"（"+week[dt_now.weekday()]+"）", subject)
    mail_subject.delete('1.0', 'end')
    mail_subject.insert('1.0', subject)

def read_body(txt):
    with open(txt, encoding="utf-8") as f:
        body = f.read() #宛先読み込み
    mail_body.delete('1.0', 'end')
    mail_body.insert('1.0',body)

#メール作成
def make_mail():
    #instance生成(メール)
    mail = outlook.CreateItem(0)
    mail.bodyFormat = 1
    #送信先
    mail.to = mail_to.get('1.0', 'end -1c')
    mail.cc = mail_cc.get('1.0', 'end -1c')
    mail.bcc = mail_bcc.get('1.0', 'end -1c')
    #件名
    mail.subject = mail_subject.get('1.0', 'end -1c')
    #本文
    mail.body = mail_body.get('1.0', 'end -1c')
    #メール確認
    mail.display(True)

#メール送信チェック
def send_check():
    #if send_check_button.current():
    if var.get():
        send_button['state'] = "normal"
    else:
        send_button['state'] = "disabled"

#メール送信
def send_mail():
    #instance生成(メール)
    mail = outlook.CreateItem(0)
    mail.bodyFormat = 1
    #送信先
    mail.to = mail_to.get('1.0', 'end -1c')
    mail.cc = mail_cc.get('1.0', 'end -1c')
    mail.bcc = mail_bcc.get('1.0', 'end -1c')
    #件名
    mail.subject = mail_subject.get('1.0', 'end -1c')
    #本文
    mail.body = mail_body.get('1.0', 'end -1c')
    #メール確認
    mail.Send()
    #root.quit()
###################################################################


style = ttk.Style()
style.theme_use("winnative")
style.configure("office.TCombobox", selectbackground="blue", fieldbackground="red", padding=5)


##########メニュー##########
with open("menu.txt", encoding="utf-8") as f:
        menus = f.readlines() #メニュー読み込み
v = tk.StringVar()
combobox = ttk.Combobox(root, textvariable= v, values = menus, height=1)
combobox.bind('<<ComboboxSelected>>', get_combo)
combobox.set(menus[0])

##########宛先##########
label_to = ttk.Label(root, text="To：")
mail_to = tk.Text(root, height=2)
label_cc = ttk.Label(root, text="Cc：")
mail_cc = tk.Text(root, height=2)
label_bcc = ttk.Label(root, text="Bcc：")
mail_bcc = tk.Text(root, height=2)
read_to("宛先/1.txt")

##########件名##########
label_subject = ttk.Label(root, text="件名：")
mail_subject = tk.Text(root, height=1)
read_subject("件名/1.txt")

##########本文##########
label_body = ttk.Label(root, text="本文：")
mail_body = tk.Text(root, height=30, width=100)
read_body("本文/1.txt")

##########作成ボタン##########
make_button = ttk.Button(root, text="作成", command = make_mail)
##########送信チェックボタン##########
var = tk.BooleanVar()
var.set(False) 
send_check_button = tk.Checkbutton(root, text='そのまま送信する場合はチェック', command=send_check, variable = var)
##########送信ボタン##########
send_button = ttk.Button(root, text="送信", command = send_mail, state=tk.DISABLED)
##########閉じるボタン##########
close_button = ttk.Button(root, text="閉じる", command = root.quit)

##########ウィジェット配置##########
combobox.pack(expand=0, padx=3, pady=3)

label_to.pack(fill="x", expand=0, padx=3, pady=3)
mail_to.pack(fill="x", expand=0, padx=3, pady=3)
label_cc.pack(fill="x", expand=0, padx=3, pady=3)
mail_cc.pack(fill="x", expand=0, padx=3, pady=3)
label_bcc.pack(fill="x", expand=0, padx=3, pady=3)
mail_bcc.pack(fill="x", expand=0, padx=3, pady=3)

label_subject.pack(fill="x", expand=0, padx=3, pady=3)
mail_subject.pack(fill="x", expand=0, padx=3, pady=3)

label_body.pack(fill="x", expand=0, padx=3, pady=3)
mail_body.pack(fill="both", expand=True, padx=3, pady=3)

make_button.pack(expand=0, padx=3, pady=3)
send_check_button.pack(expand=0, padx=3, pady=3)
send_button.pack(expand=0, padx=3, pady=3)
close_button.pack(expand=0, padx=3, pady=3)
root.mainloop()
