import win32com.client
import datetime
import calendar
import re
import os

def main():
    x = get_number()
    make_mail(x)
    print("メールを作成しました。")
    
##########関数作成#################################################
#コンボボックスの選択値をゲット
def get_number():
    while True:
        print("メニュー")
        with open("メニュー.txt", encoding="utf-8") as f:
            menus = f.readlines() #メニュー読み込み
        for i,j in enumerate(menus):
            print(str(i+1)+"："+j, end = "")
        x = str(input("\n入力："))
        try:
            if x == "bye":
                quit()
            elif int(x)-1 in range(len(menus)):
                return x
                break
            else:
                os.system('cls')
                print("表示されている範囲の数字を入力してください")
        except ValueError:
            os.system('cls')
            print("数字以外を入力しないでください")

def read_to(txt):
    with open(txt, encoding="utf-8") as f:
        to_cc_bcc = f.readlines() #宛先読み込み
    return to_cc_bcc

def read_subject(txt):
    week = ["月","火","水","木","金","土","日"]
    with open(txt, encoding="utf-8") as f:
        subject = f.read() #宛先読み込み
    dt_now = datetime.datetime.now()
    if("YYYY" in subject):
        subject = re.sub(r'YYYY', dt_now.strftime('%Y'), subject)
    if("MM" in subject):
        subject = re.sub(r'MM', dt_now.strftime('%m'), subject)
    if("DD" in subject):
        subject = re.sub(r'DD', dt_now.strftime('%d'), subject)
    if("AAA" in subject):
        subject = re.sub(r'AAA', week[dt_now.weekday()], subject)
    return subject

def read_body(txt):
    with open(txt, encoding="utf-8") as f:
        body = f.read() #宛先読み込み
    return body

#メール作成
def make_mail(x):
    #魔法の言葉で、Outlookを使える状態に
    outlook = win32com.client.Dispatch("Outlook.Application")
    #instance生成(メール)
    mail = outlook.CreateItem(0)
    #メッセージ形式はテキスト＝1、HTML＝2、リッチテキスト＝3
    mail.bodyFormat = 3
    #送信先
    mail_to_cc_bcc = read_to("宛先/"+x+".txt")
    mail.to = mail_to_cc_bcc[1]
    mail.cc = mail_to_cc_bcc[4]
    mail.bcc = mail_to_cc_bcc[7]
    #件名
    mail_subject = read_subject("件名/"+x+".txt")
    mail.subject = mail_subject
    #本文
    mail_body = read_body("本文/"+x+".txt")
    mail.body = mail_body
    #メール確認
    mail.display(True)
    #mail.Send()

###################################################################

if __name__ == "__main__":
    main()