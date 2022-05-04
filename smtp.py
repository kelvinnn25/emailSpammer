import smtplib, ssl
import pandas as pd

#assumes it is located in first sheet of your worbook indexed 
df = pd.read_excel('emails.xlsx', sheetname=0) 
emailList = df['A'].tolist()

#message list from other excel file, also assumes it is located in first sheet 
df2 = pd.read_excel('messages.xlsx', sheetname=0)
messageList = df2['A'].tolist()

#using SMTP 
server = smtplib.SMTP_SSL("smtp.gmail.com", 587)
#log in with your email and password NEVER LEAVE IT IN THE CODE OR ELSE YOU WILL GET HACKED
server.login('email', 'password')

#from = your email address, 
for reciever in emailList:
    for msg in messageList: 
        server.sendmail("from", reciever, msg)
    server.quit()


