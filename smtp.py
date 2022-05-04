import smtplib, ssl
import pandas as pd
import numpy as np

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
        msgPointer = 0 
        server.sendmail("from", reciever, msg)
        msgPointer +=1

    server.quit()
df = pd.read_excel('messages.xlsx') #Here we read the excel
df.iloc[11:1001,:1] = np.nan #Here we are selecting from the first column upto the first row only and from the 11th row up to the 1000 and making the values null.
df.to_excel('location_of_new_file.xlsx') #Here we are saving the excel with the modifications we made.


