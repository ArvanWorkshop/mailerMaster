import xlrd
import pandas as pd
import random
import time
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class Excel():
    def __init__(self):
        pass
    def reademail(self, mailPath):
        data = pd.read_excel(mailPath, 'Sheet1')
        df = data.to_dict()
        return df

mailPath = "Email_list.xlsx"
reademail = Excel()
mailList = reademail.reademail(mailPath)

Total_sender_email = len(mailList['SenderEmail'])
Total_Password = len(mailList['Password'])
Total_name = len(mailList['Name'])
Total_ReceiverEmail = len(mailList['ReceiverEmail'])

print("Total Sender Emails: ", Total_sender_email)
print("Total Password: ", Total_sender_email)
print("Total Name: ", Total_sender_email)
print("Total ReceiverEmails: ", Total_sender_email,'\n')
time.sleep(2)
for i in range(Total_sender_email):
    temp = {
        'SenderEmail': mailList['SenderEmail'][i],
        'Password': mailList['Password'][i],
        'Name': mailList['Name'][i],
        'ReceiverEmail': mailList['ReceiverEmail'][i],
    }
    print(mailList['SenderEmail'][i])

        ##############
    message = MIMEMultipart("alternative")
    message["Subject"] = "heey" + " " + mailList['Name'][i]
    message["From"] = "Julia Wiesner"
    message["To"] = mailList['ReceiverEmail'][i]
    ##############

    # Create the plain-text and HTML version of your message
    text = """
    hi ,
    how are you !!
    Welcome to mailerMaster unlimited email sending get way
           """
    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    # part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    # message.attach(part2)
    time.sleep(2)
    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, local_hostname='localhost', timeout=10) as server:
        server.login(mailList['SenderEmail'][i], mailList['Password'][i])
        server.sendmail(
            mailList['SenderEmail'][i], mailList['ReceiverEmail'][i], message.as_string()
        )

    print("send was successful =====>>", "from:", mailList['SenderEmail'][i], "to", mailList['ReceiverEmail'][i], "(", mailList['Name'][i], ")", "====>>","done..")














