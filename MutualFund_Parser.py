import requests
from bs4 import BeautifulSoup
import pandas
from openpyxl import load_workbook
import webbrowser
from googlesearch import search
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

open('MutualFund_Parser_Results.txt', 'w').close()
listURL = []

df = pandas.read_excel('MutualFund_Parser_Data.xlsx', sheet_name=0)
names = list(df.iloc[1:36, 0])

remove_list = ['/returns']
print(names)
for name in names:
    for j in search(name+' moneycontrol NAV', tld='com', lang='en', num=10, start=0, stop=1):
        URL = re.sub('/return', "", j)
        listURL.append(URL)
print(listURL)


def check_price():
    for website in listURL:
        page = requests.get(website)
        soup1 = BeautifulSoup(page.content, 'html.parser')
        soup2 = BeautifulSoup(soup1.prettify(), 'html.parser')
        title = soup2.find(attrs={'class': 'page_heading'}).get_text()
        price = soup2.find(attrs={'class': 'amt'}).get_text()
        print(title.strip() + " " + price.strip())
        with open('mutualdata.txt', 'a', encoding="utf-8") as file:
            file.write(title.strip() + " " + price.strip()+"\n")
            file.close()
    print("Updated")


check_price()


def send_mail():

    mail_content = '''Hello,
    This is an e-mail with the latest mutual fund price data. Please find below the relevant attachments. 
    Thank You
    '''
    # The mail addresses and password
    sender_address = 'SENDER ADDRESS'
    sender_pass = 'SENDER PASSWORD'
    receiver_address = 'RECEIVER ADDRESS'

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Mutual fund price data'

    # The subject line
    # The body and the attachments for the mail
    message.attach(MIMEText(mail_content, 'plain'))

    attach_file_name = 'MutualFund_Parser_Results.txt'

    attach_file = open(attach_file_name, 'rb')  # Open the file as binary mode
    payload = MIMEBase('application', 'octate-stream')
    payload.set_payload((attach_file).read())

    encoders.encode_base64(payload)  # encode the attachment

    # add payload header with filename
    payload.add_header('Content-Disposition',
                       "attachment; filename= %s" % attach_file_name)
    message.attach(payload)
    text = message.as_string()

    # Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
    session.starttls()  # enable security
    # login with mail_id and password
    session.login(sender_address, sender_pass)

    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent')


send_mail()
