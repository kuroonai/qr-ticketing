#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 19 12:55:23 2019

@author:Naveen Kumar Vasudevan, 
        400107764,
        Doctoral Student, 
        The Xi Research Group, 
        Department of Chemical Engineering,
        McMaster University, 
        Hamilton, 
        Canada.
         
        naveenovan@gmail.com
        https://naveenovan.wixsite.com/kuroonai
"""

import os
import qrcode
# qr = qrcode.QRCode()
from PIL import Image, ImageDraw, ImageFont
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

loc = os.getcwd()
os.chdir(loc)


scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

'''
steps to get the json file from your google spreadsheet:
    
    1. install gspread (pip install gspread)
    2. For setting up credentials,  go to Google developers console
    https://console.developers.google.com/cloud-resource-manager
    3. Create or select an existing project
    4. select the project and go to navigation menu on the top left corner
    5. On that menu, select API & services and then credentials
    6. Select Create credentials and choose Service account key
    7. In the Create service account key page, select new service account and
        give a new name to service accound and in role dropdown list select project -> owner
    8. Now download the json file and then put it in the currect working folder
    9. In addition to these steps you might need to enable google drive and gsheets to this project
        by going to, API & services and add them via "Enable API and services button"
    10. Also disable secure apps only option for the email you are about to use.
    11. Finally give the email id under "client email" in the json file, access 
        to the spread sheet you are working on.
'''

credentials = ServiceAccountCredentials.from_json_keyfile_name(
         'name of the file.json', scope) # Your json file here

gc = gspread.authorize(credentials)

wks = gc.open("Name of the gsheet").sheet1

data = wks.get_all_values()
headers = data.pop(0)

df = pd.DataFrame(data, columns=headers)

'''
selecting based on the column names, status of ticket is given by sent or pending
'''

guests = df.loc[:, ["First name","Last name","email","ticketed by","status","Encryption","Dietary restriction"]]
guests = guests[guests['Last name'].str.strip().astype(bool)]
guests = guests.loc[guests['status'] == 'pending']


print('\nEmailed the following guest(s):\n')
for guest in guests.index:
    
    #if guests.loc[guest]['status'] == 'pending':
    
    '''
    For generating QR code here a encrytion value is used,
    use which ever values you want to be part of the generation
    Also have a font file in the format of .ttf in the working folder
    I used freebooterscript here.
    '''
    
    qr = qrcode.QRCode()
    img = qrcode.make(str(guest+1)+'\n'+guests.loc[guest]['First name'] +'\n'+ guests.loc[guest]['Last name'] +'\n'+ guests.loc[guest]['Encryption'],\
                          box_size=100)
    qrname = guests.loc[guest]['First name']+'_'+guests.loc[guest]['Last name']+'_'+guests.loc[guest]['Encryption']+'.png'
    #img.save(qrname)
    fontFile = 'freebooterscript.ttf'
    draw = ImageDraw.Draw(img)
    QRwidth, QRheight = img.size
    fontSize = 200#starting font size
    qrLabel = str(guest+1)+' '+guests.loc[guest]['First name']+' '+guests.loc[guest]['Last name']
    img_fraction = 0.90# portion of image width you want text width to be
    fontHeightMax = qr.border * qr.box_size - 10 #Maximum height for captions
    captionX = 0 #Where to start the caption X coordinate
    captionY = 0 #Where to start the caption y coordinate
    font = ImageFont.truetype(fontFile, fontSize)
    while font.getsize(qrLabel)[0] < img_fraction*QRwidth and font.getsize(qrLabel)[1] < fontHeightMax:
        fontSize += 1
        font = ImageFont.truetype(fontFile, fontSize)
    captionX = int(QRwidth - font.getsize(qrLabel)[0]) / 2 #Center the label
    draw.text((captionX, captionY), qrLabel, font=font,fill='black')
    img.save(qrname)
    
    fromaddr = "your email id here"
    frmpwd = 'your password here'
    
    toaddr = guests.loc[guest]['email']

		# instance of MIMEMultipart 
    msg = MIMEMultipart() 
    
    # storing the senders email address   
    msg['From'] = fromaddr 
    
    # storing the receivers email address  
    msg['To'] = toaddr 
    
    # storing the subject  
    msg['Subject'] = "ChEGSS Holiday Banquet 2019 - Ticket"
    
    # string to store the body of the mail 
    body = 'Hi,\n\nThis is the body of your email\n\nWith regards,\nNaveen\nnaveenovan@gmail.com\n'
    
      
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
    
    # open the file to be sent  
    filename = qrname
    attachment = open(loc+qrname, "rb") 
    
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
    
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
    
    # encode into base64 
    encoders.encode_base64(p) 
    
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 
    
    # start TLS for security 
    s.starttls() 
    
    # Authentication 
    s.login(fromaddr, frmpwd) 
    
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
    
    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 
    
    # terminating the session 
    s.quit() 
    
    print(guest+1, guests.loc[guest]['First name'], guests.loc[guest]['Last name'])
    cell = 'f%s'%str(guest+2)
    wks.update_acell(cell, "sent")
    #else:
        #continue
print('\nDone!')


