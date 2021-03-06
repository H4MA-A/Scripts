from sendgrid.helpers.mail import *
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import sendgrid
import smtplib, ssl, base64

data = pd.read_excel (r'list.xlsx') 
name_list = data["Name"].tolist() 
mail_list = data["Email"].tolist()
ID_list = data["ID"].tolist()
j=0;
count_success=0
for i in name_list:
    try:
        sg = sendgrid.SendGridAPIClient(api_key="#####REDACTED#####")
        from_email = Email("contact@securinets.com")
        to_email = To(mail_list[j])
        mail = Mail(from_email, to_email)
        mail.dynamic_template_data = {'Name' : i,'ID' : ID_list[j], 'Email' : mail_list[j]}
        mail.template_id = "#####REDACTED#####"
        response = sg.client.mail.send.post(request_body=mail.get())
        print(response.status_code)
        print(response.body)
        print("Succesfully sent to",end=" ")
        print(i,end=" - ")
        print(mail_list[j])
        count_success+=1
        print("Succes Count: ",end='')
        print(str(count_success))
        j+=1
    except Exception:
        print("Error: ",end='')
        print(i,end=" - ")
        print(mail_list[j])
        j+=1
        continue
