from sendgrid.helpers.mail import *
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import sendgrid
import smtplib, ssl, base64

data = pd.read_excel (r'intro.xlsx') 
name_list = data["Name"].tolist() 
mail_list = data["Email"].tolist()
j=0;
count_success=0
for i in name_list:
    try:
        im = Image.open(r'Certificat_Intro.jpg')
        d = ImageDraw.Draw(im)
        text_color = (226, 17, 14)
        font = ImageFont.truetype("Gourmet Hearth df.otf", 250)
        w, h = d.textsize(i, font=font)
        location = ((3508-w)/2, 1150)
        d.text(location, i, fill = text_color, font = font)
        im.save("certificate_" + i +".pdf")
        sg = sendgrid.SendGridAPIClient(api_key="#####REDACTED#####")
        from_email = Email("contact@securinets.com")
        to_email = To(mail_list[j])
        mail = Mail(from_email, to_email)
        mail.dynamic_template_data = {'Name' : i}
        mail.template_id = "#####REDACTED#####"
        #with open("certificate_" + i +".pdf", 'rb') as sigf:
        #    sig = sigf.read()
        sig = open("certificate_" + i +".pdf", "rb").read()
        encoded = base64.b64encode(sig).decode()
        attachment = Attachment()
        attachment.file_content = FileContent(encoded)
        attachment.file_type = FileType('pdf')
        attachment.file_name = FileName("certificate_" + i +".pdf")
        attachment.disposition = Disposition('attachment')
        attachment.content_id = ContentId('Example Content ID')
        mail.attachment = attachment
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
