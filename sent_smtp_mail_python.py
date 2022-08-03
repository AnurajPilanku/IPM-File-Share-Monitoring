#Anuraj Pilanku
#Sent mail from outlook which contains images in mailbody

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import basename

From = 'USSACDev@mmm.com'
To = 'P.Anuraj@cognizant.com'
cc="PraveenKumar.B3@cognizant.com"
bcc="D.ManishBabu@cognizant.com"

msgRoot = MIMEMultipart('related')
msgRoot['Subject'] = 'IPM Fileshare Monitoring'
msgRoot['From'] = From
msgRoot['Cc']=cc
msgRoot['To'] = To
msgRoot['Bcc']=bcc
msgRoot.preamble = '====================================================='
msgAlternative = MIMEMultipart('alternative')
msgRoot.attach(msgAlternative)
msgText = MIMEText('Please find the IPM File share details')
msgAlternative.attach(msgText)
msgText = MIMEText("""<html>
<body style="font-family:Times New Roman">
<br/><img src='cid:image1'<br/>
<br>
<br>
<br /><font face='Times New Roman'><b><i>Hi All, </a></i></b></font><br/>

<br /><font face='Times New Roman'><b><i>Please find the Disk space details </a></i></b></font><br/>

<br/><img src='cid:image2'<br/>

<br /><font face='Times New Roman'><b><i>Regards </a></i></b></font><br/>
<br /><font face='Times New Roman'><b><i>3M Automation Center Team </a></i></b></font><br/>
<br>
<br>
<br/><img src='cid:image3'<br/>
</body>
</html>""", 'html')
msgAlternative.attach(msgText)
fp = open("//acdev01/3M_CAC/IPM_FSM/Mail_elements/head.png", 'rb')
fp2 = open("//acdev01/3M_CAC/IPM_FSM/Mail_elements/new.png", 'rb')
fp3 = open("//acdev01/3M_CAC/IPM_FSM/Mail_elements/footer.png", 'rb')
msgImage = MIMEImage(fp.read())
msgImage1 = MIMEImage(fp2.read())
msgImage2 = MIMEImage(fp3.read())
fp.close()
fp2.close()
fp3.close()
msgImage.add_header('Content-ID', '<image1>')
msgImage1.add_header('Content-ID', '<image2>')
msgImage2.add_header('Content-ID', '<image3>')
msgRoot.attach(msgImage)
msgRoot.attach(msgImage1)
msgRoot.attach(msgImage2)

filepaths=[]
for f in filepaths or ["//acdev01/3M_CAC/SMO_AMA/3m_mailid.xlsx","//acdev01/3M_CAC/SMO_AMA/mail_details.xlsx"]:
    with open(f,"rb") as file:
        part=MIMEApplication(file.read(),Name=basename(f))
        part["Content-Disposition"]='attachment;filename="%s"'%basename(f)
        msgRoot.attach(part)
smtp = smtplib.SMTP()
smtp.connect('mailserv.mmm.com')
# smtp.login('username', 'password')
smtp.sendmail(From,To, msgRoot.as_string())
smtp.quit()  # <img src="cid:image1">
print("Email is sent successfully")