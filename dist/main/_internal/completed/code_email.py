from email.mime.image import MIMEImage
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import re

signature_format = """
    <table cellpadding="0" width="400" style="border-collapse: collapse; font-size: 10.7px;">
  <tr>
    <td style="margin: 0.1px; padding: 0px; cursor: pointer;">
      <img src="cid:SignatureImage" width="77" alt=" &quot;created with MySignature.io&quot;" style="display: block; min-width: 77px;">
    </td>
  </tr>
  <tr>
    <td style="margin: 0.1px; padding: 10px 10px 0px 0px; font: 600 13.9px / 17.7px Verdana, Geneva, sans-serif; color: rgb(0, 0, 0); text-transform: uppercase; letter-spacing: 0.4pt; cursor: pointer;">(name)</td>
  </tr>
  <tr>
    <td style="margin: 0.1px; padding: 5px 0px 0px; font: 10.7px / 13.6px Verdana, Geneva, sans-serif; color: rgb(0, 0, 1);">
      <span style="cursor: pointer;">(title)</span>&nbsp;
    </td>
  </tr>
  <tr>
    <td style="margin: 0.1px; padding: 15px 0px;">
      <table cellpadding="0" style="border-collapse: collapse;">
        <tr>
          <td height="1" style="margin: 0.1px; padding: 0px; border-top: 2px solid rgb(0, 0, 1); font: 1px / 1px Verdana, Geneva, sans-serif; width: 30px;"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr style="cursor: pointer;">
    <td style="margin: 0.1px; padding: 1px 0px; font: 10.7px / 13.6px Verdana, Geneva, sans-serif; color: rgb(0, 0, 1);">
      <!---->
      <a href="mailto:(email)" target="_blank" style="color: rgb(0, 0, 1); text-decoration: none; font-family: Verdana, Geneva, sans-serif;">(email)</a>
    </td>
  </tr>
  <tr style="cursor: pointer;">
    <td style="margin: 0.1px; padding: 1px 0px; font: 10.7px / 13.6px Verdana, Geneva, sans-serif; color: rgb(0, 0, 1);">
      <!---->
      <span style="color: rgb(0, 0, 1); text-decoration: none; font-family: Verdana, Geneva, sans-serif;"></span>
    </td>
  </tr>
  <tr style="cursor: pointer;">
    <td style="margin: 0.1px; padding: 1px 0px; font: 10.7px / 13.6px Verdana, Geneva, sans-serif; color: rgb(0, 0, 1);">
      <!---->
      <a href="tel:(num)" target="_blank" style="color: rgb(0, 0, 1); text-decoration: none; font-family: Verdana, Geneva, sans-serif;">(num)</a>
    </td>
  </tr>
    <tr style="cursor: pointer;">
    <td style="margin: 0.1px; padding: 1px 0px; font: 10.7px / 13.6px Verdana, Geneva, sans-serif; color: rgb(0, 0, 1);">
      <!---->
      <a href="(website)" target="_blank" style="color: rgb(0, 0, 1); text-decoration: none; font-family: Verdana, Geneva, sans-serif;">(website)</a>
    </td>
  </tr>

  <!---->
  <!---->
</table>
"""

email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
website_regex = r'\b(?:http[s]?://|www\.)[a-zA-Z0-9.-]+\.[a-zA-Z]{2,7}\b'
phone_regex = r'\+\d{11}'

def Email_send_function(to, subject, message, uname, pasw,attach=None):
    # print(to,subject,message,uname,pasw)

    s = smtplib.SMTP("smtp.gmail.com", 587)  # create session for gmail
    s.starttls()  # transport layer
    s.login(uname, pasw)
    msg = "Subject: {}\n\n{}".format(subject, message)
    s.sendmail(uname, to, msg)
    x = s.ehlo()
    if x[0] == 250:
        return "s"
    else:
        return "f"
    s.close()


def generateHTMLResponse(name,email,title,num,website,img):
    signature = signature_format
    signature = signature.replace("(name)", name)
    signature = signature.replace("(email)", email)
    signature = signature.replace("(title)", title)
    signature = signature.replace("(num)", num)
    signature = signature.replace("(website)", website)
    return signature

def send_email(row, uname, pasw):
    email = row['Email']
    subject = row["Subject"]
    body = re.sub(r"ASIN\s*\(\s*\)", f"ASIN ({row['ASIN']})", row["Email-Content"])


    name_split = row["SignatureName"].split("\n")
    name_split = [i for i in name_split if i != "" and i != " "]
    name = name_split[0]
    title = name_split[1]

    links = row["SignatureLinks"]
    
    num = re.findall(phone_regex, links)[0]
    website = re.findall(website_regex, links)[0]
    email_signature = re.findall(email_regex, links)[0]


    img = row["SignatureImg"]

    signature = generateHTMLResponse(name,email_signature,title,num,website,img)

    with open("signature.html", "w") as f:
        f.write(signature)

    if row["Attachments"] != "nan":
        attach = row["Attachments"]
        file = None
        with open(attach, "rb") as f:
            file = f.read()
        attachment = MIMEApplication(file)
        attachment.add_header('Content-Disposition', 'attachment', filename=attach)
    

    # create the message
    msg = MIMEMultipart()
    msg['From'] = uname
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    msg.attach(MIMEText(signature, 'html'))
    if row["Attachments"] != "nan":
        msg.attach(attachment)

    fp = open(img, 'rb')
    image = MIMEImage(fp.read())
    fp.close()

    image.add_header('Content-ID', '<SignatureImage>')
    msg.attach(image)

    s = smtplib.SMTP("smtp.gmail.com", 587)  
    s.starttls() 
    s.login(uname, pasw)
    s.sendmail(uname, email, msg.as_string())
    x = s.ehlo()
    if x[0] == 250:
        s.close()
        return "s"
    else:
        s.close()
        return "f"
    

    