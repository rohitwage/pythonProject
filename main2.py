import pandas as pd
import smtplib
from email.message import EmailMessage

df = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv')

print(df.tail(5))
new_df = df.tail(10)
print(df.count())

msg = EmailMessage()
msg['Subject'] = 'Status | Jmeter Validation'
msg['From']='rohit.wage@tietoevry.com'
msg['To'] = 'rohit.wage@tietoevry.com'

body = """\
<html>
  <head> <style>
  .p1 {
  font-family: "sans-serif";
    font-size: 15px;
    }</style></head>
  <body>
 
    <p class="p1">Hi!<br>
    <br></br>
       Please find the attached report of Jenkins job - ServiceNow Jmeter as below..<br>
       <br>Latest top 10 status:
       """+new_df.to_html()+"""
       <font color ="blue">
       <br> Thanks! </br>
       <br></br>
    </p>
    
  </body>
</html>
"""

with open('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv') as myfile:
    data = myfile.read()
    filename=myfile.name
    msg.set_content('------------------------------------')
    msg.add_alternative(body, subtype='html')
    msg.add_attachment(data.encode(encoding="ascii"), maintype="application", subtype="csv", filename=filename)
with smtplib.SMTP_SSL('smtp.office365.com', 587) as server:
    server=smtplib.SMTP_SSL('smtp.office365.com', 587)
    server.login("rohit.wage@tietoevry.com", "06September@2021")
    server.send_message(msg)

print("email sent")