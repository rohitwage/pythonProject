import win32com.client as win32

import pandas as pd

df1 = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv', usecols=["TestCase"])
FinalStatus = df1.tail(1)
# print(FinalStatus.iloc[0]['TestCase'])
FStatus = FinalStatus.iloc[0]['TestCase']

df = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv')

print(df.tail(5))
new_df = df.tail(20)
print(df.count())

body = """\
<html>
  <head> <style>
  .p1 {
  font-family: "sans-serif";
    font-size: 15px;
    }</style>
    <style> 
    br, table, th, td {{font-size:10pt; border:1px solid black; border-collapse:collapse; text-align:left;}}
    th, td {{padding: 5px;}}
        </style></head>
  <body>

    <p class="p1">Hi!<br>
    <br></br>
       Please find the attached report of Jenkins job - Discovery ServiceNow - Jmeter validation and summary as below..<br>
       <br>Latest top 10 status:
       """ + new_df.to_html(index=False, col_space=150, justify='right',classes='format.css') + """
       <font color ="blue">
       <br> Thanks! </br>
       <br>Note: This is an auto generated email please do not reply</br>
    </p>

  </body>
</html>
"""

with open('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv') as myfile:
    data = myfile.read()
    filename = myfile.name

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'rohit.wage@tietoevry.com'
mail.Subject = 'Discovery validation Status :'+FStatus
mail.Body = '***************'
mail.HTMLBody = body
attachment = 'C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv'
mail.Attachments.Add(attachment)

mail.Send()

print("email sent")