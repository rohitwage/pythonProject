import win32com.client as win32
import pandas as pd

df = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv')
df1 = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv', usecols=["Test"])
df2 = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv', usecols=["Status"])
df3 = pd.read_csv('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv', usecols=["TestCase"])

df1_toList = df1.values.tolist()
df1_toList.reverse()
df3_toList = df3.values.tolist()
df3_toList.reverse()

#This script retrives the count of the last report from the csv
i = 0
for x in df3_toList:
    i += 1
    if x == ['TC01_DiscNow_T01_LaunchURL']:
        break

#This script retrives count of rows till Final summary row comes up
j = 0
for y in df3_toList:
    j += 1
    if y == ['Final summary ']:
        break

print("counter :", i, j)


FinalTestName = df.tail(j)
FinalTestStatus = df.tail(j)
FName = FinalTestName.iloc[0]['Test']
Fstatus = FinalTestStatus.iloc[0]['Status']

print(FName)

#The below piece of code is for to reset the index of dataframe that is coming up in the email report with attached to below HTML
new_df = df.tail(i).reset_index()
print(new_df.reset_index(inplace=True))
new_df.index = new_df.index + 1

#Coverting the dataframe into html form inorder to enclose in the email report
html_df = new_df[['Test', 'Date', 'Time', 'Status']]
html = html_df.to_html()

#print(html)

body = """\
<html>
  <head> <style>
  .p1 {
  font-family: "sans-serif";
    font-size: 15px;
    }
    table {background-color: #dee1ff;
    width: 90%; font-size:12pt; border-style:groove; white; border-collapse:collapse; text-align:left;  
    }
    tr:nth-child(even) {background-color: #f2f2f2;}
    </style>
    </head>
  <body>

    <p class="p1">Hi!<br>
    <br></br>
       Please find the attached report of Jenkins job - Discovery ServiceNow - Jmeter validation and summary as below..
       <br>Latest records:
       <br>""" + html + """</br>
       <font color ="blue">
       <br> Thanks! </br>
       <br>Note: This is an auto generated email please do not reply</br>
    </p>

  </body>
</html>
"""
#
with open('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv') as myfile:
    data = myfile.read()
    filename = myfile.name

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'DLMSSSITSMDiscovery@tieto.com;rohit.wage@tietoevry.com'
mail.Subject = 'Discovery Status | '+FName+': '+Fstatus
mail.Body = '***************'
mail.HTMLBody = body
attachment = 'C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv'
mail.Attachments.Add(attachment)

mail.Send()

#print("email sent")