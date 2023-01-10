import smtplib
import os
import pandas as pd
from email.message import EmailMessage

# specify the folder where the files are located
folder_path = "folder path" # write here the path of the folder

#Hoja1 is where the emails are located
data = pd.read_excel('file path', sheet_name='Hoja1') # Choose the file and the excel sheet you store the emails of your distribution list
#Hoja2 is the name of the pptx files
data1 = pd.read_excel('file path', sheet_name='Hoja2') # Same file but with the other sheet with the names of the files
#translating into lists
emails = data.values.tolist()
pptx_file = data1.values.tolist()

email_id = "youremail@gmail.com"
email_pass = "yourpassword"

count = 1
for i in range(len(emails)):
   receiver = emails[i][0]
   msg = EmailMessage()
   msg['Subject'] = 'Resultados pruebas psicológicas - Informe'
   msg['From'] = email_id
   msg['To'] = receiver
   msg.set_content('Estimado/a, \nEl presente correo es para ... \nAtentamente, \nEquipo de Investigación')
   file = os.path.join(folder_path, pptx_file[i][0])
   with open(file,'rb') as f:
      file_data = f.read()
      file_name = f.name
      msg.add_attachment(file_data, maintype='application', subtype = 'vnd.openxmlformats-officedocument.presentationml.presentation', filename=file_name)
   with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
      smtp.login(email_id, email_pass)
      smtp.send_message(msg)
      print(count)
      count = count+1
