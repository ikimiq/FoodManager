import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

wb = openpyxl.load_workbook(filename='food.xlsx')
wb.active = 0
sheet = wb.active
wb_output = openpyxl.Workbook()
wb_output.active = 0
sheet_output = wb_output.active
sheet_output['A1'] = 'Заказ, с учетом остатков'
for i in range(2, 6):
	sheet_output['A' + str(i)] = sheet['B'+str(i)].value - sheet['C'+str(i)].value


msg = MIMEMultipart()
message = "Thank you"
password = "kingking2281337"
msg['From'] = "testtesttesttest9990@gmail.com"
msg['To'] = "nekrylov1618@gmail.com"
msg['Subject'] = "privet"

msg.attach(MIMEText(message, 'plain'))

server = smtplib.SMTP('smtp.gmail.com: 587')
 
server.starttls()

server.login(msg['From'], password)

server.sendmail(msg['From'], msg['To'], msg.as_string())
 
server.quit()
 
print ("successfully sent email to %s:" % (msg['To']))

wb.save('food.xlsx')
wb_output.save('food1.xlsx')
