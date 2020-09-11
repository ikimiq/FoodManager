import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

wb = openpyxl.load_workbook(filename='food.xlsx') #открытие файла для чтения
wb.active = 0 #установка рабочего листа на первый
sheet = wb.active
wb_output = openpyxl.Workbook() #создание и открытие файла для записи
wb_output.active = 0 #установка рабочего листа на первый
sheet_output = wb_output.active
sheet_output['A1'] = 'Заказ, с учетом остатков' # на первый столбец таблицы записывается эта фраза
for i in range(2, 6):
	sheet_output['A' + str(i)] = sheet['B'+str(i)].value - sheet['C'+str(i)].value #заносится в столбец A разница между заказом и остатком


msg = MIMEMultipart()
message = "Thank you" #сообщение которое отправялется
password = "kingking2281337" # пароль от почты с которой идет отправление
msg['From'] = "testtesttesttest9990@gmail.com" #почта с которой отправляется
msg['To'] = "sudamessage@gmail.com" # почта куда отправляется
msg['Subject'] = "privet" # тема сообщения

msg.attach(MIMEText(message, 'plain'))

server = smtplib.SMTP('smtp.gmail.com: 587') # подключение к серверу отправки сообщений

server.starttls() # хз точно, но как понял делает подключение к почте безопасным

server.login(msg['From'], password) #логин в почту с которой будет отправляться сообщение

server.sendmail(msg['From'], msg['To'], msg.as_string()) # отправка сообщения

server.quit()

print ("successfully sent email to %s:" % (msg['To'])) # вывод того что сообщение отправилось

wb.save('food.xlsx') #закрытие файла из которго читали
wb_output.save('food1.xlsx') # закрытие файла в который запсиывали

