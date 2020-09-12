import openpyxl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage                   
from email.mime.audio import MIMEAudio 
import smtplib


path = r'C:\Users\nekry\Desktop\PYTHON PROJECTS\Foodmanager' #в какой папке искать файл с отчетом
files = os.listdir(path) #список всех файлов в папке
for file in files: #цикл оставляет в списке файлов только файлы с расширениеv .xlsx
	if not (file.endswith(".xlsx") and os.path.isfile(file)):
		files.remove(file)

new = max(files, key=os.path.getctime)#самый новый файл
files = [os.path.join(path, file) for file in files] #добавление полного пути к файлу
new_path = max(files, key=os.path.getctime) #самый новый файл

wb = openpyxl.load_workbook(filename=new) #открытие файла для чтения
wb.active = 0 #установка рабочего листа на первый
sheet = wb.active
wb_output = openpyxl.Workbook() #создание и открытие файла для записи
wb_output.active = 0 #установка рабочего листа на первый
sheet_output = wb_output.active
sheet_output['A1'] = 'Заказ, с учетом остатков' # на первый столбец таблицы записывается эта фраза
for i in range(2, 6):
	sheet_output['A' + str(i)] = sheet['B'+str(i)].value - sheet['C'+str(i)].value #заносится в столбец A разница между заказом и остатком

wb.save(new) #закрытие файла из которго читали
index = new.index('.xlsx')
new_output = new[:index] + '_output.xlsx' #название файла с результатом
wb_output.save(new_output) # закрытие файла в который запсиывали

msg = MIMEMultipart() #создание объекта сообщения
message = "заказ:" #сообщение которое отправялется
password = "kingking2281337" # пароль от почты с которой идет отправление
msg['From'] = "testtesttesttest9990@gmail.com" #почта с которой отправляется
msg['To'] = "sudamessage@gmail.com" # почта куда отправляется
msg['Subject'] = "Поставки" # тема сообщения

msg.attach(MIMEText(message, 'plain'))

server = smtplib.SMTP('smtp.gmail.com: 587') # подключение к серверу отправки сообщений

server.starttls() # хз точно, но как понял делает подключение к почте безопасным

server.login(msg['From'], password) #логин в почту с которой будет отправляться сообщение

filepath = new_output         # если файл в той же папке то не менять если в другой то прописыаать весь путь
filename = os.path.basename(filepath)                     # Только имя файла

if os.path.isfile(filepath):
	ctype, encoding = mimetypes.guess_type(filepath)
	maintype, subtype = ctype.split('/', 1)
	with open(filepath, 'rb') as fp:
		file = MIMEBase(maintype, subtype)
		file.set_payload(fp.read())
		fp.close()
	encoders.encode_base64(file)
	file.add_header('Content-Disposition', 'attachment', filename=filename)
	msg.attach(file)    
server.sendmail(msg['From'], msg['To'], msg.as_string()) # отправка сообщения
server.quit()
print ("successfully sent email to %s:" % (msg['To'])) # вывод того что сообщение отправилось
