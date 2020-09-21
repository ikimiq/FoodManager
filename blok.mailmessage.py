import smtplib
from email.mime.text import MIMEText
from email.header import Header

# Настройки
mail_sender = 'borovboris43@gmail.com'
mail_receiver = 'sergey_yelizarov@inbox.ru'
username = 'borovboris43@gmail.com'
password = 'qazxc741'
server = smtplib.SMTP('smtp.gmail.com:587')

# Формируем тело письма
subject = u'Тестовый email от ' + mail_sender
body = u'Это тестовое письмо оптправлено с помощью smtplib'
msg = MIMEText(body, 'plain', 'utf-8')
msg['Subject'] = Header(subject, 'utf-8')

# Отпавляем письмо
server.starttls()
server.ehlo()
server.login(username, password)
server.sendmail(mail_sender, mail_receiver, msg.as_string())
server.quit()

#_______________
# Ссылки на инструкции:
# https://defpython.ru/otpravka_poczty_s_pomosczu_smtplib - простая настройка
# http://codius.ru/articles/Python_%D0%9A%D0%B0%D0%BA_%D0%BE%D1%82%D0%BF%D1%80%D0%B0%D0%B2%D0%B8%D1%82%D1%8C_%D0%BF%D0%B8%D1%81%D1%8C%D0%BC%D0%BE_%D0%BD%D0%B0_%D1%8D%D0%BB%D0%B5%D0%BA%D1%82%D1%80%D0%BE%D0%BD%D0%BD%D1%83%D1%8E_%D0%BF%D0%BE%D1%87%D1%82%D1%83 - сложно и не понятно
# http://python-3.ru/page/imap-email-python - принимаем и анализируем письма
