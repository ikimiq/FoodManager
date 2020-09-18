import wmi
import datetime

# Блок проверки нахождения id в базе пользователей программы.
# Получение id HDD
c = wmi.WMI()
for item in c.Win32_PhysicalMedia():
    if "PHYSICALDRIVE" in str(item.Tag).upper():
        serialNo = item.SerialNumber
        break

#Запрос в БД (Есть ли в ней строка равная переменной serialNo)


#________________________________________

# Запуск этой функции долько в автоматическом режиме (переключатель в екселе)



# Получае текущего дня и месяца
time = str(datetime.datetime.now())
timeX = (time[8:10])

#Нужно прочитать из икселя день отправки сообщения - PlanTimeX

if (timeX == PlanTimeX):
    # основная функция программы и к началу постоянного цикла проверки времени
else:
    time.sleep(30)
    # вернёт программу к началу постоянного цикла проверки времени

        

