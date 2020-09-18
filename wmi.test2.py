import wmi

c = wmi.WMI()
for item in c.Win32_PhysicalMedia():
    if "PHYSICALDRIVE" in str(item.Tag).upper():
        serialNo = item.SerialNumber
        break
    
#строка для демонстрации вывода
print (serialNo)
input()
