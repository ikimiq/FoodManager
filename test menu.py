menuststus = "1"
while (menuststus == "1"):
    menuststus = "0"
    print ("Добро пожаловать в FoodManager!\n\n1.Запустить программу\n2.Настройки автозапуска\n3.Правовая информация")
    per=input()
    
    if (per=="1"):
        #pern="1"       # Демонстрационная строка. В программе быть не должно
        menustatus2 = "1"
        if (pern=="1"): #Должна читаться из екселя
            while (menustatus2 == "1"):
                menustatus2 = "0"
                print("\n1.Запустить программу\n2.Запустить(больше не напоминать)\n3.К настройкам")
                per2=input()
                if (per2=="1"):
                    print(start)
                    # Запуск основной прграммы def или иначе
                elif (per2=="2"):
                    # Нужно изменить pern на 0 в екселе
                     print(start)
                    # Запуск основной прграммы def или иначе
                elif (per2=="3"):
                    print("#Тут будет лежать инструкция по настройте автозапуска")
                else:
                    print("Вы допустили ошибку. Введите повторно.\n")
                    menustatus2 = "1"
        
    elif (per=="2"):
        print("#Тут будет лежать инструкция по настройте автозапуска")
        
    elif (per=="3"):
        print("#Тут будет лежать правовая информация на использование программы.\n Возможно лучше вывести при первом запуске один раз")

    else:
        print("Вы допустили ошибку. Введите повторно.\n\n")
        menuststus = "1"

# Снова стартовое меню    


