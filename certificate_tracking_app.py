from get_certificate_expiration_date import get_certificate_expiration_date
from get_annual_control import get_annual_control
from easygui import msgbox, buttonbox, multenterbox
import os
import openpyxl


wb = openpyxl.load_workbook("certificate_tracking_app.xlsx")
sheet_certdays = wb["certdays"]

# Сохраняем в списке названия кнопок
options = ['Сведения об аттестатах','Ежегодный контроль','Открыть файл "Excel', 'Задать количество дней для аттестатов и ежегодного контроля','Отчет']
options_2 = ["Вставить в Word"]
fields_names = ["Количество дней для аттестатов", "Количество дней для ежегодного контроля"]

# Код, привязывающий функции к кнопкам
if get_certificate_expiration_date() == "OK": # "ОК" должно быть на английской раскладке
    while True:
        button = buttonbox(f"\tНажмите на соответствующую кнопку для получения сведений об аттестатах: ","ЦИТСиЗИ МВД по Республике Калмыкия", choices=options)
        if button == options[0]:
            get_certificate_expiration_date()
        elif button == options[1]:
            get_annual_control()
        elif button == options[2]:
            os.startfile("certificate_tracking_app.xlsx")
            break
        elif button == options[3]:
            try:
                input_fields = multenterbox("Задайте количество дней", "ЦИТСиЗИ МВД по Республике Калмыкия",fields_names)
                sheet_certdays["B1"] = int(input_fields[0]) 
                sheet_certdays["B2"] = int(input_fields[1])                
            except ValueError:
                msgbox("Введите целые числа","ЦИТСиЗИ МВД по Республике Калмыкия")
            else:
                wb.save("certificate_tracking_app.xlsx")
                break
        elif button == options[4]:
            btn = buttonbox(f"Нажмите на кнопку Word для создания отчета", "ЦИТСиЗИ МВД по Республике Калмыкия", choices=options_2)
        else:
            break


      







    









    



    
        
            
            
    
    
    
    


    


