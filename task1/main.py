import sys
import os
from docx import Document

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from laba_logging import *
InitLogFile()

from replacement import *

InitFile(__file__)

def err():
    print("Вы что-то сделали неверно, попробуйте еще раз")
    return input()


def Referat():
    print("Подготовка к созданию...")
    doc = Document("D:\\projects\\VisualStudioCode\\Laba_2_2_5_Graphics\\templates\\Реферат.docx")
    dict_data = {
        '__authors__' : "me"
    }
    print("Введите авторов через запятую:    ")
    tmp = input().replace(",", ",\n")
    while len(tmp.split(".")) < 3:
        tmp = err().replace(",", ",\n")
    dict_data['__authors__'] = tmp


    print("Ведите правообладателя (ФИО) или Название организации:    ")
    dict_data['__director__'] = input()

    print("Ведите название вашего модуля:    ")
    dict_data['__program__'] = input()
    name = f"Реферат {dict_data['__program__']}.docx"

    print("Придумайте аннотацию к вашему модулю:    ")
    print("\t- Для чего нужен модуль. \n\t- Что использует модуль для работы \n\t- Область применения данного модуля")
    tmp = input()
    while (tmp.find("модул") < 0) or (tmp.find("для") < 0) or (len(tmp) < 15):
        tmp = err()
    dict_data['__annotation__'] = tmp

    print("Укажите тип ЭВМ:    ")
    dict_data['__type__'] = input()

    print("Укажите язык создания:    ")
    tmp = input()
    while True:
        f = True
        for tmp_ in tmp.split(", "):
            if tmp_ not in ['C++', "python", "java", 'питон', "C#", "C", ]:
                tmp = err()
                f = False
                break
        if f == True:
            break
    dict_data['__language__'] = tmp
    
    print("Укажите операционную систему:    ")
    dict_data['__os__'] = input()
    
    print("Укажите затрачиваемое количество памяти (в Мб):    ")
    dict_data['__memory__'] = input()
    print("Сохраняется...    ")

    replace(doc, dict_data)
    doc.save(name)
    print(f"Сохранено {name}! ")




def ConsentToInformation():
    '''
Белов Александр Кириллович
22.04.2007
admin
Россия, Совхозная, 39а
Русское
0
ксорти
50
'''
    print("Подготовка к созаднию...")
    doc = Document("D:\\projects\\VisualStudioCode\\Laba_2_2_5_Graphics\\templates\\Согласие на указание сведений об авторе.docx")
    this_dict = {
        '__name__' : "program"
        }

    print("Введите ФИО полностью: ")
    this_dict['__full_user_name__'] = input()
    user_name = this_dict['__full_user_name__'].split(" ")
    for i in range(0, len(user_name)):
        if i == 0:
            this_dict['__user_name__'] = user_name[i] + " "
        else:
            this_dict['__user_name__'] += user_name[i][0] + ". "
    

    print("Введите дату рождения DD.MM.YYYY: ")
    date = input().split(".")
    while True:
        try:
            this_dict['__day__'] =   str(int(date[0]))
            this_dict['__month__'] = str(int(date[1]))
            this_dict['__year__'] =  str(int(date[2]))
            break
        except:
            print("что то пошло не так, попробуйте еще")

    print("Введите свою должность: ")
    this_dict['__post__'] = input()

    print("Введите место постоянного жительства, включая указание страны: ")
    this_dict['__residence__'] = input()    
    
    print("Введите гражданство: ")
    this_dict['__nationality__'] = input()
    
    print("Выбрать что-то одно из нижеприведенного: \n0 - если Физ. Лицо \n1 - если Юр. Лицо")
    while(1):
        a = input()
        if a == "0":
            print("секунду ... ")
            this_dict['__info__'] = this_dict['__user_name__'] + " " + this_dict['__residence__']
            break
        if a == "1":
            print("Введите наименование, место нахождения, основной государственный регистрационный номер (ОГРН)  и идентификационный номер налогоплательщика (ИНН)")
            this_dict['__info__'] = input()
            break
        else:
            print("что то не так, попробуйте еще раз")

    
    
    print("Введите название проекта: ")
    this_dict['__name__'] = input()

    print("Укажите сколько процентов работы вы выполняли: ")
    percent = input()
    while True:
        try:
            this_dict['__percent__'] = str(int(percent)%101)
            break
        except:
            print("что то пошло не так, попробуйте еще")

    # print("Укажите сколько процентов работы вы выполняли: ")
    # this_dict['__residence__'] = input()

    replace(doc, this_dict)

    doc.save("Согласие на указание сведений об " + this_dict['__full_user_name__'] + ".docx")


ConsentToInformation()
# Referat()
# A.K.<tkjd^W
# sghfggfgn
# zdsfbhdgnfgh
# модуль нужен для тестирования
# ывпв
# C
# zdfgdf
# 10
'''
Белов Алексаендр Кириллович
22.04.2007
разработчик
россия д5
российское
0
бот
20
'''
