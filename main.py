import pandas as pd
import csv
import json
import os
import smtplib
import re

from config import *


def filter_text(text: str) -> str:
    '''
    фильтрует текст для дальнейшего сравнения
    '''
    
    # к нижнему регистру 'ABC' -> 'abc'
    text = text.lower()
    # удаление пробелов по краям ' abc ' -> 'abc'
    text = text.strip()
    # найти все знаки препинания и заменить их на пустоту ''
    # [^ ] - 'набор символов кроме'
    # \w - символы из которых состоят слова (word) -> a-z0-9
    # \s - символы пробела (space) -> \t \r \n \s
    expression = r"[^\w\s]"
    text = re.sub(expression, "", text)
    
    return text


def valid_column(name_column: list) -> list:
    '''
    проверяет название столбцов
    '''
    
    shablon_name = ['name', 'фио', 'имя']
    shablon_addr = ['email', 'емаил', 'почта']
    name_index, addr_index = None, None

    for i in range(len(name_column)):
        name_column[i] = filter_text(name_column[i])
        if name_column[i] in shablon_name:
            name_index = i
        elif name_column[i] in shablon_addr:
            addr_index = i
        
        if name_index != None and addr_index != None:
            return [name_index, addr_index]
        
    if name_index == None and addr_index == None:
        return print('Что-то пошло не так с названиями столбцов таблицы --_--')


def valid_email(to_addrs_dict: dict) -> dict:
    '''
    проверка email на корректность
    '''
    
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    error_lst = []
    for name, email in to_addrs_dict.items():
        if re.match(pattern, email) is not None:
            continue
        else:
            error_lst.append(name)
            
    for name in error_lst:
        to_addrs_dict.pop(name)
        print(f'Некорректный email {email} удален из рассылки')
        
    return to_addrs_dict

    
def get_to_addr_xls(path_file) -> dict:
    '''
    получение словаря с именем и email из xls
    '''
    
    to_addr = {}
    
    with open(path_file, 'rb') as file:
        file = pd.read_excel(file)
        name_column = file.columns.tolist()
        # проверка названия столбцов
        column_index = valid_column(name_column)
        name = file.iloc[:, column_index[0]].tolist()
        addr = file.iloc[:, column_index[1]].tolist() 
 
        for index in range(len(name)):
            to_addr[name[index]] = addr[index]
        
    return to_addr

    
def get_to_addr_xlsx(path_file) -> dict:
    '''
    получение словаря с именем и email из xlsx
    '''
    
    to_addr = {}
    
    with open(path_file, 'rb') as file:
        file = pd.read_excel(file)
        name_column = file.columns.tolist()
        # проверка названия столбцов
        column_index = valid_column(name_column)
        name = file.iloc[:, column_index[0]].tolist()
        addr = file.iloc[:, column_index[1]].tolist() 
 
        for index in range(len(name)):
            to_addr[name[index]] = addr[index]
        
    return to_addr


def get_to_addr_csv(path_file) -> dict:
    '''
    получение словаря с именем и email из csv
    '''
    
    to_addr = {}
    
    with open(path_file, 'r', newline='') as file:  
        file = csv.reader(file)
        # убрать строчку с заголовками
        name_column = next(file)
        for row in file:
            row = row[0].split(';')
            to_addr[row[0]] = row[1]
        
    return to_addr


def get_to_addr_json(path_file) -> dict:
    '''
    получение словаря с именем и email из json
    '''
    
    to_addr = {}
    
    with open(path_file, 'r') as file: 
        file = file.read()
        to_addr = json.loads(file)
         
    return to_addr


def create_message(to_addr: dict, templ_text: list) -> list:
    '''
    подставляет имена в текст сообщения
    '''
    text = []
    for name in to_addr.keys():
        shablon = templ_text[0].replace('to_addr_name', name)
        text.append(shablon)
    
    return text


def log_mail(from_addr: str, psw: str):
    '''
    авторизация в почте и открытие соединения
    '''
    
    # server = input('Введите сервис, откуда необходимо осуществить рассылку:\n\tyandex\tgoogle\tyahoo\tmicrosoft\n')
    
    server = 'yandex'
    # создание объекта smtplib
    # 587 - порт
    if server == 'yandex':
        # yandex  НЕ МЕНЯТЬ ПОРТ
        smtpObj = smtplib.SMTP_SSL('smtp.yandex.ru:465')
    # elif server == 'google':
        # google -> не поддерживает больше сторонние приложения
        # smtpObj = smtplib.SMTP_SSL('smtp.gmail.com', 587)
        # smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    # elif server == 'yahoo':
        # yahoo -> нельзя создать пароль для приложения -> не работает
        # smtpObj = smtplib.SMTP_SSL('smtp.mail.yahoo.com:465')
    # elif server == 'microsoft':
        # microsoft -> нельзя создать почту даже
        # smtpObj = smtplib.SMTP_SSL('smtp.live.com', 587)
    else:
        print('Что-то пошло не так в создании объекта smtplib --_--')
    
    print('Соеднение открыто\n')
    
    try:
        # авторизация
        smtpObj.login(from_addr, psw)
        print('Авторизация/аутентификация прошла успешно\n')
        
    except smtplib.SMTPAuthenticationError:
        print('Что-то пошло не так при авторизации/аутентификации --_--')
    
    return smtpObj
    

def close_mail(smtpObj):
    # закрыть соединение
    smtpObj.quit()
    print('\nСоеднение закрыто\n')
        
        
def send_mail(smtpObj, from_addr: str, to_addrs: str, msg: str):
    '''
    отправка сообщения
    '''
    
    # отправка сообщения
    smtpObj.sendmail(from_addr, to_addrs, msg)
    print(f'Сообщение отправлено на почту:\n{to_addrs}')
       
    
if __name__ == '__main__':
  
    if os.listdir(DIR):
        lst_files = [f for f in os.listdir(DIR) if f.endswith(EXT)]
        print(f'В данной директории найдены подходящие файлы:\n{lst_files}\n')
        
        name_file = './' + input('Введите название необходимого файла:\n')
        if name_file.find('.xlsx') != -1:
            to_addr_dict = get_to_addr_xlsx(name_file)
        elif name_file.find('.xls') != -1:
            to_addr_dict = get_to_addr_xls(name_file)
        elif name_file.find('.csv') != -1:
            to_addr_dict = get_to_addr_csv(name_file)
        elif name_file.find('.json') != -1:
            to_addr_dict = get_to_addr_json(name_file)
        else:
            print('Что-то пошло не так --_--')
        
        to_addr_dict = valid_email(to_addr_dict)
        obj = log_mail(FROM_ADDR, PASSWORD)
        
        text = create_message(to_addr_dict, TEMPLATES_TEXT)
        
        for name, addr in to_addr_dict.items():
            for i in text:
                if name in i:
                    text_person = i
                    BODY = '\r\n'.join((
                        'From: %s' % FROM_ADDR,
                        'Subject: %s' % SUBJECT,
                        '',
                        text_person
                    )).encode(encoding='utf-8')
                    send_mail(obj, FROM_ADDR, addr, BODY)
                    break
                    
        # закрыть соединение
        close_mail(obj)
    
    else:
        print('В данной директории нет подходящих файлов')
