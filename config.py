# yandex
FROM_ADDR = '@yandex.ru'
PASSWORD = ''
'''
# google
FROM_ADDR_google = '@google.com'
PASSWORD_google = ''
# microsoft
FROM_ADDR_microsoft = '@microsoft.com'
PASSWORD_microsoft = ''
# yahoo
FROM_ADDR_yahoo = '@yahoo.com'
PASSWORD_yahoo = ''
'''

# путь к папке с файлами
DIR = r''
# возможные расширения файлов
EXT = ('.csv', '.xlsx', '.xls', '.json',)
# шаблоны текста
TEMPLATES_TEXT = [
        f'Здравствуйте, to_addr_name\n\nСобрание пройдет __.__ в __.__ в аудитории mm\n\nС уважением, Mr. X',
        f'Здравствуйте, to_addr_name\n\nПредлагаем пройти что-то где-то\n\nС уважением, Mr. X'
]
# тема письма
SUBJECT = 'Mr. X'
