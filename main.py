import os
import requests
import json
import pandas as pd
from datetime import datetime

# на время запила
from pprint import pprint
from decorator import decor_arg



def get_MOEX_tickers_slovar():
    ''' Получение словаря тикеров торгуемых акций на Московской бирже через API
    формируется в словарь. Пример:
    {'ABRD': {'BOARDID': 'TQBR',
          'BOARDNAME': 'Т+: Акции и ДР - безадрес.',
          'CURRENCYID': 'SUR',
          'DECIMALS': 1,
          'FACEUNIT': 'SUR',
          'FACEVALUE': 1,
          'INSTRID': 'EQIN',
          'ISIN': 'RU000A0JS5T7',
          'ISSUESIZE': 98000184,
          'LATNAME': 'Abrau-Durso ao',
          'LISTLEVEL': 3,
          'LOTSIZE': 10,
          'MARKETCODE': 'FNDT',
          'MINSTEP': 0.5,
          'PREVADMITTEDQUOTE': 193,
          'PREVDATE': '2021-12-20',
          'PREVLEGALCLOSEPRICE': 193,
          'PREVPRICE': 193,
          'PREVWAPRICE': 193,
          'REGNUMBER': '1-02-12500-A',
          'REMARKS': None,
          'SECID': 'ABRD',
          'SECNAME': 'Абрау-Дюрсо ПАО ао',
          'SECTORID': None,
          'SECTYPE': '1',
          'SETTLEDATE': '2021-12-23',
          'SHORTNAME': 'АбрауДюрсо',
          'STATUS': 'A'},
          ...
    '''
    dict_TICKs = {}
    URL ="http://iss.moex.com/iss/engines/stock/markets/shares/boards/TQBR/securities.json?iss.meta=off&iss.only=securities&securities"
    response = requests.get(URL).json()
    number_of_dict = 1
    for row in response['securities']['data']:
        dict_TICKs[number_of_dict] = {
                f"{response['securities']['columns'][0]}" : row[0],
                f"{response['securities']['columns'][1]}" : row[1],
                f"{response['securities']['columns'][2]}" : row[2],
                f"{response['securities']['columns'][3]}" : row[3],
                f"{response['securities']['columns'][4]}" : row[4],
                f"{response['securities']['columns'][5]}" : row[5],
                f"{response['securities']['columns'][6]}" : row[6],
                f"{response['securities']['columns'][7]}" : row[7],
                f"{response['securities']['columns'][8]}" : row[8],
                f"{response['securities']['columns'][9]}" : row[9],
                f"{response['securities']['columns'][10]}" : row[10],
                f"{response['securities']['columns'][11]}" : row[11],
                f"{response['securities']['columns'][12]}" : row[12],
                f"{response['securities']['columns'][13]}" : row[13],
                f"{response['securities']['columns'][14]}" : row[14],
                f"{response['securities']['columns'][15]}" : row[15],
                f"{response['securities']['columns'][16]}" : row[16],
                f"{response['securities']['columns'][17]}" : row[17],
                f"{response['securities']['columns'][18]}" : row[18],
                f"{response['securities']['columns'][19]}" : row[19],
                f"{response['securities']['columns'][20]}" : row[20],
                f"{response['securities']['columns'][21]}" : row[21],
                f"{response['securities']['columns'][22]}" : row[22],
                f"{response['securities']['columns'][23]}" : row[23],
                f"{response['securities']['columns'][24]}" : row[24],
                f"{response['securities']['columns'][25]}" : row[25],
                f"{response['securities']['columns'][26]}" : row[26],
                f"{response['securities']['columns'][27]}" : row[27]
        }
        number_of_dict +=1
    return dict_TICKs

def json_create_MOEX_tickers(dictionary,filename):
    '''    Создает json словарь со всеми акциями торгуемыми на бирже
    ******************** РАБОТАЕТ, доработки не требуется ********************'''
    jsonData = json.dumps(dictionary)
    with open(filename,'w',encoding ='utf-8') as file:
        file.write(jsonData)

def xlsx_create_MOEX_tickers(data):
    '''    Создает xlsx книгу со всеми ценными бумагами торгуемыми на Московской бирже    '''
    count = len(list(data.keys()))
    df = pd.DataFrame(data)
    data_in_file = df.transpose()
    filename = 'TICKs.xlsx'
    data_in_file.to_excel(filename, sheet_name='TICK', index=False)
    dict = data_in_file.to_dict('index')
    filename_j = 'TICKs.json'
    json_create_MOEX_tickers(dict,filename_j)

@decor_arg('logs_update.xlsx')
def write_MOEX_tickers():
    dictionary = get_MOEX_tickers_slovar()
    xlsx_create_MOEX_tickers(dictionary)

def read_MOEX_tickers():
    '''    Считываем словарь с тикерами акций из файла    '''
    with open('TICKs.json', 'r', encoding='utf-8') as fh:
        data = json.load(fh)
    return data

@decor_arg('logs.xlsx')
def read_MOEX_tickers_figi(SECID):
    ''' Возращает всю информацию по указанному тикеру в виде словаря '''
    data = read_MOEX_tickers()
    for key in data.keys():
        iskomoe = data[key]['SECID']
        if SECID == iskomoe:
            return data[key]

def file_yes():
    filename = 'logs.xlsx'
    filename2 = 'logs_update.xlsx'
    if os.path.isfile(filename) != True:
        res = pd.DataFrame()
        res.to_excel(filename, sheet_name = 'logs', index=False)
    if os.path.isfile(filename2) != True:
        res2 = pd.DataFrame()
        res2.to_excel(filename2, sheet_name = 'logs', index=False)


a = 0
file_yes()
write_MOEX_tickers()
while a != 'x':
    a = input('''Для обновления файлов введите "update"
    Для вывода всех тикеров введите "look"
    Для вывода информации по конкретному инструменту введите его тикер
    P.s. декоратор применяется к запросам по тикеру
    ''')
    if a == 'update':
        write_MOEX_tickers()
    elif a == 'look':
        tickers = read_MOEX_tickers()
        for i in tickers:
            print(tickers[i]['SECID'])
    elif len(a) > 1:
        read_MOEX_tickers_figi(a)
