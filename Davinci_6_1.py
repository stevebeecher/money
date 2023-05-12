  
################################################
################################################
#  Project Authors: John Hirdt, Steve Beecher
#  Project Support: Sergio Ali 
#  Project: Davinci Code
#  Date: 12.17.2021
#  Version: 6.0.0
#  1100 lines --> --> XXX Lines
################################################
################################################
###########  Imports  ##########################
################################################
import gspread
import ast
import pytz
import traceback
import certifi
import json
#import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from urllib.request import Request, urlopen
import sched, time 
import schedule 
import send_text_message
import linecache
import sys
import requests
import random 
import ssl
################################################
###########  Functions   #######################
################################################
def time_is():
    date_time = str(datetime.now(pytz.timezone('US/Eastern')))
    date = date_time.split(' ')[0]
    time = date_time.split(' ')[1]
    hour = int(time.split(':')[0])
    minute = int(time.split(':')[1])
    #second = int(time.split(':')[2].split('.')[0])
    second = int(15)
    return hour, minute, second
################################################
def get_today():
    date = str(datetime.now(pytz.timezone('US/Eastern')))
    month = date.split('-')[1]
    if int(month) == 10 or int(month) == 11 or int(month) == 12:
        pass
    else:
        month = month.split('0')[-1]

    day = date.split('-')[2].split(' ')[0]
    if int(day) < 10: day = day.split('0')[1] #gets rid of the zero from days less than 10
    year = date.split('-')[0]
    today = month+'/'+day+'/'+year
    return today
################################################
def get_sheet_instance(symb):
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('DavinciTwoPointO-e22f61dd66df.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1veNKTTtlS8lOBT49YvdQT0J5JOmfbSUb5j_lQCyiRZ0")
    sheet_instance = sheet.worksheet(symb)
    return sheet_instance
################################################

def new_row(symb, row_dict):
    today = get_today()
    sheet_instance = get_sheet_instance(symb)
    next_row = next_available_row(sheet_instance)
    row = int(next_row)
    row_dict[symb] = row
    A = today
    if(symb == "AAPL"): col = 'B'
    if(symb == "AMZN"): col = 'C'
    if(symb == "BABA"): col = 'D'
    if(symb == "GS"): col = 'E'
    if(symb == "META"): col = 'F'
    if(symb == "JD"): col = 'G'
    if(symb == "MRVL"): col = 'H'
    if(symb == "MSFT"): col = 'I'
    if(symb == "NFLX"): col = 'J'
    if(symb == "NVDA"): col = 'K'
    if(symb == "QQQ"): col = 'L'
    if(symb == "SNAP"): col = 'M'
    if(symb == "SPY"): col = 'N'
    if(symb == "TSLA"): col = 'O'
    if(symb == "WFC"): col = 'P'
    if(symb == "DKNG"): col = 'Q'
    if(symb == "NIO"): col = 'R'
    if(symb == "DIS"): col = 'S'
    if(symb == 'JNJ'): col = 'T'
    if(symb == 'NET'): col = 'U'
    if(symb == 'PYPL'): col = 'V'
    if(symb == 'QRVO'): col = 'W'
    if(symb == 'SQQQ'): col = 'X'
    if(symb == 'PINS'): col = 'Y'
    if(symb == 'EBAY'): col = 'Z'
    if(symb == 'CAT'): col = 'AA'
    if(symb == 'DHI'): col = 'AB'
    if(symb == 'GOOGL'): col = 'AC'
    if(symb == 'ZM'): col = 'AD'
    if(symb == 'DT'): col = 'AE'
    if(symb == 'QFIN'): col = 'AF'
    if(symb == 'F'): col = 'AG'
    if(symb == 'BA'): col = 'AH'
    if(symb == 'JETS'): col = 'AI'
    B = "=StockData!" + col + '8'
    C = "=StockData!" + col + '2' # open
    D = "=D" + str(row-1) + "+1" 
    E = "=B" +str(row) + "-$X$2"
    F = "=(G" + str(row) + "+E" + str(row) + ")/2"
    G = "=E" + str(row) + "+$AC$2"
    H = "=B" + str(row) + "-$AB$2"
    I = "=StockData!" + col + "4" # low 
    J = "=E" + str(row) + "+$J$2*$AC$2"
    K = "=StockData!" + col + "5" #close  
    if(symb == "AAPL"): reg = "=0.1863*(D"+str(row)+')+47.546'
    if(symb == "AMZN"): reg = "=0.1611*(D"+str(row)+')+80'
    if(symb == "BABA"): reg = "=0.1564*(D"+str(row)+')+217.84'
    if(symb == "GS"): reg = "=1.2335*(D"+str(row)+')+206.63'
    if(symb == "META"): reg = "=0.1387*(D"+str(row)+')+176.66'
    if(symb == "JD"): reg = "=0.2027*(D"+str(row)+')+22.952'
    if(symb == "MRVL"): reg = "=0.0522*(D"+str(row)+')+21.586'
    if(symb == "MSFT"): reg = "=0.3085*(D"+str(row)+')+135.19'
    if(symb == "NFLX"): reg = "=1.0441*(D"+str(row)+')+246.91'
    if(symb == "NVDA"): reg = "=0.292*(D"+str(row)+')+29.2'
    if(symb == 'QQQ'): reg = "=0.1208*(D"+str(row)+')+95.709'
    if(symb == 'SNAP'): reg = "=0.0847*(D"+str(row)+')+8.3067'
    if(symb == 'SPY'): reg = "=0.101*(D"+str(row)+')+205.08'
    if(symb == 'TSLA'): reg = "=0.8081*(D"+str(row)+')+80.85'
    if(symb == "WFC"): reg = "=-0.0162*(D"+str(row)+')+27.398'
    if(symb == "DKNG"): reg = "=0.1711*(D"+str(row)+')+12.659'
    if(symb == "NIO"): reg = "=0.2254*(D"+str(row)+')-9.1507'
    if(symb == "DIS"): reg = "=0.3399*(D"+str(row)+')+88.887'
    if(symb == "JNJ"): reg = "=0.0564*(D"+str(row)+')+134.88'
    if(symb == "NET"): reg = "=0.3559*(D"+str(row)+')+25.341'
    if(symb == "PYPL"): reg = "=0.4727*(D"+str(row)+')+79.715'
    if(symb == "QRVO"): reg = "=0.4727*(D"+str(row)+')+79.715'
    if(symb == "SQQQ"): reg = "=-0.4533*(D"+str(row)+')+206.95'
    if(symb == 'PINS'): reg = "=0.2582*(D"+str(row)+')+4.4069'
    if(symb == 'EBAY'): reg = "=0.0774*(D"+str(row)+')+40.321'
    if(symb == 'CAT'): reg = "=0.4712*(D"+str(row)+')+90.463'
    if(symb == 'DHI'): reg = "=0.0674*(D"+str(row)+')+35.307'
    if(symb == 'GOOGL'): reg = "=0.0964*(D"+str(row)+')+73.136'
    if(symb == 'ZM'): reg = "=1.5398*(D"+str(row)+')+92.041'
    if(symb == 'DT'): reg = "=0.0731*(D"+str(row)+')+21.302'
    if(symb == 'QFIN'): reg = "=0.022*(D"+str(row)+')+7.3941'
    if(symb == 'F'): reg = "=0.0289*(D"+str(row)+')+4.2409'
    if(symb == 'BA'): reg ="=0.3627*(D"+str(row)+')+129.91'
    if(symb == 'JETS'): reg = "=0.0469*(D"+str(row)+')+12.327'
    L = reg
    M = "=I" +str(row) + "-L" + str(row)
    N = "=M" + str(row -1) + "-M" + str(row)
    #day_num = int(sheet_instance.cell(row, 4).value) + 1
    day_num = row - 3
    if(symb == "AAPL"): reg1 = "(0.1863*("+str(day_num)+')+47.546))'
    if(symb == "AMZN"): reg1 = "(0.1611*("+str(day_num)+')+80))'
    if(symb == "BABA"): reg1 = "(0.1564*("+str(day_num)+')+217.84))'
    if(symb == "GS"): reg1 = "(1.2335*("+str(day_num)+')+206.63))'
    if(symb == "META"): reg1 = "(0.1387*("+str(day_num)+')+176.66))'
    if(symb == "JD"): reg1 = "(0.2027*("+str(day_num)+')+22.952))'
    if(symb == "MRVL"): reg1 = "(0.0522*("+str(day_num)+')+21.586))'
    if(symb == "MSFT"): reg1 = "(0.3085*("+str(day_num)+')+135.19))'
    if(symb == "NFLX"): reg1 = "(1.0441*("+str(day_num)+')+246.91))'
    if(symb == "NVDA"): reg1 = "(0.292*("+str(day_num)+')+29.2))'
    if(symb == "QQQ"): reg1 = "(0.1208*("+str(day_num)+')+95.709))'
    if(symb == "SNAP"): reg1 = "(0.0847*("+str(day_num)+')+8.3067))'
    if(symb == "SPY"): reg1 = "(0.101*("+str(day_num)+')+205.08))'
    if(symb == "TSLA"): reg1 = "(0.8081*("+str(day_num)+')+80.85))'
    if(symb == "WFC"): reg1 = "(-0.0162*("+str(day_num)+')+27.398))'
    if(symb == "DKNG"): reg1 = "(0.1711*("+str(day_num)+')+12.659))'
    if(symb == "NIO"): reg1 = "(0.2254*("+str(day_num)+')-9.1507))'
    if(symb == "DIS"): reg1 = "(0.3399*("+str(day_num)+')+88.887))'
    if(symb == "JNJ"): reg1 = "(0.0564*("+str(day_num)+')+134.88))'
    if(symb == "NET"): reg1 = "(0.3559*("+str(day_num)+')+25.341))'
    if(symb == "PYPL"): reg1 = "(0.4727*("+str(day_num)+')+79.715))'
    if(symb == "QRVO"): reg1 = "(0.4727*("+str(day_num)+')+79.715))'
    if(symb == "SQQQ"): reg1 = "(-0.4533*("+str(day_num)+')+206.95))'
    if(symb == "PINS"): reg1 = "(0.2582*("+str(day_num)+')+4.4069))'
    if(symb == "EBAY"): reg1 = "(0.0774*("+str(day_num)+')+40.321))'
    if(symb == "CAT"): reg1 = "(0.4712*("+str(day_num)+')+90.463))'
    if(symb == 'DHI'): reg1 = "(0.0674*("+str(day_num)+')+35.307))'
    if(symb == "GOOGL"): reg1 = "(0.0964*("+str(day_num)+')+73.136))'
    if(symb == "ZM"): reg1 = "(1.5398*("+str(day_num)+')+92.041))'
    if(symb == "DT"): reg1 = "(0.0731*("+str(day_num)+')+21.302))'
    if(symb == "QFIN"): reg1 = "(0.022*("+str(day_num)+')+7.3941))'
    if(symb == 'F'): reg1 = "(0.0289*("+str(day_num)+')+4.2409))'
    if(symb == 'BA'): reg1 = "(0.3627*("+str(day_num)+')+129.91))'
    if(symb == 'JETS'): reg1 = "(0.0469*("+str(day_num)+')+12.327))'
    O = "=((M" + str(row) + "-$O$1) + " + str(reg1)
    P = "=StockData!" + col + "3" #put high here
    Q = "=S"+str(row) + "-($Q$2*$AC$2)"
    R = "=(S"+str(row)+"+Q" + str(row)+")/2"
    S = "=B"+ str(row)+"+$X$2"
    T = "=S" + str(row) + "+$AB$2"
    U = "=(K" + str(row) + "-I" + str(row)+")/I" + str(row)
    V = "=(P" + str(row) + "-K" + str(row) + ")/K" + str(row)
    W = "=B" + str(row) + "-$AG$1"
    X = "=X" + str(row - 1) #previous days atr here, for pre/mkt, actual atr added in later during pull top 5
    Y = "=(I" + str(row) + "-L" + str(row) + ")/I" + str(row)
    Z = "=U"+str(row)+"-V"+str(row)
    AA = "=(I"+str(row)+"-I"+str(row-1) + ")/I"+str(row-1)
    AB = "=AVERAGE(C" + row +',E' +row+":H"+row+',J'+row+',W'+row+',Q'+row-1+',R'+row-1+')'
    AC = "=Z" + str(row) + "+AA" + str(row) 
    AD = '=ABS(E' + str(row) + '-I'+str(row)+')'
    AE = "=AVERAGE(AD" + str(int(row-41))+":AD"+str(row) + ")"
    AF = "=(AF"+str(int(row-1))+"+1)"
    AG = "=C" + str(row) + "-I" + str(row)
    AH = "=AG" + str(row) + "-$AG$1"
    AI = "=(AH" + str(row) + ")^2"
    AJ = ""
    AK = ""
    AL = ""
    AM = ""
    AN = ""
    AO = ""
    AP = ""
    AQ = ""
    AR = ""
    AS = ""
    AT = "=B" + str(row) + "*(2/($AT$2 + 1)) + AT" + str(row-1) + "*(1-(2/($AT$2 + 1)))" #EMA 10
    AU = "=B" + str(row) + "*(2/($AU$2 + 1)) + AU" + str(row-1) + "*(1-(2/($AU$2 + 1)))" #EMA 21
    AV = "=AVERAGE(K"+str(row)+":K" + str(row-50) + ")"
    AW = "=(index("+'$K3:$K, counta('+'$K3:$K))-AV'+str(row)+')/AV'+str(row)
    AX = "=(index("+'$K3:$K, counta('+'$K3:$K))-AU'+str(row)+')/AU'+str(row)
    AY = "=(index("+'$K3:$K, counta('+'$K3:$K))-AT'+str(row)+')/AT'+str(row)
    sheet_instance.batch_update([{'range': 'A'+str(row)+':AY'+str(row), 'values':[[A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY]]}], value_input_option = 'USER_ENTERED')
    return row_dict
################################################
def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)
################################################
def get_row_dict( symb, row_dict):
    today = get_today()
    sheet_instance = get_sheet_instance(symb)
    next_row = next_available_row(sheet_instance)
    row = int(next_row)
    row_dict[symb] = row- 1
    return row_dict
################################################
def end_the_day(symb, row_dict):
    #this code runs only once for each stock
    today = get_today()
    sheet_instance = get_sheet_instance('StockData')
    row = row_dict[symb]
    O = ""; H = ""; L = ""; C = ""; T = ""; P = "";
    if symb == "AAPL":
        O = sheet_instance.cell(2, 2).value
        H = sheet_instance.cell(3, 2).value
        L = sheet_instance.cell(4, 2).value
        C = sheet_instance.cell(5, 2).value
        T = sheet_instance.cell(6, 2).value
        P = sheet_instance.cell(8, 2).value
    if symb == "AMZN":
        O = sheet_instance.cell(2, 3).value
        H = sheet_instance.cell(3, 3).value
        L = sheet_instance.cell(4, 3).value
        C = sheet_instance.cell(5, 3).value
        T = sheet_instance.cell(6, 3).value
        P = sheet_instance.cell(8, 3).value
    if symb == "BABA":
        O = sheet_instance.cell(2, 4).value
        H = sheet_instance.cell(3, 4).value
        L = sheet_instance.cell(4, 4).value
        C = sheet_instance.cell(5, 4).value
        T = sheet_instance.cell(6, 4).value
        P = sheet_instance.cell(8, 4).value
    if symb == "GS":
        O = sheet_instance.cell(2, 5).value
        H = sheet_instance.cell(3, 5).value
        L = sheet_instance.cell(4, 5).value
        C = sheet_instance.cell(5, 5).value
        T = sheet_instance.cell(6, 5).value
        P = sheet_instance.cell(8, 5).value
    if symb == "META":
        O = sheet_instance.cell(2, 6).value
        H = sheet_instance.cell(3, 6).value
        L = sheet_instance.cell(4, 6).value
        C = sheet_instance.cell(5, 6).value
        T = sheet_instance.cell(6, 6).value
        P = sheet_instance.cell(8, 6).value
    if symb == "JD":
        O = sheet_instance.cell(2, 7).value
        H = sheet_instance.cell(3, 7).value
        L = sheet_instance.cell(4, 7).value
        C = sheet_instance.cell(5, 7).value
        T = sheet_instance.cell(6, 7).value
        P = sheet_instance.cell(8, 7).value
    if symb == "MRVL":
        O = sheet_instance.cell(2, 8).value
        H = sheet_instance.cell(3, 8).value
        L = sheet_instance.cell(4, 8).value
        C = sheet_instance.cell(5, 8).value
        T = sheet_instance.cell(6, 8).value
        P = sheet_instance.cell(8, 8).value
    if symb == "MSFT":
        O = sheet_instance.cell(2, 9).value
        H = sheet_instance.cell(3, 9).value
        L = sheet_instance.cell(4, 9).value
        C = sheet_instance.cell(5, 9).value
        T = sheet_instance.cell(6, 9).value
        P = sheet_instance.cell(8, 9).value
    if symb == "NFLX":
        O = sheet_instance.cell(2, 10).value
        H = sheet_instance.cell(3, 10).value
        L = sheet_instance.cell(4, 10).value
        C = sheet_instance.cell(5, 10).value
        T = sheet_instance.cell(6, 10).value
        P = sheet_instance.cell(8, 10).value
    if symb == "NVDA":
        O = sheet_instance.cell(2, 11).value
        H = sheet_instance.cell(3, 11).value
        L = sheet_instance.cell(4, 11).value
        C = sheet_instance.cell(5, 11).value
        T = sheet_instance.cell(6, 11).value
        P = sheet_instance.cell(8, 11).value
    if symb == "QQQ":
        O = sheet_instance.cell(2, 12).value
        H = sheet_instance.cell(3, 12).value
        L = sheet_instance.cell(4, 12).value
        C = sheet_instance.cell(5, 12).value
        T = sheet_instance.cell(6, 12).value
        P = sheet_instance.cell(8, 12).value
    if symb == "SNAP":
        O = sheet_instance.cell(2, 13).value
        H = sheet_instance.cell(3, 13).value
        L = sheet_instance.cell(4, 13).value
        C = sheet_instance.cell(5, 13).value
        T = sheet_instance.cell(6, 13).value
        P = sheet_instance.cell(8, 13).value
    if symb == "SPY":
        O = sheet_instance.cell(2, 14).value
        H = sheet_instance.cell(3, 14).value
        L = sheet_instance.cell(4, 14).value
        C = sheet_instance.cell(5, 14).value
        T = sheet_instance.cell(6, 14).value
        P = sheet_instance.cell(8, 14).value
    if symb == "TSLA":
        O = sheet_instance.cell(2, 15).value
        H = sheet_instance.cell(3, 15).value
        L = sheet_instance.cell(4, 15).value
        C = sheet_instance.cell(5, 15).value
        T = sheet_instance.cell(6, 15).value
        P = sheet_instance.cell(8, 15).value
    if symb == "WFC":
        O = sheet_instance.cell(2, 16).value
        H = sheet_instance.cell(3, 16).value
        L = sheet_instance.cell(4, 16).value
        C = sheet_instance.cell(5, 16).value
        T = sheet_instance.cell(6, 16).value
        P = sheet_instance.cell(8, 16).value
    if symb == "DKNG":
        O = sheet_instance.cell(2, 17).value
        H = sheet_instance.cell(3, 17).value
        L = sheet_instance.cell(4, 17).value
        C = sheet_instance.cell(5, 17).value
        T = sheet_instance.cell(6, 17).value
        P = sheet_instance.cell(8, 17).value
    if symb == "NIO":
        O = sheet_instance.cell(2, 18).value
        H = sheet_instance.cell(3, 18).value
        L = sheet_instance.cell(4, 18).value
        C = sheet_instance.cell(5, 18).value
        T = sheet_instance.cell(6, 18).value
        P = sheet_instance.cell(8, 18).value 
    if symb == "DIS":
        O = sheet_instance.cell(2, 19).value
        H = sheet_instance.cell(3, 19).value
        L = sheet_instance.cell(4, 19).value
        C = sheet_instance.cell(5, 19).value
        T = sheet_instance.cell(6, 19).value
        P = sheet_instance.cell(8, 19).value
    if symb == "JNJ":
        O = sheet_instance.cell(2, 20).value
        H = sheet_instance.cell(3, 20).value
        L = sheet_instance.cell(4, 20).value
        C = sheet_instance.cell(5, 20).value
        T = sheet_instance.cell(6, 20).value
        P = sheet_instance.cell(8, 20).value   
    if symb == "NET":
        O = sheet_instance.cell(2, 21).value
        H = sheet_instance.cell(3, 21).value
        L = sheet_instance.cell(4, 21).value
        C = sheet_instance.cell(5, 21).value
        T = sheet_instance.cell(6, 21).value
        P = sheet_instance.cell(8, 21).value
    if symb == "PYPL":
        O = sheet_instance.cell(2, 22).value
        H = sheet_instance.cell(3, 22).value
        L = sheet_instance.cell(4, 22).value
        C = sheet_instance.cell(5, 22).value
        T = sheet_instance.cell(6, 22).value
        P = sheet_instance.cell(8, 22).value 
    if symb == "QRVO":
        O = sheet_instance.cell(2, 23).value
        H = sheet_instance.cell(3, 23).value
        L = sheet_instance.cell(4, 23).value
        C = sheet_instance.cell(5, 23).value
        T = sheet_instance.cell(6, 23).value
        P = sheet_instance.cell(8, 23).value
    if symb == "SQQQ":
        O = sheet_instance.cell(2, 24).value
        H = sheet_instance.cell(3, 24).value
        L = sheet_instance.cell(4, 24).value
        C = sheet_instance.cell(5, 24).value
        T = sheet_instance.cell(6, 24).value
        P = sheet_instance.cell(8, 24).value 
    if symb == "PINS":
        O = sheet_instance.cell(2, 25).value
        H = sheet_instance.cell(3, 25).value
        L = sheet_instance.cell(4, 25).value
        C = sheet_instance.cell(5, 25).value
        T = sheet_instance.cell(6, 25).value
        P = sheet_instance.cell(8, 25).value 
    if symb == "EBAY":
        O = sheet_instance.cell(2, 26).value
        H = sheet_instance.cell(3, 26).value
        L = sheet_instance.cell(4, 26).value
        C = sheet_instance.cell(5, 26).value
        T = sheet_instance.cell(6, 26).value
        P = sheet_instance.cell(8, 26).value
    if symb == "CAT":
        O = sheet_instance.cell(2, 27).value
        H = sheet_instance.cell(3, 27).value
        L = sheet_instance.cell(4, 27).value
        C = sheet_instance.cell(5, 27).value
        T = sheet_instance.cell(6, 27).value
        P = sheet_instance.cell(8, 27).value
    if symb == "DHI":
        O = sheet_instance.cell(2, 28).value
        H = sheet_instance.cell(3, 28).value
        L = sheet_instance.cell(4, 28).value
        C = sheet_instance.cell(5, 28).value
        T = sheet_instance.cell(6, 28).value
        P = sheet_instance.cell(8, 28).value 
    if symb == "GOOGL":
        O = sheet_instance.cell(2, 29).value
        H = sheet_instance.cell(3, 29).value
        L = sheet_instance.cell(4, 29).value
        C = sheet_instance.cell(5, 29).value
        T = sheet_instance.cell(6, 29).value
        P = sheet_instance.cell(8, 29).value
    if symb == "ZM":
        O = sheet_instance.cell(2, 30).value
        H = sheet_instance.cell(3, 30).value
        L = sheet_instance.cell(4, 30).value
        C = sheet_instance.cell(5, 30).value
        T = sheet_instance.cell(6, 30).value
        P = sheet_instance.cell(8, 30).value
    if symb == 'DT':  
        O = sheet_instance.cell(2, 31).value
        H = sheet_instance.cell(3, 31).value
        L = sheet_instance.cell(4, 31).value
        C = sheet_instance.cell(5, 31).value
        T = sheet_instance.cell(6, 31).value
        P = sheet_instance.cell(8, 31).value
    if symb == "QFIN":
        O = sheet_instance.cell(2, 32).value
        H = sheet_instance.cell(3, 32).value
        L = sheet_instance.cell(4, 32).value
        C = sheet_instance.cell(5, 32).value
        T = sheet_instance.cell(6, 32).value
        P = sheet_instance.cell(8, 32).value 
    if symb == "F":
        O = sheet_instance.cell(2, 33).value
        H = sheet_instance.cell(3, 33).value
        L = sheet_instance.cell(4, 33).value
        C = sheet_instance.cell(5, 33).value
        T = sheet_instance.cell(6, 33).value
        P = sheet_instance.cell(8, 33).value 
    if symb == "BA":
        O = sheet_instance.cell(2, 34).value
        H = sheet_instance.cell(3, 34).value
        L = sheet_instance.cell(4, 34).value
        C = sheet_instance.cell(5, 34).value
        T = sheet_instance.cell(6, 34).value
        P = sheet_instance.cell(8, 34).value
    if symb == "JETS":
        O = sheet_instance.cell(2, 35).value
        H = sheet_instance.cell(3, 35).value
        L = sheet_instance.cell(4, 35).value
        C = sheet_instance.cell(5, 35).value
        T = sheet_instance.cell(6, 35).value
        P = sheet_instance.cell(8, 35).value  
    sheet_instance_1 = get_sheet_instance(symb)
    row = row_dict[symb]
    sheet_instance_1.update_cell(row, 3, O) #lock in open
    sheet_instance_1.update_cell(row, 16, H) #lock in high
    sheet_instance_1.update_cell(row, 9, L) #lock in low
    sheet_instance_1.update_cell(row, 11, C) #lock in close
    sheet_instance_1.update_cell(row, 28, T) #lock in time
    sheet_instance_1.update_cell(row, 2, P) #lock in pre/mkt 
    print(symb, ' locked in')
################################################
def pull_atr(symb, row_dict):
    today = get_today()
    sheet_instance = get_sheet_instance(symb)
    row = row_dict[symb]
   
    atr = "=0.2*(4*X"+str(row-1) + "+ max(p"+str(row) + "-I"+str(row)+",abs(p"+str(row) +"-k"+str(row-1) + "), abs(I"+str(row)+"-K"+str(row-1)+")))" 
    sheet_instance.update_cell(row, 24, atr)
################################################
def PrintException():
    face = open('errors.txt', 'a')
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    hour, minute, second = time_is()
    message = str(filename) + ' ' + str(lineno) + ' ' + str(exc_obj) + '' + 'at ' + str(hour) + ':' + str(minute)
    #message = 'Exception in ({}, Line {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj) + 'at ' + str(hour), ':' + str(minute)
    face.write( message )
    print(message)
    print(exc_type)
    face.write('\n')
    face.close()
    #time.sleep(20)
################################################
def plus_minus():
    user_agent_list = [
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1.1 Safari/605.1.15',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:77.0) Gecko/20100101 Firefox/77.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:77.0) Gecko/20100101 Firefox/77.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
    ]
    url = "https://finance.yahoo.com/quote/%5EDJI/history?p=%5EDJI"
    user_agent = random.choice(user_agent_list) 
    #req = Request(url, headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.13; rv:63.0) Gecko/20100101 Firefox/63.0'})
    req = Request(url, headers = {'User-Agent': user_agent})
    page = urlopen(req, timeout = 20)
    html_bytes = page.read()
    html = html_bytes.decode("utf-8")
    i_1 = html.find("% or ") + 5
    i_2 = i_1 + 50
    sp_change = html[i_1: i_2].split(' points')[0]
    sp_value = html[i_1 -50: i_2 + 200].split(' points to ')[-1].split(' points')[0]
    i_3 = html.find("% or ", i_2) + 5
    i_4 = i_3 + 50
    dow_change = html[i_3:i_4].split(' points')[0]
    dow_value = html[i_3 - 50: i_4 + 200].split(' points to ')[-1].split(' points')[0]
    i_5 = html.find("% or ", i_4) + 5
    i_6 = i_5 + 50
    nasdaq_change = html[i_5:i_6].split(' points')[0]
    nasdaq_value = html[i_5 - 50: i_6 + 200].split(' points to ')[-1].split(' points')[0]
    sheet_instance = get_sheet_instance('MoneyTree')
    sheet_instance.batch_update([{'range': 'D18:F18', 'values': [[sp_value, dow_value, nasdaq_value]]}, {'range': 'D19:F19', 'values': [[sp_change, dow_change, nasdaq_change]]}], value_input_option = 'USER_ENTERED') 
################################################
def pre_market_low(symb):
    return "=index(" + symb + "!w4:w,counta(" + symb + "!w4:w))"
################################################
def classification(symb):
    a = "=MoneyTree"
    if(symb == "AAPL"): a = a + "!T2"
    if(symb == "AMZN"): a = a + "!T3"
    if(symb == "BABA"): a = a + "!T4"
    if(symb == "GS"): a = a + "!T5"
    if(symb == "META"): a = a + "!T6"
    if(symb == "JD"): a = a + "!T7"
    if(symb == "MRVL"): a = a + "!T8"
    if(symb == "MSFT"): a = a + "!T9"
    if(symb == "NFLX"): a = a + "!T10"
    if(symb == "NVDA"): a = a + "!T11"
    if(symb == "QQQ"): a = a + "!T22"
    if(symb == "SNAP"): a = a + "!T13"
    if(symb == "SPY"): a = a + "!T23"
    if(symb == "TSLA"): a = a + "!T15"
    if(symb == "WFC"): a = a + "!T16"
    if(symb == "DKNG"): a = a + "2!T5"
    if(symb == "NIO"): a = a + "2!T12"
    if(symb == "DIS"): a = a + "2!T4"
    if(symb == 'JNJ'): a = a + "2!T10"
    if(symb == 'NET'): a = a + "2!T11"
    if(symb == 'PYPL'): a = a + "2!T14"
    if(symb == 'QRVO'): a = a + "2!T15"
    if(symb == 'SQQQ'): a = a + "!T21"
    if(symb == 'PINS'): a = a + "2!T13"
    if(symb == 'EBAY'): a = a + "2!T6"
    if(symb == 'CAT'): a = a + "2!T7"
    if(symb == 'DHI'): a = a + "2!T3"
    if(symb == 'GOOGL'): a = a + "2!T16"
    if(symb == 'ZM'): a = a + "2!T17"
    if(symb == 'DT'): a = a + "!T6"
    if(symb == 'QFIN'): a = a + "!T13"
    if(symb == 'F'): a = a + "2!T8"
    if(symb == 'BA'): a = a + "2!T2"
    if(symb == 'JETS'): a = a + "2!T9"
    return a
################################################
def batch_pre(stocks):
    pre_data = {}
    key = "624b51b375a1fae6682dc955e6159fc1"
    for sym in stocks:
      url = "https://financialmodelingprep.com/api/v3/quote/" +sym+ ",FB?apikey=" + key
      context = ssl.create_default_context(cafile=certifi.where())
      response = urlopen(url, context = context)
      data = response.read().decode("utf-8")
      r = json.loads(data) 
      dataIN = r[0]
      p = dataIN['price']
      pre_data[sym] = ("","", pre_market_low(sym),"","",classification(sym), p, 'test')
    return pre_data
################################################
def write_batch_pre(stock_data):
    sheet_instance = get_sheet_instance('StockData')
   
    #data for open
    B0 = stock_data['AAPL'][0]; B1 = stock_data['AMZN'][0]
    B2 = stock_data['BABA'][0]; B3 = stock_data['GS'][0]
    B4 = stock_data['META'][0]; B5 = stock_data['JD'][0]
    B6 = stock_data['MRVL'][0]; B7 = stock_data['MSFT'][0]
    B8 = stock_data['NFLX'][0]; B9 = stock_data['NVDA'][0]
    B10 = stock_data['QQQ'][0]; B11 = stock_data['SNAP'][0]
    B12 = stock_data['SPY'][0]; B13 = stock_data['TSLA'][0]
    B14 = stock_data['WFC'][0]; B15 = stock_data['DKNG'][0]
    B16 = stock_data['NIO'][0]; B17 = stock_data['DIS'][0]
    B18 = stock_data['JNJ'][0]; B19 = stock_data['NET'][0]
    B20 = stock_data['PYPL'][0]; B21 = stock_data['QRVO'][0]
    B22 = stock_data["SQQQ"][0]; B23 = stock_data['PINS'][0]
    B24 = stock_data['EBAY'][0]; B25 = stock_data["CAT"][0]
    B26 = stock_data['DHI'][0]; B27 = stock_data["GOOGL"][0]
    B28 = stock_data["ZM"][0]; B29 = stock_data["DT"][0]
    B30 = stock_data["QFIN"][0]; B31 = stock_data["F"][0]
    B32 = stock_data["BA"][0]; B33 = stock_data["JETS"][0]
    #data for high
    C0 = stock_data['AAPL'][1]; C1 = stock_data['AMZN'][1]
    C2 = stock_data['BABA'][1]; C3 = stock_data['GS'][1]
    C4 = stock_data['META'][1]; C5 = stock_data['JD'][1]
    C6 = stock_data['MRVL'][1]; C7 = stock_data['MSFT'][1]
    C8 = stock_data['NFLX'][1]; C9 = stock_data['NVDA'][1]
    C10 = stock_data['QQQ'][1]; C11 = stock_data['SNAP'][1]
    C12 = stock_data['SPY'][1]; C13 = stock_data['TSLA'][1]
    C14 = stock_data['WFC'][1]; C15 = stock_data['DKNG'][1]
    C16 = stock_data['NIO'][1]; C17 = stock_data['DIS'][1]
    C18 = stock_data['JNJ'][1]; C19 = stock_data['NET'][1]
    C20 = stock_data['PYPL'][1]; C21 = stock_data['QRVO'][1]
    C22 = stock_data["SQQQ"][1]; C23 = stock_data['PINS'][1]
    C24 = stock_data['EBAY'][1]; C25 = stock_data["CAT"][1]
    C26 = stock_data['DHI'][1]; C27 = stock_data["GOOGL"][1]
    C28 = stock_data["ZM"][1]; C29 = stock_data["DT"][1]
    C30 = stock_data["QFIN"][1]; C31 = stock_data["F"][1]
    C32 = stock_data["BA"][1]; C33 = stock_data["JETS"][1]
    #data for low 
    D0 = stock_data['AAPL'][2]; D1 = stock_data['AMZN'][2]
    D2 = stock_data['BABA'][2]; D3 = stock_data['GS'][2]
    D4 = stock_data['META'][2]; D5 = stock_data['JD'][2]
    D6 = stock_data['MRVL'][2]; D7 = stock_data['MSFT'][2]
    D8 = stock_data['NFLX'][2]; D9 = stock_data['NVDA'][2]
    D10 = stock_data['QQQ'][2]; D11 = stock_data['SNAP'][2]
    D12 = stock_data['SPY'][2]; D13 = stock_data['TSLA'][2]
    D14 = stock_data['WFC'][2]; D15 = stock_data['DKNG'][2]
    D16 = stock_data['NIO'][2]; D17 = stock_data['DIS'][2]
    D18 = stock_data['JNJ'][2]; D19 = stock_data['NET'][2]
    D20 = stock_data['PYPL'][2]; D21 = stock_data['QRVO'][2]
    D22 = stock_data["SQQQ"][2]; D23 = stock_data['PINS'][2]
    D24 = stock_data['EBAY'][2]; D25 = stock_data["CAT"][2]
    D26 = stock_data['DHI'][2]; D27 = stock_data["GOOGL"][2]
    D28 = stock_data["ZM"][2]; D29 = stock_data["DT"][2]
    D30 = stock_data["QFIN"][2]; D31 = stock_data["F"][2]
    D32 = stock_data["BA"][2]; D33 = stock_data["JETS"][2]
    #data for close 
    E0 = stock_data['AAPL'][3]; E1 = stock_data['AMZN'][3]
    E2 = stock_data['BABA'][3]; E3 = stock_data['GS'][3]
    E4 = stock_data['META'][3]; E5 = stock_data['JD'][3]
    E6 = stock_data['MRVL'][3]; E7 = stock_data['MSFT'][3]
    E8 = stock_data['NFLX'][3]; E9 = stock_data['NVDA'][3]
    E10 = stock_data['QQQ'][3]; E11 = stock_data['SNAP'][3]
    E12 = stock_data['SPY'][3]; E13 = stock_data['TSLA'][3]
    E14 = stock_data['WFC'][3]; E15 = stock_data['DKNG'][3]
    E16 = stock_data['NIO'][3]; E17 = stock_data['DIS'][3]
    E18 = stock_data['JNJ'][3]; E19 = stock_data['NET'][3]
    E20 = stock_data['PYPL'][3]; E21 = stock_data['QRVO'][3]
    E22 = stock_data["SQQQ"][3]; E23 = stock_data['PINS'][3]
    E24 = stock_data['EBAY'][3]; E25 = stock_data["CAT"][3]
    E26 = stock_data['DHI'][3]; E27 = stock_data["GOOGL"][3]
    E28 = stock_data["ZM"][3]; E29 = stock_data["DT"][3]
    E30 = stock_data["QFIN"][3]; E31 = stock_data["F"][3]
    E32 = stock_data["BA"][3]; E33 = stock_data["JETS"][3]
    #data for time
    F0 = stock_data['AAPL'][4]; F1 = stock_data['AMZN'][4]
    F2 = stock_data['BABA'][4]; F3 = stock_data['GS'][4]
    F4 = stock_data['META'][4]; F5 = stock_data['JD'][4]
    F6 = stock_data['MRVL'][4]; F7 = stock_data['MSFT'][4]
    F8 = stock_data['NFLX'][4]; F9 = stock_data['NVDA'][4]
    F10 = stock_data['QQQ'][4]; F11 = stock_data['SNAP'][4]
    F12 = stock_data['SPY'][4]; F13 = stock_data['TSLA'][4]
    F14 = stock_data['WFC'][4]; F15 = stock_data['DKNG'][4]
    F16 = stock_data['NIO'][4]; F17 = stock_data['DIS'][4]
    F18 = stock_data['JNJ'][4]; F19 = stock_data['NET'][4]
    F20 = stock_data['PYPL'][4]; F21 = stock_data['QRVO'][4]
    F22 = stock_data["SQQQ"][4]; F23 = stock_data['PINS'][4]
    F24 = stock_data['EBAY'][4]; F25 = stock_data["CAT"][4]
    F26 = stock_data['DHI'][4]; F27 = stock_data["GOOGL"][4]
    F28 = stock_data["ZM"][4]; F29 = stock_data["DT"][4]
    F30 = stock_data["QFIN"][4]; F31 = stock_data["F"][4]
    F32 = stock_data["BA"][4]; F33 = stock_data["JETS"][4]

    #data for classification
    G0 = stock_data['AAPL'][5]; G1 = stock_data['AMZN'][5]
    G2 = stock_data['BABA'][5]; G3 = stock_data['GS'][5]
    G4 = stock_data['META'][5]; G5 = stock_data['JD'][5]
    G6 = stock_data['MRVL'][5]; G7 = stock_data['MSFT'][5]
    G8 = stock_data['NFLX'][5]; G9 = stock_data['NVDA'][5]
    G10 = stock_data['QQQ'][5]; G11 = stock_data['SNAP'][5]
    G12 = stock_data['SPY'][5]; G13 = stock_data['TSLA'][5]
    G14 = stock_data['WFC'][5]; G15 = stock_data['DKNG'][5]
    G16 = stock_data['NIO'][5]; G17 = stock_data['DIS'][5]
    G18 = stock_data['JNJ'][5]; G19 = stock_data['NET'][5]
    G20 = stock_data['PYPL'][5]; G21 = stock_data['QRVO'][5]
    G22 = stock_data["SQQQ"][5]; G23 = stock_data['PINS'][5]
    G24 = stock_data['EBAY'][5]; G25 = stock_data["CAT"][5]
    G26 = stock_data['DHI'][5]; G27 = stock_data["GOOGL"][5]
    G28 = stock_data["ZM"][5]; G29 = stock_data["DT"][5]
    G30 = stock_data["QFIN"][5]; G31 = stock_data["F"][5]
    G32 = stock_data["BA"][5]; G33 = stock_data["JETS"][5]

    #data for pre_market

    H0 = stock_data['AAPL'][6]; H1 = stock_data['AMZN'][6]
    H2 = stock_data['BABA'][6]; H3 = stock_data['GS'][6]
    H4 = stock_data['META'][6]; H5 = stock_data['JD'][6]
    H6 = stock_data['MRVL'][6]; H7 = stock_data['MSFT'][6]
    H8 = stock_data['NFLX'][6]; H9 = stock_data['NVDA'][6]
    H10 = stock_data['QQQ'][6]; H11 = stock_data['SNAP'][6]
    H12 = stock_data['SPY'][6]; H13 = stock_data['TSLA'][6]
    H14 = stock_data['WFC'][6]; H15 = stock_data['DKNG'][6]
    H16 = stock_data['NIO'][6]; H17 = stock_data['DIS'][6]
    H18 = stock_data['JNJ'][6]; H19 = stock_data['NET'][6]
    H20 = stock_data['PYPL'][6]; H21 = stock_data['QRVO'][6]
    H22 = stock_data["SQQQ"][6]; H23 = stock_data['PINS'][6]
    H24 = stock_data['EBAY'][6]; H25 = stock_data["CAT"][6]
    H26 = stock_data['DHI'][6]; H27 = stock_data["GOOGL"][6]
    H28 = stock_data["ZM"][6]; H29 = stock_data["DT"][6]
    H30 = stock_data["QFIN"][6]; H31 = stock_data["F"][6]
    H32 = stock_data["BA"][6]; H33 = stock_data["JETS"][6]


    sheet_instance.batch_update([{'range': 'B2:AI2', 'values': [[B0, B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23, B24, B25, B26, B27, B28, B29, B30, B31, B32, B33]]},{'range': 'B3:AI3', 'values': [[C0, C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, C14, C15, C16, C17, C18, C19, C20, C21, C22, C23, C24, C25, C26, C27, C28, C29, C30, C31, C32, C33]]},{'range': 'B4:AI4', 'values': [[D0, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, D14, D15, D16, D17, D18, D19, D20, D21, D22, D23, D24, D25, D26, D27, D28, D29, D30, D31, D32, D33]]}, {'range': 'B5:AI5', 'values': [[E0, E1, E2, E3, E4, E5, E6, E7, E8, E9, E10, E11, E12, E13, E14, E15, E16, E17, E18, E19, E20, E21, E22, E23, E24, E25, E26, E27, E28, E29, E30, E31, E32, E33]]}, {'range': 'B6:AI6', 'values': [[F0, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16, F17, F18, F19, F20, F21, F22, F23, F24, F25, F26, F27, F28, F29, F30, F31, F32, F33]]},{'range': 'B7:AI7', 'values': [[G0, G1, G2, G3, G4, G5, G6, G7, G8, G9, G10, G11, G12, G13, G14, G15, G16, G17, G18, G19, G20, G21, G22, G23, G24, G25, G26, G27, G28, G29, G30, G31, G32, G33]]},{'range': 'B8:AI8', 'values': [[H0, H1, H2, H3, H4, H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23, H24, H25, H26, H27, H28, H29, H30, H31, H32, H33]]}], value_input_option = "USER_ENTERED")


    
    ################################################
################################################
def POLYGON_PULL(key, sym):
    url = "https://api.polygon.io/v2/snapshot/locale/us/markets/stocks/tickers/"+sym+"?apiKey=" + key

    r = requests.get(url)
    data = r.json()
    dataIN = data['ticker']['day']
    o = dataIN['o']
    c = dataIN['c']
    h = dataIN['h']
    l = dataIN['l']
    return (o, h, l, c, "")
################################################
def FMP_PULL(sym):
    key = "624b51b375a1fae6682dc955e6159fc1"
    url = "https://financialmodelingprep.com/api/v3/quote/" +sym+ ",FB?apikey=" + key
    context = ssl.create_default_context(cafile=certifi.where())
    response = urlopen(url, context = context)
    data = response.read().decode("utf-8")
    r = json.loads(data)  
    dataIN = r[0]
    o = dataIN['open']
   # print(o)
    c = dataIN['price']
    #print(c)
    h = dataIN['dayHigh']
    #print(h)
    l = dataIN['dayLow']
    #print(l)
    try:  
      e = dataIN['earningsAnnouncement'][0:10]
    except:
      e = ""
    #print(e)
    return (o, h, l, c, e)
#################################################
def write_stock_data(stocks):
    stock_data = {}
    for sym in stocks:
        tup = FMP_PULL(sym)
        stock_data[sym] = tup 
    return stock_data
################################################
def write_as_batch(stock_data):
    #stock_data is a dictionary with open, high, low, close and time updated
    sheet_instance = get_sheet_instance('StockData')
    #data for open
    B0 = float(stock_data['AAPL'][0]); B1 = float(stock_data['AMZN'][0])
    B2 = float(stock_data['BABA'][0]); B3 = float(stock_data['GS'][0])
    B4 = float(stock_data['META'][0]); B5 = float(stock_data['JD'][0])
    B6 = float(stock_data['MRVL'][0]); B7 = float(stock_data['MSFT'][0])
    B8 = float(stock_data['NFLX'][0]); B9 = float(stock_data['NVDA'][0])
    B10 = float(stock_data['QQQ'][0]); B11 = float(stock_data['SNAP'][0])
    B12 = float(stock_data['SPY'][0]); B13 = float(stock_data['TSLA'][0])
    B14 = float(stock_data['WFC'][0]); B15 = float(stock_data['DKNG'][0])
    B16 = float(stock_data['NIO'][0]); B17 = float(stock_data['DIS'][0])
    B18 = float(stock_data['JNJ'][0]); B19 = float(stock_data['NET'][0])
    B20 = float(stock_data['PYPL'][0]); B21 = float(stock_data['QRVO'][0])
    B22 = float(stock_data["SQQQ"][0]); B23 = float(stock_data['PINS'][0])
    B24 = float(stock_data['EBAY'][0]); B25 = float(stock_data["CAT"][0])
    B26 = float(stock_data['DHI'][0]); B27 = float(stock_data["GOOGL"][0])
    B28 = float(stock_data["ZM"][0]); B29 = float(stock_data["DT"][0])
    B30 = float(stock_data["QFIN"][0]); B31 = float(stock_data["F"][0])
    B32 = float(stock_data["BA"][0]); B33 = float(stock_data["JETS"][0])
    #data for high
    C0 = float(stock_data['AAPL'][1]); C1 = float(stock_data['AMZN'][1])
    C2 = float(stock_data['BABA'][1]); C3 = float(stock_data['GS'][1])
    C4 = float(stock_data['META'][1]); C5 = float(stock_data['JD'][1])
    C6 = float(stock_data['MRVL'][1]); C7 = float(stock_data['MSFT'][1])
    C8 = float(stock_data['NFLX'][1]); C9 = float(stock_data['NVDA'][1])
    C10 = float(stock_data['QQQ'][1]); C11 = float(stock_data['SNAP'][1])
    C12 = float(stock_data['SPY'][1]); C13 = float(stock_data['TSLA'][1])
    C14 = float(stock_data['WFC'][1]); C15 = float(stock_data['DKNG'][1])
    C16 = float(stock_data['NIO'][1]); C17 = float(stock_data['DIS'][1])
    C18 = float(stock_data['JNJ'][1]); C19 = float(stock_data['NET'][1])
    C20 = float(stock_data['PYPL'][1]); C21 = float(stock_data['QRVO'][1])
    C22 = float(stock_data["SQQQ"][1]); C23 = float(stock_data['PINS'][1])
    C24 = float(stock_data['EBAY'][1]); C25 = float(stock_data["CAT"][1])
    C26 = float(stock_data['DHI'][1]); C27 = float(stock_data["GOOGL"][1])
    C28 = float(stock_data["ZM"][1]); C29 = float(stock_data["DT"][1])
    C30 = float(stock_data["QFIN"][1]); C31 = float(stock_data["F"][1])
    C32 = float(stock_data["BA"][1]); C33 = float(stock_data["JETS"][1])
    #data for low 
    D0 = float(stock_data['AAPL'][2]); D1 = float(stock_data['AMZN'][2])
    D2 = float(stock_data['BABA'][2]); D3 = float(stock_data['GS'][2])
    D4 = float(stock_data['META'][2]); D5 = float(stock_data['JD'][2])
    D6 = float(stock_data['MRVL'][2]); D7 = float(stock_data['MSFT'][2])
    D8 = float(stock_data['NFLX'][2]); D9 = float(stock_data['NVDA'][2])
    D10 = float(stock_data['QQQ'][2]); D11 = float(stock_data['SNAP'][2])
    D12 = float(stock_data['SPY'][2]); D13 = float(stock_data['TSLA'][2])
    D14 = float(stock_data['WFC'][2]); D15 = float(stock_data['DKNG'][2])
    D16 = float(stock_data['NIO'][2]); D17 = float(stock_data['DIS'][2])
    D18 = float(stock_data['JNJ'][2]); D19 = float(stock_data['NET'][2])
    D20 = float(stock_data['PYPL'][2]); D21 = float(stock_data['QRVO'][2])
    D22 = float(stock_data["SQQQ"][2]); D23 = float(stock_data['PINS'][2])
    D24 = float(stock_data['EBAY'][2]); D25 = float(stock_data["CAT"][2])
    D26 = float(stock_data['DHI'][2]); D27 = float(stock_data["GOOGL"][2])
    D28 = float(stock_data["ZM"][2]); D29 = float(stock_data["DT"][2])
    D30 = float(stock_data["QFIN"][2]); D31 = float(stock_data["F"][2])
    D32 = float(stock_data["BA"][2]); D33 = float(stock_data["JETS"][2])
    #data for close 
    E0 = float(stock_data['AAPL'][3]); E1 = float(stock_data['AMZN'][3])
    E2 = float(stock_data['BABA'][3]); E3 = float(stock_data['GS'][3])
    E4 = float(stock_data['META'][3]); E5 = float(stock_data['JD'][3])
    E6 = float(stock_data['MRVL'][3]); E7 = float(stock_data['MSFT'][3])
    E8 = float(stock_data['NFLX'][3]); E9 = float(stock_data['NVDA'][3])
    E10 = float(stock_data['QQQ'][3]); E11 = float(stock_data['SNAP'][3])
    E12 = float(stock_data['SPY'][3]); E13 = float(stock_data['TSLA'][3])
    E14 = float(stock_data['WFC'][3]); E15 = float(stock_data['DKNG'][3])
    E16 = float(stock_data['NIO'][3]); E17 = float(stock_data['DIS'][3])
    E18 = float(stock_data['JNJ'][3]); E19 = float(stock_data['NET'][3])
    E20 = float(stock_data['PYPL'][3]); E21 = float(stock_data['QRVO'][3])
    E22 = float(stock_data["SQQQ"][3]); E23 = float(stock_data['PINS'][3])
    E24 = float(stock_data['EBAY'][3]); E25 = float(stock_data["CAT"][3])
    E26 = float(stock_data['DHI'][3]); E27 = float(stock_data["GOOGL"][3])
    E28 = float(stock_data["ZM"][3]); E29 = float(stock_data["DT"][3])
    E30 = float(stock_data["QFIN"][3]); E31 = float(stock_data["F"][3])
    E32 = float(stock_data["BA"][3]); E33 = float(stock_data["JETS"][3])
    #data for time
    F0 = stock_data['AAPL'][4]; F1 = stock_data['AMZN'][4]
    F2 = stock_data['BABA'][4]; F3 = stock_data['GS'][4]
    F4 = stock_data['META'][4]; F5 = stock_data['JD'][4]
    F6 = stock_data['MRVL'][4]; F7 = stock_data['MSFT'][4]
    F8 = stock_data['NFLX'][4]; F9 = stock_data['NVDA'][4]
    F10 = stock_data['QQQ'][4]; F11 = stock_data['SNAP'][4]
    F12 = stock_data['SPY'][4]; F13 = stock_data['TSLA'][4]
    F14 = stock_data['WFC'][4]; F15 = stock_data['DKNG'][4]
    F16 = stock_data['NIO'][4]; F17 = stock_data['DIS'][4]
    F18 = stock_data['JNJ'][4]; F19 = stock_data['NET'][4]
    F20 = stock_data['PYPL'][4]; F21 = stock_data['QRVO'][4]
    F22 = stock_data["SQQQ"][4]; F23 = stock_data['PINS'][4]
    F24 = stock_data['EBAY'][4]; F25 = stock_data["CAT"][4]
    F26 = stock_data['DHI'][4]; F27 = stock_data["GOOGL"][4]
    F28 = stock_data["ZM"][4]; F29 = stock_data["DT"][4]
    F30 = stock_data["QFIN"][4]; F31 = stock_data["F"][4]
    F32 = stock_data["BA"][4]; F33 = stock_data["JETS"][4]
    sheet_instance.batch_update([{'range': 'B2:AI2', 'values': [[B0, B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23, B24, B25, B26, B27, B28, B29, B30, B31, B32, B33]]},{'range': 'B3:AI3', 'values': [[C0, C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, C14, C15, C16, C17, C18, C19, C20, C21, C22, C23, C24, C25, C26, C27, C28, C29, C30, C31, C32, C33]]},{'range': 'B4:AI4', 'values': [[D0, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, D14, D15, D16, D17, D18, D19, D20, D21, D22, D23, D24, D25, D26, D27, D28, D29, D30, D31, D32, D33]]}, {'range': 'B5:AI5', 'values': [[E0, E1, E2, E3, E4, E5, E6, E7, E8, E9, E10, E11, E12, E13, E14, E15, E16, E17, E18, E19, E20, E21, E22, E23, E24, E25, E26, E27, E28, E29, E30, E31, E32, E33]]}, {'range': 'B6:AI6', 'values': [[F0, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16, F17, F18, F19, F20, F21, F22, F23, F24, F25, F26, F27, F28, F29, F30, F31, F32, F33]]}], value_input_option = "USER_ENTERED")

################################################
def good_morning(stocks):
    today = get_today()
    sheet_instance = get_sheet_instance('QFIN') 
    r = int(next_available_row(sheet_instance)) - 1  
    if sheet_instance.cell(r, 1).value == today: flag = True
    else: flag = False
    if flag == True:
        row_dict = {}
        for sym in stocks:
            row_dict = get_row_dict(sym, row_dict)
            print(sym, ' loaded')
            time.sleep(5)
    if flag == False:
        row_dict = {}
        for sym in stocks:
            row_dict = new_row(sym, row_dict)
            print("New Day Loaded for: " + sym)
            time.sleep(12) #need to check this timing
        time.sleep(100)
    return row_dict



#####################################################
def good_night(stocks, row_dict, name_tuples):
    time.sleep(60) #wait to make sure we are after 1600
    try:
        stock_data = write_stock_data(stocks)
        write_as_batch(stock_data) #run one more time 
        plus_minus()
        for sym in stocks:
            end_the_day(sym, row_dict)
            time.sleep(21)
    except Exception as e:
        PrintException()
        stock_data = write_stock_data(stocks)
        write_as_batch(stock_data) #run one more time 
        plus_minus()
        for sym in stocks:
            end_the_day(sym, row_dict)
            time.sleep(21)
    #check_for_resets('MoneyTree')
    #check_for_resets('MoneyTree2')
    
    #name_tuples = check_name(name_tuples)
    time.sleep(100)
    #earnings_dates(stocks)
################################################
def run_pre_market(stocks, hour, minute):
    if (hour == 7) or (hour == 8):
        try:
            pre_data = batch_pre(stocks)
            write_batch_pre(pre_data)
        except Exception as e: PrintException()
        time.sleep(21) 
    if (hour == 9) and (minute < 29):
        try:
            pre_data = batch_pre(stocks)
            write_batch_pre(pre_data)
        except Exception as e: PrintException()
        time.sleep(21)
    else: pass
################################################
def prepare_for_open(stocks, hour, minute, row_dict):
    if (hour == 9) and (minute == 35):
        try:
            for sym in stocks:
                pull_atr(sym, row_dict)
                time.sleep(20)
                #print(sym, ' is done!')
        except Exception as e:
            PrintException()
            for sym in stocks:
                pull_atr(sym, row_dict)
                time.sleep(20)
        time.sleep(40)
################################################
def run_open_market(stocks, hour, minute, name_tuples):
    if (hour == 9) and (minute >= 32):
        try:
            stock_data = write_stock_data(stocks)
            write_as_batch(stock_data)
            plus_minus()
        except EOFError as e1:
            print('EOF ERRORS: ', e1)
        except Exception as e:
            PrintException()
        time.sleep(20)
    if (hour >= 10) and (hour < 16):
        try:
            stock_data = write_stock_data(stocks)
            write_as_batch(stock_data)
            plus_minus()
        except EOFError as e1:
            print('HERE IS THE ERROR: ', e1)
        except Exception as e:
            print('No Its Here: ', e)
            PrintException()
        time.sleep(20)
    return name_tuples
################################################
#############  Main  ###########################
################################################
if __name__ == "__main__":
    stocks = ['AAPL', 'AMZN', 'BABA', 'GS', 'META', 'JD', 'MRVL', 'MSFT', 'NFLX', 'NVDA', 'QQQ', 'SNAP', 'SPY', 'TSLA', 'WFC', 'DKNG', 'NIO', 'DIS', "JNJ", 'NET', 'PYPL', 'QRVO', 'SQQQ', 'PINS', 'EBAY', 'CAT', 'DHI', 'GOOGL', 'ZM', 'DT', 'QFIN', 'F', 'BA', 'JETS']
    name_tuples = [('AAPL', ""), ('AMZN', ""), ('BABA', ""), ('GS', ""), ('META', ""), ('JD', ""), ('MRVL', ""), ('MSFT', ""), ('NFLX', ""), ('NVDA', ""), ('QQQ', ""), ('SNAP', ""), ('SPY', ""), ('TSLA', ""), ('WFC', ""), ('DKNG', ""), ('NIO', ""), ('DIS', ""), ('JNJ', ""), ('NET', ""), ('PYPL', ""), ('QRVO', ""), ('SQQQ', ""), ('PINS', ""), ('EBAY', ""), ("CAT", ""), ('DHI', ""), ('GOOGL', ""), ('ZM', ""), ('DT', ""), ('QFIN', ""), ('F', ""), ('BA', ""), ('JETS', "")]
    #################################################
    row_dict = good_morning(stocks)
    #################################################
    while True:
        hour, minute, second = time_is()
        run_pre_market(stocks, hour, minute)
        prepare_for_open(stocks, hour, minute, row_dict)
        name_tuples = run_open_market(stocks, hour, minute, name_tuples)
        if (hour >= 16) and (minute >= 18): break
    good_night(stocks, row_dict, name_tuples)
#################################################
#####         END                           #####
#################################################