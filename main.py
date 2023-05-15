from urllib.request import urlopen
from datetime import datetime
import certifi
import json
import ssl
import urllib.request

def Stock_LookUp(sym):
    key = "624b51b375a1fae6682dc955e6159fc1"
    url = "https://financialmodelingprep.com/api/v3/quote/" +sym+ ",FB?apikey=" + key
    context = ssl.create_default_context(cafile=certifi.where())
    response = urlopen(url, context = context)
    data = response.read().decode("utf-8")
    r = json.loads(data)  
    dataIN = r[0]
    return dataIN

def dif_YearlyHigh(sym):
    dataIN = Stock_LookUp(sym)
    yH = dataIN['yearHigh']
    c = dataIN['price']
    return yH - c

def dif_YearLow(sym):
  dataIN = Stock_LookUp(sym)
  lH = dataIN['yearLow']
  c = dataIN['price']
  return c - lH
  
while True:
  a = input("Enter symbl:")
  print()
  print(a + " is " + str(dif_YearlyHigh(a)) + " points away from the yearly high.")
  print(a + " is " + str(dif_YearLow(a)) + " points away from the yearly low.")
  print()