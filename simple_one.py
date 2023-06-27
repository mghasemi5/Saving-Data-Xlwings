import requests
from bs4 import BeautifulSoup
import xlwings as xw


url = 'https://finance.yahoo.com'

response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, 'html.parser')
#scrap the S&P 500 price
sp_price_element = soup.find('fin-streamer', {'data-symbol': 'ES=F', 'data-field': 'regularMarketPrice'})
sp_price = sp_price_element['value']
#scrap the DOW30 price
dow30_price_element = soup.find('fin-streamer', {'data-symbol': 'YM=F', 'data-field': 'regularMarketPrice'})
dow30_price = dow30_price_element['value']
#scrap the nasdaq price
nasdaq_price_element = soup.find('fin-streamer', {'data-symbol': 'NQ=F', 'data-field': 'regularMarketPrice'})
nasdaq_price = nasdaq_price_element['value']
#scrap the RUSSEL price
russel_price_element = soup.find('fin-streamer', {'data-symbol': 'RTY=F', 'data-field': 'regularMarketPrice'})
russel_price = russel_price_element['value']

# Connect to the active Excel application
wb = xw.Book()
sheet = wb.sheets.active

# Write the index prices to specific cells
sheet.range('A1').value = 'S&P 500'
sheet.range('A2').value = sp_price
sheet.range('B1').value = 'DOW30'
sheet.range('B2').value = dow30_price
sheet.range('C1').value = 'Nasdaq'
sheet.range('C2').value = nasdaq_price
sheet.range('D1').value = 'Russel2000'
sheet.range('D2').value = russel_price

# Save and close the workbook
wb.save('index_prices.xlsx')
wb.close()