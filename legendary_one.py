import requests
from bs4 import BeautifulSoup
import xlwings as xw
from datetime import datetime
import time

url = 'https://finance.yahoo.com'
wb = xw.Book()
# Connect to the active Excel application
sheet = wb.sheets.active
# Define our column headers and their colors
sheet.range('A1').value = 'Time'  # Header of A column
sheet.range('A:A').color = (0, 255, 0)  # Green for timestamp column
sheet.range('B1').value = 'S&P 500'  # Header of B column
sheet.range('B:B').color = (255, 0, 0)  # Red for S&P 500 price column
sheet.range('C1').value = 'DOW 30'  # Header of C column
sheet.range('C:C').color = (0, 0, 255)  # Blue for DOW30 price column
sheet.range('D1').value = 'NASDAQ'  # Header of D column
sheet.range('D:D').color = (255, 255, 0)  # Yellow for Nasdaq price column
sheet.range('E1').value = 'RUSSEL 2000'  # Header of D column
sheet.range('E:E').color = (255, 0, 255)  # Purple for Russel2000 price column




# Print the address of the last empty row



while True:
    response = requests.get(url)
    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    # scrap the S&P 500 price
    sp_price_element = soup.find('fin-streamer', {'data-symbol': 'ES=F', 'data-field': 'regularMarketPrice'})
    sp_price = sp_price_element['value']
    # scrap the DOW30 price
    dow30_price_element = soup.find('fin-streamer', {'data-symbol': 'YM=F', 'data-field': 'regularMarketPrice'})
    dow30_price = dow30_price_element['value']
    # scrap the nasdaq price
    nasdaq_price_element = soup.find('fin-streamer', {'data-symbol': 'NQ=F', 'data-field': 'regularMarketPrice'})
    nasdaq_price = nasdaq_price_element['value']
    # scrap the RUSSEL price
    russel_price_element = soup.find('fin-streamer', {'data-symbol': 'RTY=F', 'data-field': 'regularMarketPrice'})
    russel_price = russel_price_element['value']

    # Retrieve the current timestamp
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Write the index prices to specific cells
    sheet.range(f"A{sheet.cells.last_cell.row}").end('up').offset(1).value = timestamp
    sheet.range(f"B{sheet.cells.last_cell.row}").end('up').offset(1).value = sp_price
    sheet.range(f"C{sheet.cells.last_cell.row}").end('up').offset(1).value =dow30_price
    sheet.range(f"D{sheet.cells.last_cell.row}").end('up').offset(1).value = nasdaq_price
    sheet.range(f"E{sheet.cells.last_cell.row}").end('up').offset(1).value = russel_price
    # Save and close the workbook
    wb.save('index_prices.xlsx')
    time.sleep(300)


