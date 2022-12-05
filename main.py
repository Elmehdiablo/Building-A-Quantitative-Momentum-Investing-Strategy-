import pandas as pd 
import xlsxwriter
from scipy.stats import percentileofscore as score
from statistics import mean
import math
import requests
from sec import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv("S&P500.csv")
lst_stocks = []
my_column = [
    'Stock',
    'Price',
    'One-year price return',
    'One-year momentum return',
    '6 months price return',
    '6 months momentum return',
    '3 months price return',
    '3 months momentum return',
    '1 month price return',
    '1 month momentum return',
    'HQM SCORE',
    'Number of shares to Buy'

]
hqm_dataframe = pd.DataFrame(columns=my_column)
for stock in stocks['Symbol']:
    lst_stocks.append(stock)

def chunk(lst,n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]
lst_of_lst_stocks = list(chunk(lst_stocks, 100))   
lst_of_lst_stocks.pop()     

strings_stocks_lst = []
for j in lst_of_lst_stocks:
    strings_stocks_lst.append(','.join(j))
for item in strings_stocks_lst:
    batch_url_call = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={item}&types=quote,stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_url_call).json()
    for stck in item.split(","):
        try:
         hqm_dataframe = hqm_dataframe.append(
            pd.Series([
                stck,
                data[stck]['quote']['latestPrice'],
                data[stck]['stats']['year1ChangePercent']/100,
                'N/A',
                data[stck]['stats']['month6ChangePercent']/100,
                'N/A',
                data[stck]['stats']['month3ChangePercent']/100,
                'N/A',
                data[stck]['stats']['month1ChangePercent']/100,
                'N/A',
                'N/A',
                'N/A'

             ],index=my_column
             ),ignore_index=True
          )
        except KeyError:
            pass
time_periods = ['One-year','6 months','3 months','1 month']
for row in hqm_dataframe.index :
    for time_period in time_periods : 
        hqm_dataframe.loc[row,f'{time_period} momentum return'] = score(hqm_dataframe[f'{time_period} price return'],hqm_dataframe.loc[row,f'{time_period} price return'] )/100
for row in hqm_dataframe.index:
    momentum_percent = []
    for time_period in time_periods:
        momentum_percent.append(hqm_dataframe.loc[row,f'{time_period} momentum return'])
    hqm_dataframe.loc[row,'HQM SCORE']   = mean(momentum_percent) /100
hqm_dataframe.sort_values('HQM SCORE', ascending= False , inplace = True )
        

porfolio_size = input("Hello ! please enter your porfolio size : ")
try:
    porfolio = float(porfolio_size)
except ValueError:
    print("Please enter a number . \n")
    porfolio_size = input("Hello ! please enter your porfolio size : ")
    porfolio = float(porfolio)
#money_youhave()
hqm_dataframe = hqm_dataframe[:50]
hqm_dataframe.reset_index(drop=True , inplace=True)
position_size= porfolio/len(hqm_dataframe.index)
for i in hqm_dataframe.index:
    hqm_dataframe.loc[i,'Number of shares to Buy'] = math.floor(position_size/hqm_dataframe.loc[i,'Price'])
writer = pd.ExcelWriter('anas_investing.xlsx')    
hqm_dataframe.to_excel(writer,'Momentm Trades',index = False)

front_color = '#ffffff'
background_color = '#0a0a23'

string_format = writer.book.add_format({
    'font_color' : front_color,
    'bg_color'   : background_color,
     'border' : 1
})
dollar_format = writer.book.add_format({
    'num_format' : '$0.00',
    'font_color' : front_color,
    'bg_color'   : background_color,
    'border'     : 1 
})

percent_format  = writer.book.add_format({
    'num_format' : '0.0%',
    'font_color' : front_color,
    'bg_color'   : background_color,
    'border'     : 1 
})

integer_format = writer.book.add_format({
    'num_format' : '0',
    'font_color' :  front_color,
    'bg_color'   : background_color,
    'border'     : 1
})

column_template ={
    'A':['Stock',string_format],
    'B':['Price',dollar_format],
    'C':['One-year price return',percent_format ],
    'D': ['One-year momentum return',percent_format ],
    'E': ['6 months price return',percent_format ],
    'F': ['6 months momentum return',percent_format ],
    'G': ['3 months price return',percent_format ],
    'H': ['3 months momentum return',percent_format ],
    'I': ['1 month price return',percent_format ],
    'J': ['1 month momentum return',percent_format ],
    'K':['HQM SCORE',percent_format ],
    'L':['Number of shares to Buy',integer_format] 
   
}
for colum in column_template.keys():
    writer.sheets['Momentm Trades'].set_column(f'{colum}:{colum}',22,column_template[colum][1])
    writer.sheets['Momentm Trades'].write(f'{colum}1',column_template[colum][0],column_template[colum][1])
writer.save()    

