##HAS TO BE ON PORT 9050### 


import os
import pandas as pd
import numpy as np
import matplotlib as plt
import io
import plotly.graph_objs as go
import plotly.express as px
from datetime import date, timedelta, datetime
import yfinance as yf


isin=pd.read_excel('C:/Users/molly/Desktop/cusip-isin.xlsx')

isin=isin.drop(columns=['ISIN','SEDOL','Cusip'])

account_dict={'APFC' : 'Alaska Permanent Fund',
              'URS' : 'Utah State Retirement Systems',
              'GOF' : 'CDAM GLOBAL OPPORTUNITIES FUND'}

accounts=['GOF','APFC','URS']
dfs_may={}
dfs_june={}
filtered_dfs={}
all_holdings={}
URS={}
APFC={}
GOF={}

January=[]
February=[]
March=[]
April=[]
May=[]
June=[]
July=[]
August=[]
September=[]
October=[]
November=[]
December=[]
trade_files={
   # 'January':'Trades 31.01.xlsx',
   # 'February' : 'Trades 29.02.xlsx',
   # 'March' :'Trades 31.03.xlsx',
    'April' : 'Trades 30.04.xlsx',
    'May' : 'Trades 31.05.xlsx',
    'June' : 'Trades 30.06.xlsx',
   # 'July' : 'Trades 31.07.xlsx',
   # 'August' : 'Trades 31.08.xlsx',
   # 'September' : 'Trades 30.09.xlsx',
   # 'October' : 'Trades 31.10.xlsx',
   # 'November' : 'Trades 30.11.xlsx',
   # 'December' : 'Trades 31.12.xlsx'
}
months= [##January, February,March,
    'April', 'May', 'June'
    ##,July, August, September, October, November, December
    ]

holdings_files= {
    ##'January': 'Holdings 31.01.xlsx', 
    ##'February': 'Holdings 29.02.xlx',
    ##'March' : 'Holdings 31.03.xlsx',
    'April' :'Holdings 30.04.xlsx', 
    'May' : 'Holdings 31.05.xlsx', 
    'June': 'Holdings 30.06.xlsx'
    ##,'July' : 'Holdings 31.07.xlsx', 'August' : 'Holdings 31.08.xlsx, 'September':'Holdings 30.09.xlsx',
    #  'October' : 'Holdings 31.10.xlsx, 'November':'Holdings 30.11.xlsx, 'December':'Holdings 31.12.xlsx'
}

##NEED TO JOIN HOLDINGS AND TRADING BEFORE MERGING TOGETHER##

##ALSO NEED TO JOIN WITH PRICING BEFORE ##


holdings=pd.DataFrame()
##holdings
for account in accounts:
    
    for month in months:
         tickers = yf.download('ATZ.TO ENGH.TO MTY.TO OTEX.TO BYD.TO TFII CIGI WDC SPGI SF.ST ZD QLYS NOW LGIH CNM ELV DAVA IP.MI DNP.WA IMCD.AS', period="1mo")
         Close=pd.DataFrame(tickers['Close'])

         holdings_file=holdings_files[month]
         file=(f'C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/{account}/{holdings_file}')
         try:
             fd=pd.read_excel(file)
             fd = fd.rename(columns={'Description':'Company Name'})
             fd['Trade Date']=pd.to_datetime(fd['Holding Scenario'].str[:10])
             fd=fd.loc[:,['Company Name', 
                             'Notional Quantity',
                             'Trade Date']]
             fd=fd[fd['Notional Quantity'] > 0]
             if fd.empty:
                 print("Empty")
             else:
                 holdings=pd.concat([holdings, fd], ignore_index=True)
         except FileNotFoundError:
             print(f"File not found: {file}")
         except Exception as e:
             print(f"Error processing file")
    all_trades = pd.DataFrame()
    ##Trade Data May
    for month in months:
        trade_file=trade_files[month]
        file=(f'C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/{trade_file}')
        try:
          df=pd.read_excel(file)
          df=df.rename(columns={'Description':'Company Name',
                                  'Notional Quantity' : 'Quantity'})
          le_name=account_dict[f'{account}']
          filtered= df[(df['LE Name']==le_name) & (df['ISIN'].str.len()>0)]
          df['Trade Date']=pd.to_datetime(df['Trade Date'])
          combined_filtered = pd.merge(left=filtered, right=isin, on='Company Name', how='left')
          combined_filtered['Ticker'] = combined_filtered['Ticker'].str.split().str[0]
          
          if combined_filtered.empty:
              print(f"Trade DataFrame for {account} in {month} is empty")
          else:
              grouped_df=combined_filtered.groupby(['ISIN','Trade Date', 'Company Name', 'Ticker']).agg(
                  {'Quantity':'sum'}).reset_index()
              grouped_df= grouped_df.sort_values(by=['ISIN', 'Trade Date'])
              grouped_df['Cumulative Quantity']=grouped_df.groupby(['ISIN', 'Company Name'])['Quantity'].cumsum()
              all_trades=pd.concat([all_trades, grouped_df], ignore_index=True) 
        except FileNotFoundError:
            print(f"File not found: {file}")
        except Exception as e:
            print(f"Error processing file {file}")
    if holdings.empty:
        print(f"No valid holdings data for account {account}")
    else:
        holdings_f = pd.merge(left=holdings, right=all_trades, on='Company Name', how='left')
        holdings_f['Cumulative Total'] = holdings_f['Notional Quantity'] + holdings_f['Quantity']
        holdings_f['Cumulative Total'].fillna(holdings_f['Notional Quantity'], inplace=True)
        holdings_f['Trade Date_y'].fillna(holdings_f['Trade Date_x'], inplace=True)
    ##create an equity dataset for each account
    equities = holdings_f['Ticker'].unique()
    print(equities)
    for ticker in equities: 
        equity_acc = holdings_f[holdings_f['Ticker'] == ticker]
        if account == 'URS' :
          URS[ticker]=equity_acc
        elif account == 'APFC':
            APFC[ticker]=equity_acc
        elif account == 'GOF':
            GOF[ticker]=equity_acc

print(URS['ZD'])
print(URS['IP'])
##TRANSFORM THE DICTIONARIES INTO EQUITY DATAFRAME WITH A COLUMN FOR EACH ACCOUNT 

print(type(GOF))
     

########
##open daily nav file
file_directory = "G:/OPS/Server Data/"
file = "NAV.xlsx"
excel = os.path.join(file_directory, file)
graph_data = pd.read_excel(excel)
search_file=date.today() - timedelta(days=1)
new= search_file.strftime("%Y-%m-%d") + '.xlsx'

##add new data to existing both CDAM NAV and MSCI
direct= 'C:\\Users\\molly\\Documents\\NAV_XL'
new=os.path.join(direct,new)
print(new)
if os.path.exists(new):
     excel = pd.read_excel(new)
     NAV2 = excel.iloc[31]
     NAV = graph_data._append(NAV2)
     wrld = yf.download('IWRD.L', period='1d')
     wrld.reset_index(inplace=True)
     wrld['MSCI'] = (wrld['Close'] - wrld['Close'].shift(1)) / wrld['Close'].shift(1) 
     wrld= wrld.drop(columns=['Open','High','Low','Close', 'Adj Close', 'Volume'])
     NAV = NAV.rename(columns={'Position Scenario Date Adjusted':'Date'})
     wrld = pd.merge(NAV, wrld, how='left', on='Date')
     wrld['MSCI_x'].fillna(wrld['MSCI_y'], inplace=True)
     wrld['Daily PNL (URS)'] = wrld['Utah Net MV(Dirty)'] -  wrld['Utah Net MV(Dirty)'].shift(1)
     wrld['Daily PNL (APFC)'] = wrld['Alaska Net MV(Dirty)'] - wrld['Alaska Net MV(Dirty)'].shift(1)
     wrld['Daily PNL (GOF)'] = wrld['CDAM Net MV(Dirty)'] - wrld['CDAM Net MV(Dirty)'].shift(1)
     wrld['Daily Returns (URS)'] = wrld['Daily PNL (URS)'] / wrld['Utah Net MV(Dirty)'].shift(1)
     wrld['Daily Returns (APFC)'] = wrld['Daily PNL (APFC)'] / wrld['Alaska Net MV(Dirty)'].shift(1)
     wrld['Daily Returns (GOF)'] = wrld['Daily PNL (GOF)'] / wrld['CDAM Net MV(Dirty)'].shift(1)
     wrld.to_excel('G:/OPS/Server Data/NAV_2.xlsx')
     ##os.remove(new)
else:
     print("No file")

##create graph dataframe
URS = {'Date': graph_data['Position Scenario Date Adjusted'],
       'MV' : graph_data['Utah Net MV(Dirty)'], 
       'Daily PNL' : graph_data['Daily PNL (URS)'], 
       'URS Return' : 100*(graph_data['Daily Return(URS)']),
       'APFC Return' : 100*(graph_data['Daily Returns (APFC)']),
       'GOF Return' : 100*graph_data['Daily Returns (GOF)'],
       'MSCI': (graph_data['MSCI'])
}
data = pd.DataFrame(URS)
dataset_column=data.columns[3:6]
end_date=data['Date'].max()
yesterday =pd.to_datetime((datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d'))
print(yesterday)
start_date=data['Date'].min()

eq_dict={'ATZ.TO' : 'ATZ',
'ENGH.TO' : 'ENGH',
'MTY.TO' : 'MTY',
'OTEX.TO' : 'OTEX',
'BYD.TO' : 'BYD',
'TFII' : 'TFII',
'CIGI' : 'CIGI',
'WDC' : 'WDC',
'ZD' : 'ZD',
'QLYS' : 'QLYS',
'NOW' : 'NOW',
'LGIH' : 'LGIH',
'CNM' : 'CNM',
'ELV' : 'ELV',
'DAVA' : 'DAVA',
'IP.MI' : 'IP',
'DNP.WA' : 'DNP',
'IMCD.AS' : 'IMCD'}



prices={}
##download equity info
tickers = yf.download('ATZ.TO ENGH.TO MTY.TO OTEX.TO BYD.TO TFII CIGI WDC SPGI SF.ST ZD QLYS NOW LGIH CNM ELV DAVA IP.MI DNP.WA IMCD.AS', period="1mo")

Close=pd.DataFrame(tickers['Close'])
Close.reset_index(inplace=True)
Close['Date']=pd.to_datetime(Close['Date'])
yesterday_data=Close[Close['Date']==yesterday]
print(yesterday_data)
pivoted_close=yesterday_data.melt(id_vars=['Date'], var_name='Ticker', value_name='Price')
pivoted_close['Ticker']=pivoted_close['Ticker'].replace(eq_dict)
print(pivoted_close)

ticker_column=Close.columns[1:19]
print(pivoted_close)

excel=pd.ExcelFile('C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/IRR.xlsx')

equities=excel.sheet_names
print(equities)


start_dates=[]
end_dates=[]
cum_quants=[]
cash_flows=[]
irrs=[]
grs=[]
return_p={}

import pyxirr
from pyxirr import xirr



for equity in equities:
    equity_df = pd.read_excel(excel, sheet_name=equity)
    dates=pd.to_datetime(equity_df['Trade Date'])
    not_quant_sum = equity_df['Notional Quantity'].sum()
    l_trade= (not_quant_sum) * (pivoted_close.loc[pivoted_close['Ticker']== equity, 'Price'])
    dates=equity_df['Trade Date'].apply(lambda x: x.toordinal() if pd.notnull(x) else None)
    values=equity_df['Trading Net Proceeds'].copy()
    values=values._append(pd.Series([l_trade]), ignore_index=True)
    end_date=yesterday.toordinal()
    dates=pd.concat([dates, pd.Series([end_date])], ignore_index=True)
    returns = pd.DataFrame({
         'Dates' :dates,
         'Values' : values
    })
    start_date=min(dates)
    period=(end_date - start_date)/365
    price_max= pivoted_close.loc[pivoted_close['Ticker']== equity, 'Price']
    price_min=equity_df.loc[equity_df['Trade Date']==equity_df['Trade Date'].min(), 'Trade Price'].values[0]
    try:
         irr=xirr(dates, values)
         CAGR= (price_max/price_min)**(1/period)-1
    except:
         print('Try Again')
         irr=2
         CAGR=1
    grs.append(CAGR)
    irrs.append(irr)


ir = pd.DataFrame({
     'Equities' : equities,
     'IRRS' : irrs,
     'CAGRS' : grs
})
print(ir)
print(type(ir))
##IR.to_excel('C:/Users/molly/Desktop/IRR.xlsx')



options_ds =[{'label' : col, 'value':col} for col in dataset_column]
options_eq=[{'label' : col, 'value' :col} for col in ticker_column]

month_options = [{'label':'January', 'value' :1}, {'label': 'February', 'value':2}, {'label':'March', 'value':3}, 
                 {'label' :'April', 'value': 4}, {'label':'May', 'value': 5}, {'label':'June', 'value' :6}, 
                 {'label' :'July', 'value': 7}, {'label' : 'August', 'value' :8}, {'label':'September', 'value':9},
                 {'label' :'October', 'value':10}, {'label':'November', 'value':11}, {'label':'December', 'value':12}]
year_options=[{'label':str(year), 'value':year} for year in range(2020, 2025)]


import dash
from dash import Dash, dcc, html, Input, Output, dash_table

app = Dash(__name__)
app.title = "CDAM Returns"


app.layout = html.Div(children=[
     html.Div(
          children=[
               html.Img(src='/assets/logo.png', style={'height' :'50px', 'float':'right', 'margin-right':'20px'}),
               html.H2('CDAM', style={'display':'inline','verticalAlign':'middle'})
          ]),

              html.Div(style={'width':'20%', 'padding' :'10px'},
                  children=[
                      html.H5('Account'),
                      dcc.Dropdown(
                         id='dataset-dropdown',
                         options=options_ds,
                         value=options_ds[0]['value'],
                         clearable=False
                ),
               

                         html.H5('Start Date'),
                         html.Div(
                               children=[
                                     dcc.Dropdown(
                              id='start-month-dropdown',
                              options=month_options,
                              value=month_options[0]['value'],
                              clearable=False
                         ),
                         dcc.Dropdown(
                              id='start-year-dropdown',
                              options=year_options,
                              value=year_options[0]['value'],
                              clearable=False
                         )  ]),
                         html.H5('End Date'),
                         html.Div(
                              children=[
                                   dcc.Dropdown(
                                        id='end-month-dropdown',
                                        options=month_options,
                                        value=month_options[11]['value'],
                                        clearable=False
                                   ),
                                   dcc.Dropdown(
                                        id='end-year-dropdown',
                                        options=year_options,
                                        value=year_options[-1]['value'],
                                        clearable=False
                                   )
                              ]
                         ),

                    html.H5('Equity'),
                    dcc.Dropdown(
                         id='equity-dropdown',
                         options=options_eq,
                         value=options_eq[0]['value'],
                         clearable=False
                    )
                  ]
                  
              ),
             html.Div(style={'width':'80%', 'padding' : '10px'},
               children=[
               dcc.Graph(id='returns-graph'),
               dcc.Graph(id='line-graph')
               ], 
          ),
          html.Link(rel='stylesheet', href="assets/styles_1.css")
])

@app.callback(
          Output('line-graph','figure'),
          [Input('equity-dropdown', 'value')]
        
)

def update_graph_line(selected_equity):
     if selected_equity in Close.columns:
          y_data=Close[selected_equity]
     else:
          print(f"Column '{selected_equity}' not in portfolio")
     line_fig=go.Figure()

     line_fig.add_trace(go.Scatter(x=Close['Date'], y=y_data, mode='lines', name=selected_equity))

     
     line_fig.update_layout(   
          title=f'{selected_equity} Close Price (1 Month)',
          xaxis={'title':'Date'},
          yaxis={'title':'Price'},
          colorway=['#7e9cbf']
          )
     return line_fig
     
@app.callback(
          Output('returns-graph', 'figure'),
          [
           Input('dataset-dropdown','value'),
           Input('start-month-dropdown', 'value'),
           Input('start-year-dropdown', 'value'),
           Input('end-month-dropdown', 'value'),
           Input('end-year-dropdown', 'value'),
         
        ]
)

def update_graph_bar(selected_dataset, start_month, start_year, end_month, end_year):
    start_date=pd.to_datetime(date(start_year, start_month, 1))
    end_date= pd.to_datetime(date(end_year, end_month, 1))
    data['Date'] = pd.to_datetime(data['Date'])
    filtered_data=data[(data['Date'] >= start_date) & (data['Date']<= end_date)]

    if selected_dataset in filtered_data.columns:
         column_data = filtered_data[selected_dataset]
    else:
         raise ValueError(f"Column '{selected_dataset}' does not exist in the Dataframe")

    selected_trace=go.Bar(
          x=filtered_data['Date'],
          y=column_data,
          name=selected_dataset,
          hovertemplate='%{y:.2f}%<extra></extra>'
        )
    msci_trace=go.Scatter(
        x=filtered_data['Date'],
        y=filtered_data['MSCI'],
        mode='lines+markers',
        name = 'MSCI', 
        hovertemplate='%{y:.2f}%<extra></extra>',
        marker=dict(color='blue')

    )

    return{
         'data' : [selected_trace, msci_trace],
         'layout' :go.Layout(
             title=f'{selected_dataset} Data with MSCI Benchmark',
             xaxis={'title':'Date'},
             yaxis={'title':'Returns (%)'},
             colorway=['#17b897'],
             hovermode='closest',
             barmode= 'group'
        )
    }
type(start_date)
type(data['Date'])
if __name__ == "__main__":
    app.run_server(host='192.168.3.38', port=9050 ,debug=True )


##html.Div(id='output-container-date-picker-range'),
##  html.Button('1 Month', id='btn-1-month', n_clicks=0),
##  html.Button('3 Months', id='btn-3-month', n_clicks=0),
#   html.Button('6 Months', id='btn-6-month', n_clicks=0),

##def update_date_range(n_clicks_1, c_clicks_3, n_clicks_6):
#  ctx=dash.callback_context
#  if not ctx.triggered:
#  button_id= 'btn-1-month'
#  else:
#   button_id=ctx.triggered[0]['prop_id'].split('.')[0]
#   if button_id=='btn-1-month':
#    return end_date - timedelta(days=30)
#  elif button_id=='btn-3-month':
#   return end_date - timedelta(days=90)
#  elif button_id=='btn-6-month':
#   return end_date - timedelta(days=180)

##html.Div(
          #          children=[
           #              dash_table.DataTable(
            #                  id='canada-table',
              #                columns=[{"name":i, "id":i} for i in canada.columns],
             #                 data=canada.to_dict('records'),
               #               )
                #              ]
              # ),