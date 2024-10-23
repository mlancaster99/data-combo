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
import plotly.express as px
from dateutil.relativedelta import relativedelta
from labelmap import LabelMap
import dash
from dash import Dash, dcc, html, Input, Output, dash_table
from dash_table.Format import Format, Scheme
import calendar

URS_Hold = 'C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/URS'
APFC_Hold = 'C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/APFC'
GOF_Hold = 'C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/GOF'
metrics_m = pd.read_excel('C:/Users/molly/Documents/Practise/Coding/App_development/Factsheet Data/Metrics.xlsx')
metrics_folder = 'C:/Users/molly/Documents/Practise/Coding/App_development/Factsheet Data'
metrics_excel = pd.read_excel(os.path.join(metrics_folder, 'Metrics MSCI.xlsx'))
fig = px.line()
sector_m = pd.read_excel(os.path.join(URS_Hold, 'Sector_m.xlsx'))
country_m = pd.read_excel(os.path.join(URS_Hold, 'Country_m.xlsx'))
market_m = pd.read_excel(os.path.join(URS_Hold, 'Market_m.xlsx'))
country_dict = pd.read_excel(os.path.join(URS_Hold, 'Country dictionary.xlsx'))
metrics_dicte = pd.read_excel(os.path.join(URS_Hold, 'Metrics Dict.xlsx'))
cap_dict = pd.read_excel(os.path.join(URS_Hold, 'Cap dictionary.xlsx'))
country_names_dict = dict(zip(country_dict['Country'], country_dict['Name']))
cap_names_dict = dict(zip(cap_dict['Market Cap'], cap_dict['Name']))
metrics_dict = dict(zip(metrics_dicte['Name'], metrics_dicte['Metric']))
#URS = pd.read_excel(os.path.join(URS_Hold, 'Equity_r.xlsx'))
attribution = pd.read_excel(os.path.join(URS_Hold, 'Equity_r.xlsx'))
APFC = pd.read_excel(os.path.join(APFC_Hold, 'Equity_r_A.xlsx'))
GOF = pd.read_excel(os.path.join(GOF_Hold, 'Equity_r_G.xlsx'))
dates = market_m['Date'].unique()
nav = 'G:/OPS/Server Data/NAV_2.xlsx'
nav2 = pd.read_excel(nav)
accounts  = ['APFC', 'URS', 'GOF']
accounts_dictionary = {
    'APFC': 'Alaska Net MV(Dirty)',
    'URS' : 'Utah Net MV(Dirty)',
    'GOF' : 'CDAM Net MV(Dirty)'
}
account_data = {
        'URS': attribution,
        'APFC': APFC,
        'GOF': GOF
    }
today = datetime.now()
yesterday=datetime.now() - timedelta(days=1, hours=0, minutes=0, seconds=0, microseconds=0)
if yesterday.weekday()>= 5:
    yesterday = yesterday - timedelta(days=yesterday.weekday() - 4)
last_month = today - relativedelta(months=1)
last = yesterday.year - 1
last_day = calendar.monthrange(last_month.year, last_month.month)[1]
mtd = pd.to_datetime(f'{last_month.year}' + f'-{last_month.month}' + f'-{last_day}')
last = today - relativedelta(years=1)
last_year = last.year
ytd = pd.to_datetime(f'{last_year}' + f'-12' + f'-31')

isin = pd.read_excel('C:/Users/molly/Desktop/cusip-isin.xlsx')
nams = pd.read_excel("C:/Users/molly/Documents/Practise/Coding/App_development/Holdings Data etc/ISIN Names.xlsx")
isin_names_dict = dict(zip(nams['ISIN'], nams['Name']))

isin=isin.drop(columns=['ISIN','SEDOL','Cusip'])
company_to_ticker = pd.Series(isin.Ticker.values, index=isin['Company Name']).to_dict()

metrics_options = [{'label': i, 'value': j} for i, j in metrics_dict.items()]

fig = px.area()
performance = pd.DataFrame()

external_stylesheet= [
    {
        "href" : (
            "https://fonts.googleapis.com/css2?"
            "family=Urbanist:ital,wght@0,100..900;1,100..900&display=swap"
            "family=Bodoni+Moda:wght@400;700&"
            ),
            "rel":"stylesheet",
        },
]
app = Dash(__name__, external_stylesheets=external_stylesheet)
app.title = "CDAM Returns"

app.layout = html.Div(
    children=[
        html.Div(
            children=[
                #html.P(src='/assets/logo.png', className="header-emoji"),
                html.H2('CDAM Portfolio', className="header-title")
                ], 
                className="header",
                ),
        html.Div(children=[
            html.H3(children="Attribution", className="menu-title"),
            dcc.Dropdown(
                    id='account-dropdown',
                    options = [
                    {'label': 'Account 1', 'value':'URS'},
                    {'label': 'Account 2', 'value':'APFC'},
                    {'label': 'Account 3', 'value':'GOF'}],
                    value='Account 1',  
                    clearable=False,  
                    className="dropdown",
                    style={'width': '250px'}),
                    dcc.Dropdown(
                id='date-dropdown',
                options=[],
                value = 'August 2024',
                className="dropdown",
                style={'width': '250px',
                'z-index':'1'}
                ),
            dcc.Dropdown(
                id='area-dropdown',
                options=[
                    {'label': 'Sector', 'value':'Sector'},
                    {'label': 'Country', 'value':'Country'},
                    {'label': 'Equity', 'value':'ISIN'},
                    {'label': 'Market Cap', 'value': 'Market Cap'}
                ] ,
                value='Sector',  # Set default value 
                className="dropdown",
                style={'width': '250px'}
                ),
                
            dcc.Dropdown(
                id = "time-dropdown",
                options=[
                    {'label':'1 Month', 'value': 1},
                    {'label':'3 Months', 'value': 2},
                    {'label':'6 Months', 'value': 5},
                    {'label' : 'Year To Date', 'value':10},
                    {'label':'1 Year', 'value': 11},
                    {'label':'3 Years', 'value': 35},
                    {'label':'5 Years', 'value': 59},
                    {'label':'Since Inception', 'value':300}],
                value=1,
                className="dropdown",
                style={'width': '250px'}
                ),
                
                ],className="menu",
                ),
        html.Br(),
     ##   html.Div(id='perf-container'),
        html.Br(),
        html.Div(id='attribution-table-container'),
        html.Br(),
        html.Div(
            children=[
                dcc.Dropdown(
                    id="metrics-drop",
                    options = metrics_options,
                           ## value='P_E',
                            className="dropdown",
                            clearable=False,
                            style={'width': '300px'}
                            ),
                            ],
                        className="menu",
                        ),
        html.Div(
            children=dcc.Graph(id='metrics-graph'
            ),
            className = "card",
                ),
        html.Div(
            children=[
                dcc.Dropdown(
                    id="data-drop",
                    options = [{"label": "Sector", "value" : 'Sector'},
                            {"label": "Country", "value" : 'Country'},
                            {"label": "Market Cap", "value" : 'Market Cap'}],
                            value='Sector',
                            className='dropdown',
                            style={'width': '300px'}
                            )
                            ],
                        className="menu",
                        
                        ),
                html.Div(
                    children=dcc.Graph(id='area-graph'
                    ),
              className="card",
        )
    ],
    className="wrapper",
)      


@app.callback(
    Output('date-dropdown', 'options'),
    Input('account-dropdown', 'value')
    )

def update_dropdowns(selected_account):
    if selected_account in account_data:
        dates = account_data[selected_account]['Date'].unique()
        dates = pd.to_datetime(dates)
        time = [dt.to_pydatetime() for dt in dates]
        time.sort(reverse = True)
        return [{'label': i.strftime("%B %Y"), 'value': i} for i in time]
    return []


##@app.callback(
##    Output('perf-container', 'children'),
##    Input('account-dropdown', 'value')
##)
#def performance_tables(selected_account):
#    column_b = accounts_dictionary[selected_account]
#    month = nav2.loc[nav2['Date']==mtd, column_b].values[0]
#    today = nav2.loc[nav2['Date']==yesterday, column_b].values[0]
#    year = nav2.loc[nav2['Date']==ytd, column_b].values[0]
#    if selected_account == 'URS':
#        red = nav2[nav2['Date']>=ytd]
#        redemptions_y = red['Redemption U'].sum()
 ##       redm = nav2[nav2['Date']>=mtd]
    #    redemptions_m = redm['Redemption U'].sum()
   # elif selected_account == 'APFC':
   #     red = nav2[nav2['Date']>=ytd]
   #     redemptions_y= red['Redemption A'].sum()
   #     redm = nav2[nav2['Date']>=mtd]
   #     redemptions_m = redm['Redemption A'].sum()
    #else:
   #     red = nav2[nav2['Date']>= ytd]
    #    redemptions_y =  red['Redemption C'].sum()
   #     redm = nav2[nav2['Date']>=mtd]
   #     redemptions_m = redm['Redemption C'].sum()
   # year_td = 100*(today - year + redemptions_y)/ year
  #  month_td = 100*(today - month + redemptions_m)/ month
  #  row = {'Account': selected_account, 'MTD': month_td, 'YTD': year_td}
  #  performance._append(row, ignore_index=True)
  #  data= performance.to_dict('records')
   # return html.Div(
  #      children=[
  #          dash_table.DataTable(
  #              id='perf-container',
  #              data=data,
  #              columns=[
  #                  {'name': 'Account', 'id': 'Account'},
  #                  {'name':'Year to Date Return (%)', 'id':'YTD',
  #                  "type": "numeric", "format": Format(precision=2, scheme=Scheme.fixed)},
  #                  {'name': 'Month to Date Return (%)', 'id': 'MTD',
  #                  "type": "numeric", "format" : Format(precision=2, scheme=Scheme.fixed)}
  #              ],
  #              )
  #      ]
  #  )


@app.callback(
    Output('attribution-table-container', 'children'),
    [Input('date-dropdown', 'value'),
    Input('area-dropdown', 'value'),
    Input('time-dropdown', 'value'),
    Input('account-dropdown', 'value')]
)


def update_attr(selected_date, selected_area, selected_time, selected_account):
    if selected_account in account_data:
        attrib = account_data[selected_account]
        end = pd.to_datetime(selected_date)
        print(end)
        if pd.notnull(end) : 
            end_year = pd.to_datetime(end).year
        else:
            print('No year given')
        if selected_time == 10:
            end = end.strftime('%d/%m/%Y')
            new_file = attrib[(attrib['Date'] >= f'31/01/{end_year}') & (attrib['Date'] <= end)]
        elif selected_time in (2, 5, 11, 35, 59): 
            start = end - relativedelta(months=selected_time)
            new_file = attrib[(attrib['Date'] >= start) & (attrib['Date'] <= end)]
        elif selected_time == 300:
            start = attrib['Date'].min()
            new_file = attrib[(attrib['Date'] >= start) & (attrib['Date'] <= end)]
        else: 
            new_file = attrib[attrib['Date'] == end]
        values = new_file[selected_area].unique()
        df_a = pd.DataFrame()
        for value in values:
            value_df = new_file[new_file[selected_area] == value]
            equities = value_df['ISIN'].unique()
            pe= pd.DataFrame()
            for equity in equities:
                equit = value_df[value_df['ISIN']==equity]
                eq_cont = (equit['Contribution to Return'] +1).prod()-1
                con_row = {selected_area: value, 'Contrib':eq_cont}
                pe = pe._append(con_row, ignore_index=True)
            contrib = 100 * (round(pe['Contrib'].sum(), 4))
            row = {selected_area: value, 'Contribution to Return': contrib}
            df_a = df_a._append(row, ignore_index=True)
        try:
            if selected_area in ['Country','Market Cap']:
                df_a = df_a[df_a[selected_area]!= 0 & (df_a[selected_area]!='CASH & HEDGING')]
            else:
                df_a = df_a[df_a[selected_area] != 'CASH & HEDGING']

            positive = df_a[df_a['Contribution to Return'] >= 0].sort_values(by='Contribution to Return', ascending=False)
            negative = df_a[df_a['Contribution to Return'] < 0].sort_values(by = 'Contribution to Return', ascending=False)
            contributn = pd.concat([positive, negative], ignore_index=True)
            w_tot = contributn['Contribution to Return'].sum()
            n_row = {selected_area: 'Total', 'Contribution to Return':w_tot}
            contribution = contributn._append(n_row, ignore_index=True)
            if selected_area == 'ISIN': 
                    contribution[selected_area] = contribution[selected_area].map(isin_names_dict)
            elif selected_area == 'Country': 
                contribution[selected_area] = contribution[selected_area].map(country_names_dict)
            elif selected_area == 'Market Cap':
                contribution[selected_area] = contribution[selected_area].map(cap_names_dict)
            else:
                print('No Mapping')
            TableData = contribution.to_dict('records')
            start = new_file['Date'].min().strftime('%Y-%m-%d')
            start = start[:-2] + "01"
            end= new_file['Date'].max().strftime('%Y-%m-%d')
            if selected_time ==1:
                sentence = f'Attribution by {selected_area} for '+end
            else:
                sentence = f'Attribution by {selected_area} from '+start+' to '+end
            return html.Div([
                html.H3(children=sentence, className="menu-title"),
                dash_table.DataTable(
                id='attribution-table',
                data=TableData,
                columns=[
                    {'name': selected_area, 'id': selected_area},
                    {'name':'Contribution to Return (%)', 'id':'Contribution to Return',
                    "type": "numeric", "format": Format(precision=2, scheme=Scheme.fixed)}
                ],
            style_header={
                            'fontFamily':'Bodoni',
                            'fontWeight' : 'bold',
                            'color': '#002060',
                            'backgroundColor' : '#A8986E'
                        },
                        style_data={
                            'fontFamily':'Urbanist',
                            'color': '#002060',
                            'backgroundColor' : '#A8986E'
                        } 
            )])
        except:
            print('No date yet')


@app.callback(
    Output('metrics-graph', 'figure'),
    Input('metrics-drop', 'value')
)

def update_metrics(selected_metric):
    color_scheme = ['#1B305C', '#C5BFA5', '#1B5C30', '#4A6572', '#8A94A6', '#D9CDBF', '#94A692', '#354B5E', '#97A6B1']
    metrics = metrics_excel[~metrics_excel['Account'].isin(['Portfolio Compounder', 'FTSE Global', 'MSCI Small'])]
    text= next((option['label'] for option in metrics_options if option['value'] == selected_metric), None)
    fig=px.line()
    fig=px.line(
        metrics, x='Date',
        y= selected_metric, 
        color='Account',
        title= text,
        labels={
            selected_metric: text,
            "Date" : "Date"
            },
        color_discrete_sequence=color_scheme)
    fig.update_layout(
            title={
                'text': text,
                'font' : {
                    'family' : 'Bodoni',
                    'size':24
                }
            },
            font={
                'family':'Urbanist',
                'size':14
            },
            plot_bgcolor='#EBECEF',  # Background color of the plotting area
            paper_bgcolor='#D1D1D5'
        )  
    return fig

@app.callback(
    Output('area-graph', 'figure'),
     Input('data-drop', 'value')
)

def update_graph(selected_dataset):
    color_scheme = ['#1B305C', '#C5BFA5', '#1B5C30', '#4A6572', '#8A94A6', '#D9CDBF', '#94A692', '#354B5E', '#97A6B1']
    fig=px.area()
    if selected_dataset == 'Country':
        country_data = country_m.copy()
        label_map = {
            "CA" : "Canada",
            "US" : "United States",
            "UK" : "United Kingdom",
            "NL" : "Netherlands",
            "PL" : "Poland",
            "IT" : "Italy",
            "SS" : "Sweden",
            "FR" : "France"
        }
        country_data['Country Label'] = country_data['Country'].map(label_map)
        fig=px.area(
            country_data, x='Date', y='Total Weight', color='Country Label',
             title=f"Weight by {selected_dataset} ",
             labels={
                 "Value":"Weight (%)"
             },
             color_discrete_sequence=color_scheme
            )   
        fig.update_layout(
            title={
                'text':f"Weight by {selected_dataset} ",
                'font' : {
                    'family' : 'Bodoni',
                    'size':24
                }
            },
            font={
                'family':'Urbanist',
                'size':14
            },
            plot_bgcolor='#EBECEF',  # Background color of the plotting area
            paper_bgcolor='#D1D1D5'
        )     
    elif selected_dataset == 'Market Cap':
        market_data=market_m[market_m['Market Cap']!= 0 ]
        fig=px.area(
            market_data, x='Date', y='Total Weight', color='Market Cap',
             title=f"Weight by {selected_dataset} ",
             labels={
                 "Value":"Weight (%)"
             },
            
             color_discrete_sequence=color_scheme
            )   
        fig.update_layout(
            title={
                'text':f"Weight by {selected_dataset} ",
                'font' : {
                    'family' : 'Bodoni',
                    'size':24
                }
            },
            font={
                'family':'Urbanist',
                'size':14
            },
            plot_bgcolor='#EBECEF',  # Background color of the plotting area
            paper_bgcolor='#D1D1D5'
        )     
    elif selected_dataset == 'Sector':
        sector_data = sector_m.copy()
        fig=px.area(
            sector_data, x='Date', y='Total Weight', color='Sector',
             title=f"Weight by {selected_dataset} ",
             labels={
                 "Value":"Weight (%)"
             },
             color_discrete_sequence=color_scheme
            )
        fig.update_layout(
            title={
                'text':f"Weight by {selected_dataset} ",
                'font' : {
                    'family' : 'Bodoni',
                    'size':24
                }
            },
            font={
                'family':'Urbanist',
                'size':14
            },
            plot_bgcolor='#EBECEF',  # Background color of the plotting area
            paper_bgcolor='#D1D1D5'
        )
    return fig


if __name__ == "__main__":
   ## context=('G:\OPS\Server Data\keys\cert.pem', 'G:\OPS\Server Data\keys\key.pem')
    app.run_server(host='192.168.3.38', port=9050 , debug=True  # , ssl_context=context
    )
