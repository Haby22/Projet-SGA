import pandas as pd
import sqlite3
from dash.dash_table.Format import Group
from dash import Dash,dcc, html,dash_table
import dash

from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from collections import OrderedDict
# csv file name



app = Dash(__name__)
#dash.Dash(
    #external_stylesheets=[dbc.themes.BOOTSTRAP])
#
df = pd.DataFrame(OrderedDict([
    ('climate', ['Sunny', 'Snowy', 'Sunny', 'Rainy']),
    ('temperature', [13, 43, 50, 30]),
    ('city', ['NYC', 'Montreal', 'Miami', 'NYC'])
]))

app.layout =  html.Div(children=[ html.Div(className='header',
                                   style={
                                            'background-image':  'url("/assets/junia.png")',#url="\img\shell.PNG",
                                            'background-size': '747px 337px',
                                            #'background-color': '#101b2b',
                                            'margin': '-8px -8px 0 -8px',
                                            'padding': '168px',
                                            #'color': 'white',
                                            'font-size': '25px',
                                            'text-align': 'center'
                                        }),
                                 html.Br(),
                                 html.Br(),
                                 html.Div(className='box-item',children=[
                                     
                                                html.Br(),
                                                html.H3('GESTION D\'ABSCENCE')
                                            ],style={
                                                'background-color': 'white',
                                                'margin': '10px 20px 25px 40px',
                                                #'padding': '10px',
                                                'width': '1200px',
                                                'height': '80px',
                                                'border-radius': '10px',
                                                'box-shadow': '12px 12px 12px #888888',
                                                'text-align': 'center',
                                                'textShadow':'2px 2px 8px #9DA2A4',
                                                }),
                                 html.Br(),html.P(),
                                 html.Div(),
                                 #html.Div(html.H3('GESTION D\'ABSCENCE',style={
                                           #     'text-align': 'center',
                                            #    'textShadow':'2px 2px 8px #9DA2A4' })),
                                 #html.Img(src="\img\shell.PNG",sizes="(max-width: 600px) 200px, 100vw"),
                            html.Div(className='data-display',children=[
                                            html.Div(className='data-item',id='graph',children=[
                                                dcc.Dropdown(
                                                    id='selection',
                                                    options=[
                                                        {'label':'Liste des eleves','value':'elev'},
                                                        {'label':'liste des abscences','value':'absent'}
                                                    ],
                                                    value='elev',
                                                    
                                                ),
                                                html.Br(),html.Br(),
                                                 ])]),
                                
                     html.Div(id='tra',children=[
                         dash_table.DataTable(id='table-dropdown',data=None,
                        columns=list(),editable=True,dropdown={'Etat_Abscence': {
                                'options': [{'label': 'Absent_Non_Justifie', 'value':'Absent_Non_Justifie'},
                                    {'label': 'Absent_Justifie', 'value':'Absent_Justifie'}]        
                                
                            }}
                    ),
                    #html.Div(id='table-dropdown-container')
                    ]),
                                                
   
                                                
            
                                               
                                                  
                                 #html.Div(id='live-update-text'),
                                 
                            #dcc.Graph(id='live-update-graph'),
                            dcc.Interval(
                                id='interval-component',
                                interval=1*50000, # in milliseconds
                                n_intervals=0)
   ])
#Output('live-update-text', 'children')
# Multiple components can update everytime interval gets fired.
@app.callback(Output('table-dropdown', 'columns'),Output('table-dropdown', 'data'),
              Input('interval-component', 'n_intervals'),
             Input('selection','value'),Input('table-dropdown','columns'),Input('table-dropdown','data'))
                                
def update_graph_live(n,selection,columns,data):
    filename = r"C:\Projet SGA\DB\BD.db"
    
    #df = pd.read_csv (filename , delimiter=';')
    

    conn = sqlite3.connect(filename)
    query = "SELECT * FROM Tableaura ;"

    df = pd.read_sql_query(query,conn)
    if selection == 'elev':
        
        
        columns=[]
        #dropdown=dict()
        for i in df.columns:

            if i == "Etat_Abscence":

                columns.append({'id': i, 'name': i, 'presentation': 'dropdown'})
            else:
                columns.append({'id': i, 'name': i})
        if data==None:
            data=df.to_dict('records')
        else:
            df=pd.DataFrame.from_dict(data)
            df.to_sql('Tableaura', conn, if_exists='replace', index = False)
            data=df.to_dict('records')
                        #editable=True,
                        
               
        #for i in range (len(df["Etat_Abscence"])):
            #df["Etat_Abscence"][i]=dcc.Dropdown(id='selection_1',options=[{'label':df["Etat_Abscence"][i],'value': df["Etat_Abscence"][i]},{'label':"Abscence justifiée",'value':"Abscence justifiée"}],value=df["Etat_Abscence"][i])
            #print(df["Etat_Abscence"][0].__dict__)
        
        #table = dbc.Table.from_dataframe(df, striped=True, bordered=True, hover=True)
        #print(table)
        #df.to_sql('Tableaura', conn, if_exists='replace', index = False)
        
        return columns,data

if __name__ == "__main__":
    
    app.run_server(debug=True)