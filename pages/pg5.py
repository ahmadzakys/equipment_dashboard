######-----Import Dash-----#####
import dash
from dash import dcc
from dash import html, callback
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output

import pandas as pd
import numpy as np

from datetime import date
from dateutil.relativedelta import relativedelta

import plotly.graph_objects as go

dash.register_page(__name__, name='Sumba')

##-----page 5

## -----LAYOUT-----
layout = html.Div([
                ## --ROW1--
                html.Br(),
                html.P(children=[html.Strong('Bulk Sumba')],
                       style={'textAlign': 'center', 'fontSize': 27, 'background-color':'white','color':'#2a3f5f','font-family':'Verdana'}),

                dbc.Row([
                    dbc.Col([
                        dbc.Row([
                            dbc.Col([dcc.Graph(id="hou_c1_sumba")], xs=4),
                            dbc.Col([dcc.Graph(id="crane_1_sumba")], xs=8)
                        ]),

                        html.Br(),

                        dbc.Row([
                            dbc.Col([dcc.Graph(id="hou_c2_sumba")], xs=4),
                            dbc.Col([dcc.Graph(id="crane_2_sumba")], xs=8)
                        ]),
                    ], xs=5),

                    dbc.Col([
                        dbc.Row([
                            dcc.Graph(id="toh_sumba")
                        ]),

                        html.Br(),

                        dbc.Row([
                            dcc.Graph(id="goh_sumba")
                        ]),

                        html.Br(),

                        dbc.Row([
                            dcc.Graph(id="au_sumba")
                        ]),
                    ], xs=4),

                    dbc.Col([
                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Conveyor-Maker Check-Up')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="ConMakerShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="ConMakerCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Crane-Maker Check-Up')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="CraneMakerShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="CraneMakerCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Engine-Maker Check-Up')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="EngMakerShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="EngMakerCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Others-ShiELD')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="OthersShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="OthersCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(id="NextDockingDate_sumba",
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 85, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(className="bi bi-calendar-week", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="NextDockingCol_sumba",
                                color="light", outline=True, style={"height": 85, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Conveyor Belt Spares')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="ConSpareShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="ConSpareCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Crane-Wire Rope Spares')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(id="CraneSpareShape_sumba", style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                id="CraneSpareCol_sumba",
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),


                        ###### Bullet Graph
                        dbc.Row([
                            dbc.Col([
                            dbc.CardGroup([
                                dbc.Card(
                                    dbc.CardBody(
                                        [
                                            html.P(children=[html.Strong('Crane')],
                                                    style={'textAlign': 'left', 
                                                        'fontSize': 13, 
                                                        'background-color':'white',
                                                        'color':'#2a3f5f',
                                                        'font-family':'Verdana'}),
                                        ]
                                    ), color="light", outline=True, style={"height": 40, "margin-bottom": "0px"}
                                ),
                                dbc.Card(
                                    html.Div(id="CraneShape_sumba", style={
                                                                            "color": "white",
                                                                            "textAlign": "center",
                                                                            "fontSize": 17,
                                                                            "margin": "auto",
                                                                        }),
                                    id="CraneCol_sumba",
                                    color="light", outline=True, style={"height": 40, "margin-bottom": "1px", "maxWidth": 25}
                                ),
                            ],),
                            ], xs=4),

                            dbc.Col([
                                dbc.CardGroup([
                                    dbc.Card(
                                        dbc.CardBody(
                                            [
                                                html.P(children=[html.Strong('Conv.')],
                                                        style={'textAlign': 'left', 
                                                            'fontSize': 13, 
                                                            'background-color':'white',
                                                            'color':'#2a3f5f',
                                                            'font-family':'Verdana'}),
                                            ]
                                        ), color="light", outline=True, style={"height": 40, "margin-bottom": "0px"}
                                    ),
                                    dbc.Card(
                                        html.Div(id="ConShape_sumba", style={
                                                                                "color": "white",
                                                                                "textAlign": "center",
                                                                                "fontSize": 17,
                                                                                "margin": "auto",
                                                                            }),
                                        id="ConCol_sumba",
                                        color="light", outline=True, style={"height": 40, "margin-bottom": "1px", "maxWidth": 25}
                                    ),
                                ],),
                            ], xs=4),

                            dbc.Col([
                                dbc.CardGroup([
                                    dbc.Card(
                                        dbc.CardBody(
                                            [
                                                html.P(children=[html.Strong('Engine')],
                                                        style={'textAlign': 'left', 
                                                            'fontSize': 13, 
                                                            'background-color':'white',
                                                            'color':'#2a3f5f',
                                                            'font-family':'Verdana'}),
                                            ]
                                        ), color="light", outline=True, style={"height": 40, "margin-bottom": "0px"}
                                    ),
                                    dbc.Card(
                                        html.Div(id="EngShape_sumba", style={
                                                                                "color": "white",
                                                                                "textAlign": "center",
                                                                                "fontSize": 17,
                                                                                "margin": "auto",
                                                                            }),
                                        id="EngCol_sumba",
                                        color="light", outline=True, style={"height": 40, "margin-bottom": "1px", "maxWidth": 25}
                                    ),
                                ],),
                            ], xs=4),
                        ], className="g-0 d-flex align-items-center"),

                        dbc.Row([
                            dbc.Col([
                                dbc.Card([
                                    dbc.CardBody([])
                                ], color="light", outline=True, style={"height": 160, "margin-bottom": "5px"}),
                            ], xs=4),

                            dbc.Col([
                                dbc.Card([
                                    dbc.CardBody([dcc.Graph(id="cs_conveyor_sumba")])
                                ], color="light", outline=True, style={"height": 160, "margin-bottom": "5px"}),
                            ], xs=4, className="g-0 d-flex align-items-center"),

                            dbc.Col([
                                dbc.Card([
                                    dbc.CardBody([dcc.Graph(id="cs_engine_sumba")])
                                ], color="light", outline=True, style={"height": 160, "margin-bottom": "5px"}),
                            ], xs=4, className="g-0 d-flex align-items-center")
                        ], className="g-0 d-flex align-items-center"),
                    ], xs=3)
                ]),      

    html.Br(),
    html.Footer('ABL',
            style={'textAlign': 'center', 
                   'fontSize': 20, 
                   'background-color':'#2a3f5f',
                   'color':'white'})

    ], style={
        'paddingLeft':'10px',
        'paddingRight':'10px',
    })

#### Callback Auto Update Chart & Data

@callback(
    [Output('hou_c1_sumba', 'figure'),
     Output('hou_c2_sumba', 'figure'),
     Output('toh_sumba', 'figure'),
     Output('goh_sumba', 'figure'),
     Output('crane_1_sumba', 'figure'),
     Output('crane_2_sumba', 'figure'),
     Output('au_sumba', 'figure'),

     Output('ConMakerShape_sumba', 'className'),
     Output('ConMakerCol_sumba', 'className'),
     Output('CraneMakerShape_sumba', 'className'),
     Output('CraneMakerCol_sumba', 'className'),
     Output('EngMakerShape_sumba', 'className'),
     Output('EngMakerCol_sumba', 'className'),
     Output('OthersShape_sumba', 'className'),
     Output('OthersCol_sumba', 'className'),
     Output('NextDockingDate_sumba', 'children'),
     Output('NextDockingCol_sumba', 'className'),
     Output('ConSpareShape_sumba', 'className'),
     Output('ConSpareCol_sumba', 'className'),
     Output('CraneSpareShape_sumba', 'className'),
     Output('CraneSpareCol_sumba', 'className'),
     Output('CraneShape_sumba', 'className'),
     Output('CraneCol_sumba', 'className'),
     Output('ConShape_sumba', 'className'),
     Output('ConCol_sumba', 'className'),
     Output('EngShape_sumba', 'className'),
     Output('EngCol_sumba', 'className'),

     Output('cs_conveyor_sumba', 'figure'),
     Output('cs_engine_sumba', 'figure'),
     ],
    Input('store', 'data')
)
def update_charts(data):
    ######################
    # Pre Processing
    ######################
    df_sumba = pd.DataFrame(data['Bulk Sumba'])
    df_sumba = df_sumba.replace(r'^\s*$', np.nan, regex=True)

    df_sumba[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']] = \
    df_sumba[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']].fillna(date.today())
    df_sumba['Last Overhaul GEN01'] = pd.to_datetime(df_sumba['Last Overhaul GEN01'], dayfirst = True)
    df_sumba['Last Overhaul GEN02'] = pd.to_datetime(df_sumba['Last Overhaul GEN02'], dayfirst = True)
    df_sumba['Last Overhaul GEN03'] = pd.to_datetime(df_sumba['Last Overhaul GEN03'], dayfirst = True)
    df_sumba['Last Overhaul GEN04'] = pd.to_datetime(df_sumba['Last Overhaul GEN04'], dayfirst = True)
    df_sumba['Next Docking Intermediate Survey (IS)'] = pd.to_datetime(df_sumba['Next Docking Intermediate Survey (IS)'], dayfirst = True)


    # Hydrolic Oil Utilization
    last = df_sumba.iloc[-1]
    hou1 = (last['CRANE 1 (DAILY TANK HYDROLICK)']-last['Last Rpl RH CR01'])/6000
    hou1 = round(hou1*100,0)
    hou1 = min(hou1, 100)

    hou2 = (last['CRANE 2 (DAILY TANK HYDROLICK)']-last['Last Rpl RH CR02'])/6000
    hou2 = round(hou2*100,0)
    hou2 = min(hou2, 100)

    # Wire Rope
    wire_rope_c1 = [['HW RHOL', last['HW RHOL CR01'], last['CRANE 1 (DAILY TANK HYDROLICK)']],
                    ['HW LHOL', last['HW LHOL CR01'], last['CRANE 1 (DAILY TANK HYDROLICK)']],
                    ['CW RHOL', last['CW RHOL CR01'], last['CRANE 1 (DAILY TANK HYDROLICK)']],
                    ['CW LHOL', last['CW LHOL CR01'], last['CRANE 1 (DAILY TANK HYDROLICK)']]]
    wire_rope_c1 = pd.DataFrame(wire_rope_c1, columns=['Crane 1', 'Last Replacement', 'Current RH'])
    wire_rope_c1['Utilization RH'] = wire_rope_c1['Current RH']-wire_rope_c1['Last Replacement']
    wire_rope_c1['Utilization %'] = round(wire_rope_c1['Utilization RH']/4000*100,0)
    util1 = []
    for i in wire_rope_c1['Utilization %']:
        util1.append(min(i, 100))
    wire_rope_c1['Utilization'] = util1
    wire_rope_c1['Max'] = 100

    wire_rope_c2 = [['HW RHOL', last['HW RHOL CR02'], last['CRANE 2 (DAILY TANK HYDROLICK)']],
                    ['HW LHOL', last['HW LHOL CR02'], last['CRANE 2 (DAILY TANK HYDROLICK)']],
                    ['CW RHOL', last['CW RHOL CR02'], last['CRANE 2 (DAILY TANK HYDROLICK)']],
                    ['CW LHOL', last['CW LHOL CR02'], last['CRANE 2 (DAILY TANK HYDROLICK)']]]
    wire_rope_c2 = pd.DataFrame(wire_rope_c2, columns=['Crane 2', 'Last Replacement', 'Current RH'])
    wire_rope_c2['Utilization RH'] = wire_rope_c2['Current RH']-wire_rope_c2['Last Replacement']
    wire_rope_c2['Utilization %'] = round(wire_rope_c2['Utilization RH']/4000*100,0)
    util2 = []
    for i in wire_rope_c2['Utilization %']:
        util2.append(min(i, 100))
    wire_rope_c2['Utilization'] = util2
    wire_rope_c2['Max'] = 100

    # Engine GOH
    engine_goh = [['DG1', last['DG1'], last['Last TOH DG1'], last['Last GOH DG1']],
                ['DG2', last['DG2'], last['Last TOH DG2'], last['Last GOH DG2']],
                ['DG3', last['DG3'], last['Last TOH DG3'], last['Last GOH DG3']],
                ['EDG', last['EDG'], last['Last TOH EDG'], last['Last GOH EDG']]]
    engine_goh = pd.DataFrame(engine_goh, columns=['Sumba', 'Current HR', 'Last TOH', 'Last GOH'])
    engine_goh['Last TOH'] = engine_goh['Last TOH'].fillna(0)
    # engine_goh['Utilization TOH'] = engine_goh['Current HR']-engine_goh['Last TOH']
    # engine_goh['Utilization GOH'] = engine_goh['Current HR']-engine_goh['Last GOH']
    engine_goh['Hours till TOH'] = 12000-engine_goh['Last GOH']
    engine_goh['Hours till GOH'] = 24000-engine_goh['Last GOH']
    engine_goh['TOH %'] = round((1-engine_goh['Hours till TOH']/12000)*100,0)
    engine_goh['GOH %'] = round((1-engine_goh['Hours till GOH']/24000)*100,0)

    TOH = []
    for i in engine_goh['TOH %']:
        TOH.append(min(i, 100))
    engine_goh['TOH'] = TOH

    GOH = []
    for i in engine_goh['GOH %']:
        GOH.append(min(i, 100))
    engine_goh['GOH'] = GOH

    # Alternator OH
    alternator_oh = [['GEN01', last['Last Overhaul GEN01'],],
                    ['GEN02', last['Last Overhaul GEN02'],],
                    ['GEN03', last['Last Overhaul GEN03'],],
                    ['GEN04', last['Last Overhaul GEN04'],]]
    alternator_oh = pd.DataFrame(alternator_oh, columns=['Sumba', 'Last Overhaul (Date)',])

    last_oh = []
    for i in alternator_oh['Last Overhaul (Date)']:
        last_oh.append(i+relativedelta(months=60))
    alternator_oh['Next Overhaul (Date)'] = last_oh

    alternator_oh['Todays Date'] = date.today()
    alternator_oh['Todays Date'] = alternator_oh['Todays Date'].astype('datetime64[ns]')
    alternator_oh['Time to Go'] = (alternator_oh['Next Overhaul (Date)'] - alternator_oh['Todays Date']).dt.days.astype('int64')
    alternator_oh['Five Years in Days'] = 1827
    alternator_oh['Utilization %'] = round((alternator_oh['Five Years in Days']-alternator_oh['Time to Go'])/ \
                                    alternator_oh['Five Years in Days']*100,0)

    util = []
    for i in alternator_oh['Utilization %']:
        util.append(min(i, 100))
    alternator_oh['Utilization'] = util
    alternator_oh['Max'] = 100

    #Critical Spare
    conveyor = last['Conveyor']
    df_conveyor = pd.DataFrame({'names' : ['progress',' '],
                    'values' :  [conveyor, 100 - conveyor]})

    engine = last['Engine']
    df_engine = pd.DataFrame({'names' : ['progress',' '],
                    'values' :  [engine, 100 - engine]})

    hou_c1 = plot_hou1(hou1)
    hou_c2= plot_hou2(hou2)
    toh = plot_toh(engine_goh['TOH'])
    goh = plot_goh(engine_goh['GOH'])
    crane_1 = plot_wr1(wire_rope_c1['Utilization'], wire_rope_c1)
    crane_2 = plot_wr2(wire_rope_c2['Utilization'], wire_rope_c2)
    au = plot_aoh(alternator_oh['Utilization'], alternator_oh)

    ConMakerShape = shape_cu(last['Conveyor-Maker Check-Up'])
    ConMakerCol = col_cu(last['Conveyor-Maker Check-Up'])
    CraneMakerShape = shape_cu(last['Crane-Maker Check-Up'])
    CraneMakerCol = col_cu(last['Crane-Maker Check-Up'])
    EngMakerShape = shape_cu(last['Engine-Maker Check-Up'])
    EngMakerCol = col_cu(last['Engine-Maker Check-Up'])
    OthersShape = shape_cu(last['Others'])
    OthersCol = col_cu(last['Others'])
    NextDockingDate = html.Strong('Next Docking Intermediate Survey (IS): ' + str(last['Next Docking Intermediate Survey (IS)'].strftime("%d %b %Y")))
    NextDockingCol = col_day_delta((last['Next Docking Intermediate Survey (IS)'] - pd.Timestamp.today()).days)
    ConSpareShape = shape_cu(last['Conveyor Belt Spares'])
    ConSpareCol = col_cu(last['Conveyor Belt Spares'])
    CraneSpareShape = shape_cu(last['Crane-Wire Rope Spares']) 
    CraneSpareCol = col_cu(last['Crane-Wire Rope Spares'])
    CraneShape = shape_cs(last['Crane'])
    CraneCol = col_cs(last['Crane'])
    ConShape = shape_cs(last['Conveyor'])
    ConCol = col_cs(last['Conveyor'])
    EngShape = shape_cs(last['Engine'])
    EngCol = col_cs(last['Engine'])

    cs_conveyor = plot_cs(df_conveyor)
    cs_engine = plot_cs(df_engine)

    return hou_c1, hou_c2, toh, goh, crane_1, crane_2, au,\
            ConMakerShape, ConMakerCol, CraneMakerShape, CraneMakerCol, EngMakerShape, EngMakerCol, OthersShape, OthersCol,\
            NextDockingDate, NextDockingCol, ConSpareShape, ConSpareCol, CraneSpareShape, CraneSpareCol,\
            CraneShape, CraneCol, ConShape, ConCol, EngShape, EngCol,\
            cs_conveyor, cs_engine


######################
# Plot Function
######################

#-- 1. Plot Hydrolic Utilization
def plot_hou1(value):
    color=''
    if value<=(100/3) : color='#488f31'
    elif value<=(100/3*2) : color='#ffe600'
    elif value<=100 : color='#de425b'

    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = value,
        number = {'suffix': "%" },
        domain = {'x': [0.1, 0.9], 'y': [0.03, 0.5]},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color,
            'bar':{'color':color},
            }))
    fig.add_annotation(x=0.5, y=0.75, text='<b>Hydrolic Oil Utilization</b><br>Crane no 1<br>(Max 6,000hrs)', 
                       font=dict(size=14), showarrow=False)
    fig.update_layout({'height':315,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':30, 'l':3, 'r':3}, "autosize": True})
    return(fig)

def plot_hou2(value):
    color=''
    if value<=(100/3) : color='#488f31'
    elif value<=(100/3*2) : color='#ffe600'
    elif value<=100 : color='#de425b'

    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = value,
        number = {'suffix': "%" },
        domain = {'x': [0.1, 0.9], 'y': [0.03, 0.5]},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color,
            'bar':{'color':color},
            }))
    fig.add_annotation(x=0.5, y=0.75, text='<b>Hydrolic Oil Utilization</b><br>Crane no 2<br>(Max 6,000hrs)', 
                       font=dict(size=14), showarrow=False)
    fig.update_layout({'height':315,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':30, 'l':3, 'r':3}, "autosize": True})
    return(fig)

#-- 2. Plot TOH & GOH
def plot_toh(value):
    color=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')
    fig = go.Figure()
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[0],
        number = {'suffix': "%" },
        domain = {'x': [0.02, 0.22], 'y': [0, 0.5]},
        title = {'text': "DG01", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[0],
            'bar':{'color':color[0]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[1],
        number = {'suffix': "%" },
        domain = {'x': [0.27, 0.47], 'y': [0, 0.5]},
        title = {'text': "DG02", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[1],
            'bar':{'color':color[1]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[2],
        number = {'suffix': "%" },
        domain = {'x': [0.52, 0.72], 'y': [0, 0.5]},
        title = {'text': "DG03", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[2],
            'bar':{'color':color[2]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[3],
        number = {'suffix': "%" },
        domain = {'x': [0.77, 0.97], 'y': [0, 0.5]},
        title = {'text': "DG04", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[3],
            'bar':{'color':color[3]},
            }))
    fig.add_annotation(x=0.5, y=0.97, text='<b>Engine Overhaul</b><br>TOH to be conducted @12,000hrs<br>', 
                       font=dict(size=18), showarrow=False)
    fig.update_layout({'height':175,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':3, 'l':3, 'r':3}, "autosize": True})
    

    return(fig)

############################################################################################################################

def plot_goh(value):
    color=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')
    fig = go.Figure()
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[0],
        number = {'suffix': "%" },
        domain = {'x': [0.02, 0.22], 'y': [0, 0.5]},
        title = {'text': "DG01", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[0],
            'bar':{'color':color[0]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[1],
        number = {'suffix': "%" },
        domain = {'x': [0.27, 0.47], 'y': [0, 0.5]},
        title = {'text': "DG02", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[1],
            'bar':{'color':color[1]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[2],
        number = {'suffix': "%" },
        domain = {'x': [0.52, 0.72], 'y': [0, 0.5]},
        title = {'text': "DG03", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[2],
            'bar':{'color':color[2]},
            }))
    fig.add_trace(go.Indicator(
        mode = "gauge+number",
        value = value[3],
        number = {'suffix': "%" },
        domain = {'x': [0.77, 0.97], 'y': [0, 0.5]},
        title = {'text': "DG04", 'font': {'size': 15}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[3],
            'bar':{'color':color[3]},
            }))
    fig.add_annotation(x=0.5, y=0.97, text='<b>Engine Overhaul</b><br>GOH to be conducted @24,000hrs<br>', 
                       font=dict(size=18), showarrow=False)
    fig.update_layout({'height':175,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':3, 'l':3, 'r':3}, "autosize": True})
    

    return(fig)

#-- 3. Plot Wire Rope
def plot_wr1(value, wire_rope_c1):
    anchos=0.5
    color=[]
    position=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')
    
    for i in value:
        if i<=50 : position.append('outside')
        else : position.append('inside')

    fig = go.Figure()
    fig.add_trace(go.Bar(x = wire_rope_c1['Crane 1'],
                         y = wire_rope_c1['Max'],
                         marker_color='#ededed',
                         hovertemplate='%{y:y}%',
                         width = anchos, name = 'Crane 1'))
    fig.add_trace(go.Bar(x = wire_rope_c1['Crane 1'], 
                         y = value,
                         text=value.apply(lambda x: '{0:1.0f}%'.format(x)),
                         textposition=position,
                         textfont_size=17,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 1'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 1</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 20,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=15)))
    fig.update_layout({'height':315,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':50, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

############################################################################################################################

def plot_wr2(value, wire_rope_c2):
    anchos = 0.5
    color=[]
    position=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')
    
    for i in value:
        if i<=50 : position.append('outside')
        else : position.append('inside')

    fig = go.Figure()
    fig.add_trace(go.Bar(x = wire_rope_c2['Crane 2'],
                         y = wire_rope_c2['Max'],
                         marker_color='#ededed',
                         hovertemplate='%{y:y}%',
                         width = anchos, name = 'Crane 2'))
    fig.add_trace(go.Bar(x = wire_rope_c2['Crane 2'], 
                         y = value,
                         text=value.apply(lambda x: '{0:1.0f}%'.format(x)),
                         textposition=position,
                         textfont_size=17,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 2'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 2</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 20,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=15)))
    fig.update_layout({'height':315,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':50, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

#-- 5. Plot Alternator OH
def plot_aoh(value, alternator_oh):
    anchos = 0.5
    color=[]
    position=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')

    for i in value:
        if i<=5 : position.append('outside')
        else : position.append('inside')

    fig = go.Figure()
    fig.add_trace(go.Bar(x = alternator_oh['Max'][::-1],
                         y = alternator_oh['Sumba'][::-1],
                         marker_color='#ededed',
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.add_trace(go.Bar(x = value[::-1], 
                         y = alternator_oh['Sumba'][::-1],
                         text=value[::-1].apply(lambda x: '{0:1.0f}%'.format(x)),
                         textposition=position[::-1],
                         textfont_size=15,
                         marker_color =color[::-1],
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.update_xaxes(visible=False)
    fig.update_layout(title = "<b>Alternator Utilization</b> <br> To be cleaned and service every 5 years",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 20,
                      showlegend=False,
                      yaxis = dict(tickfont = dict(size=15)))
    fig.update_layout({'height':250,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':60,'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

#-- 6. Plot Critical Space
# plotly
def plot_cs(df):
    color=''
    if df['values'][0]<=(50) : color='#d9534f'
    elif df['values'][0]<=(80) : color='#f0ad4e'
    elif df['values'][0]<=100 : color='#02b875'

    fig = go.Figure(data=[go.Pie(labels=df['names'], values=df['values'], 
                                 hole=0.5,  
                                 textinfo='none',
                                 marker_colors=[color, '#ededed'])])
    fig.update_layout(title = "<b>Critical Spare</b>",
                      title_x=0.5,
                      title_font_size = 11,
                      showlegend=False,)
    fig.update_layout({'height':120,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':30, 'b':3, 'l':3, 'r':3}, "autosize": True})
    val=str(df['values'][0])
    fig.add_annotation(x=0.5, y=0.5, text=f'<b>{val}%</b>', 
                           font=dict(size=11), showarrow=False)
    return(fig)


def shape_cu(value):
    shape=''
    if value == 'Green': shape="bi bi-check-square"
    elif value == 'Yellow': shape="bi bi-x-square"
    elif value == 'Red': shape="bi bi-x-square"
    elif value == 'Gray': shape="bi bi-square"
    return(shape)

def col_cu(value):
    col=''
    if value == 'Green': col="bg-success"
    elif value == 'Yellow': col="bg-warning"
    elif value == 'Red': col="bg-danger"
    elif value == 'Gray': col="bg-secondary"
    return(col)

def shape_cs(value):
    shape=''
    if value<=(50) : shape="bi bi-x-square"
    elif value<=(80) : shape="bi bi-x-square"
    elif value<=100 : shape="bi bi-check-square"
    else : shape="bi bi-square"
    return(shape)

def col_cs(value):
    col=''
    if value<=(50) : col="bg-danger"
    elif value<=(80) : col="bg-warning"
    elif value<=100 : col="bg-success"
    else : col="bg-secondary"
    return(col)

def col_day_delta(value):
    col=''
    if value<=180 : col="bg-danger"
    elif value<=365 : col="bg-warning"
    else: col="bg-success"
    return(col)



