######-----Import Dash-----#####
import dash
from dash import dcc
from dash import html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output

import pandas as pd
import numpy as np

from datetime import date
from dateutil.relativedelta import relativedelta

from data import df_natuna

import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

dash.register_page(__name__, name='Natuna')

##-----page 5 data

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
    fig.add_annotation(x=0.5, y=0.5, text='<b>Hydrolic Oil Utilization</b><br>Crane no 1<br>(Max 6,000hrs)', 
                       font=dict(size=14), showarrow=False)
    fig.update_layout({'height':650,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':150, 'l':3, 'r':3}, "autosize": True})
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
def plot_wr1(value):
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
    fig.update_layout({'height':650,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':50, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)


#-- 5. Plot Alternator OH
def plot_aoh(value):
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
                         y = alternator_oh['Natuna'][::-1],
                         marker_color='#ededed',
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.add_trace(go.Bar(x = value[::-1], 
                         y = alternator_oh['Natuna'][::-1],
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


######################
# Pre Processing
######################

# Hydrolic Oil Utilization
last = df_natuna.iloc[-1]
hou1 = (last['CRANE']-last['Last Rpl RH CR01'])/6000
hou1 = round(hou1*100,0)
hou1 = min(hou1, 100)

# Wire Rope
wire_rope_c1 = [['HW RHOL', last['HW RHOL CR01'], last['CRANE']],
                ['HW LHOL', last['HW LHOL CR01'], last['CRANE']],
                ['CW RHOL', last['CW RHOL CR01'], last['CRANE']],
                ['CW LHOL', last['CW LHOL CR01'], last['CRANE']],
                ['EW RHOL', last['EW RHOL CR01'], last['CRANE']],]
wire_rope_c1 = pd.DataFrame(wire_rope_c1, columns=['Crane 1', 'Last Replacement', 'Current RH'])
wire_rope_c1['Utilization RH'] = wire_rope_c1['Current RH']-wire_rope_c1['Last Replacement']
wire_rope_c1['Utilization %'] = [round(wire_rope_c1['Utilization RH'][0]/4000*100,0),
                                round(wire_rope_c1['Utilization RH'][1]/4000*100,0),
                                round(wire_rope_c1['Utilization RH'][2]/4000*100,0),
                                round(wire_rope_c1['Utilization RH'][3]/4000*100,0),
                                round(wire_rope_c1['Utilization RH'][4]/2000*100,0),]
util1 = []
for i in wire_rope_c1['Utilization %']:
    util1.append(min(i, 100))
wire_rope_c1['Utilization'] = util1
wire_rope_c1['Max'] = 100


# Engine GOH
engine_goh = [['DG1', last['HARBOR GEN 1'], last['Last TOH DG1'], last['Last GOH DG1']],
              ['DG2', last['HARBOR GEN 2'], last['Last TOH DG2'], last['Last GOH DG2']],
              ['DG3', last['MAIN GENERATOR 3'], last['Last TOH DG3'], last['Last GOH DG3']],
              ['EDG', last['MAIN GENERATOR 4'], last['Last TOH EDG'], last['Last GOH EDG']]]
engine_goh = pd.DataFrame(engine_goh, columns=['Natuna', 'Current HR', 'Last TOH', 'Last GOH'])
engine_goh['Last TOH'] = engine_goh['Last TOH'].fillna(0)
engine_goh['Utilization TOH'] = engine_goh['Current HR']-engine_goh['Last TOH']
engine_goh['Utilization GOH'] = engine_goh['Current HR']-engine_goh['Last GOH']
engine_goh['Hours till TOH'] = 12000-engine_goh['Utilization TOH']
engine_goh['Hours till GOH'] = 24000-engine_goh['Utilization GOH']
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
alternator_oh = pd.DataFrame(alternator_oh, columns=['Natuna', 'Last Overhaul (Date)',])

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


hou_c1 = plot_hou1(hou1)
toh = plot_toh(engine_goh['TOH'])
goh = plot_goh(engine_goh['GOH'])
crane_1 = plot_wr1(wire_rope_c1['Utilization'])
au = plot_aoh(alternator_oh['Utilization'])

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

ConMakerShape = shape_cu(last['Conveyor-Maker Check-Up'])
ConMakerCol = col_cu(last['Conveyor-Maker Check-Up'])
CraneMakerShape = shape_cu(last['Crane-Maker Check-Up'])
CraneMakerCol = col_cu(last['Crane-Maker Check-Up'])
EngMakerShape = shape_cu(last['Engine-Maker Check-Up'])
EngMakerCol = col_cu(last['Engine-Maker Check-Up'])
OthersShape = shape_cu(last['Others'])
OthersCol = col_cu(last['Others'])
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

## -----LAYOUT-----
layout = html.Div([
                ## --ROW1--
                html.Br(),
                html.P(children=[html.Strong('Bulk Natuna')],
                       style={'textAlign': 'center', 'fontSize': 27, 'background-color':'white','color':'#2a3f5f','font-family':'Verdana'}),

                dbc.Row([
                    dbc.Col([
                        dbc.Row([
                            dbc.Col([dcc.Graph(figure=hou_c1)], xs=4),
                            dbc.Col([dcc.Graph(figure=crane_1)], xs=8)
                        ]),
                    ], xs=5),

                    dbc.Col([
                        dbc.Row([
                            dcc.Graph(figure=toh)
                        ]),

                        html.Br(),

                        dbc.Row([
                            dcc.Graph(figure=goh)
                        ]),

                        html.Br(),

                        dbc.Row([
                            dcc.Graph(figure=au)
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
                                html.Div(className=ConMakerShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=ConMakerCol,
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
                                html.Div(className=CraneMakerShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=CraneMakerCol,
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
                                html.Div(className=EngMakerShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=EngMakerCol,
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Others')],
                                                style={'textAlign': 'left', 
                                                    'fontSize': 15, 
                                                    'background-color':'white',
                                                    'color':'#2a3f5f',
                                                    'font-family':'Verdana'}),
                                    ]
                                ), color="light", outline=True, style={"height": 55, "margin-bottom": "5px"}
                            ),
                            dbc.Card(
                                html.Div(className=OthersShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=OthersCol,
                                color="light", outline=True, style={"height": 55, "margin-bottom": "5px", "maxWidth": 75}
                            ),
                        ],),

                        dbc.CardGroup([
                            dbc.Card(
                                dbc.CardBody(
                                    [
                                        html.P(children=[html.Strong('Next Docking Intermediate Survey (IS): ' + str(last['Next Docking Intermediate Survey (IS)'].strftime("%d %b %Y")))],
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
                                className=NextDockingCol,
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
                                html.Div(className=ConSpareShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=ConSpareCol,
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
                                html.Div(className=CraneSpareShape, style={
                                                                        "color": "white",
                                                                        "textAlign": "center",
                                                                        "fontSize": 25,
                                                                        "margin": "auto",
                                                                    }),
                                className=CraneSpareCol,
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
                                    html.Div(className=CraneShape, style={
                                                                            "color": "white",
                                                                            "textAlign": "center",
                                                                            "fontSize": 17,
                                                                            "margin": "auto",
                                                                        }),
                                    className=CraneCol,
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
                                        html.Div(className=ConShape, style={
                                                                                "color": "white",
                                                                                "textAlign": "center",
                                                                                "fontSize": 17,
                                                                                "margin": "auto",
                                                                            }),
                                        className=ConCol,
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
                                        html.Div(className=EngShape, style={
                                                                                "color": "white",
                                                                                "textAlign": "center",
                                                                                "fontSize": 17,
                                                                                "margin": "auto",
                                                                            }),
                                        className=EngCol,
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
                                    dbc.CardBody([])
                                ], color="light", outline=True, style={"height": 160, "margin-bottom": "5px"}),
                            ], xs=4),

                            dbc.Col([
                                dbc.Card([
                                    dbc.CardBody([])
                                ], color="light", outline=True, style={"height": 160, "margin-bottom": "5px"}),
                            ], xs=4)
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
