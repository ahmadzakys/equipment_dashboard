######-----Import Dash-----#####
import dash
from dash import dcc
from dash import html, callback
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
from dash.exceptions import PreventUpdate

from pptx import Presentation
from pptx.util import Cm, Pt

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.io as pio
from plotly.subplots import make_subplots

import os
from datetime import date
from dateutil.relativedelta import relativedelta

dash.register_page(__name__, name='Download')

##-----page

## -----LAYOUT-----
layout = html.Div([
                html.Br(),
                
                html.Div([
                    html.P('Download slides here', id='date_download', style={'fontSize': 15, 'color':'#2a3f5f','font-family':'Verdana'}),
                    html.Button('Download Slides', 
                                      id='download_button', 
                                      n_clicks=0, 
                                      style={'fontSize': 15, 'color':'#2a3f5f','font-family':'Verdana','display': 'inline-block'}), 
                    dcc.Download(id='download')
                    ]),

    html.Footer('ABL',
            style={'textAlign': 'center', 
                   'fontSize': 20, 
                   'background-color':'#2a3f5f',
                   'color':'white',
                   'position': 'fixed',
                    'bottom': '0',
                    'width': '100%'})

    ], style={
        'paddingLeft':'10px',
        'paddingRight':'10px',
    })

#### Callback Auto Update Chart & Data

@callback(
    [Output('download', 'data')],
    [Input('download_button', 'n_clicks'),
     Input('store', 'data')]
)
def update_charts(n_clicks, data):
    if n_clicks == 0:
        raise PreventUpdate
    ######################
    # Pre Processing
    ######################
    df_sumatra = pd.DataFrame(data['Bulk Sumatra'])
    df_sumatra = df_sumatra.replace(r'^\s*$', np.nan, regex=True)

    df_sumatra[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']] = \
    df_sumatra[['Last Overhaul GEN01', 'Last Overhaul GEN02', 'Last Overhaul GEN03', 'Last Overhaul GEN04']].fillna(date.today())
    df_sumatra['Last Overhaul GEN01'] = pd.to_datetime(df_sumatra['Last Overhaul GEN01'], dayfirst = True)
    df_sumatra['Last Overhaul GEN02'] = pd.to_datetime(df_sumatra['Last Overhaul GEN02'], dayfirst = True)
    df_sumatra['Last Overhaul GEN03'] = pd.to_datetime(df_sumatra['Last Overhaul GEN03'], dayfirst = True)
    df_sumatra['Last Overhaul GEN04'] = pd.to_datetime(df_sumatra['Last Overhaul GEN04'], dayfirst = True)
    df_sumatra['Next Docking Intermediate Survey (IS)'] = pd.to_datetime(df_sumatra['Next Docking Intermediate Survey (IS)'], dayfirst = True)

    # Hydrolic Oil Utilization
    last = df_sumatra.iloc[-1]
    hou1 = (last['CRANE 1']-last['Last Rpl RH CR01'])/6000
    hou1 = round(hou1*100,0)
    hou1 = min(hou1, 100)

    hou2 = (last['CRANE 2']-last['Last Rpl RH CR02'])/6000
    hou2 = round(hou2*100,0)
    hou2 = min(hou2, 100)

    # Wire Rope
    wire_rope_c1 = [['HW RHOL', last['HW RHOL CR01'], last['CRANE 1']],
                    ['HW LHOL', last['HW LHOL CR01'], last['CRANE 1']],
                    ['CW RHOL', last['CW RHOL CR01'], last['CRANE 1']],
                    ['CW LHOL', last['CW LHOL CR01'], last['CRANE 1']]]
    wire_rope_c1 = pd.DataFrame(wire_rope_c1, columns=['Crane 1', 'Last Replacement', 'Current RH'])
    wire_rope_c1['Utilization RH'] = wire_rope_c1['Current RH']-wire_rope_c1['Last Replacement']
    wire_rope_c1['Utilization %'] = round(wire_rope_c1['Utilization RH']/4000*100,0)
    util1 = []
    for i in wire_rope_c1['Utilization %']:
        util1.append(min(i, 100))
    wire_rope_c1['Utilization'] = util1
    wire_rope_c1['Max'] = 100

    wire_rope_c2 = [['HW RHOL', last['HW RHOL CR02'], last['CRANE 2']],
                    ['HW LHOL', last['HW LHOL CR02'], last['CRANE 2']],
                    ['CW RHOL', last['CW RHOL CR02'], last['CRANE 2']],
                    ['CW LHOL', last['CW LHOL CR02'], last['CRANE 2']]]
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
    engine_goh = pd.DataFrame(engine_goh, columns=['Sumatra', 'Current HR', 'Last TOH', 'Last GOH'])
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

    # Conveyor Belt Replacement
    conveyor_br = [['BC1', last['BC1'], last['Last RPL HR BC1']],
                ['BC2', last['BC2'], last['Last RPL HR BC2']],
                ['BC3', last['BC3'], last['Last RPL HR BC3']],
                ['BC4', last['BC4'], last['Last RPL HR BC4']],
                ['BC5', last['BC5'], last['Last RPL HR BC5']],
                ['SHL1', last['SHL1'], last['Last RPL HR SHL1']],
                ['SHL2', last['SHL2'], last['Last RPL HR SHL2']]]
    conveyor_br = pd.DataFrame(conveyor_br, columns=['Sumatra', 'Current RH', 'Last Rpl HR'])
    conveyor_br['Utilization (hrs)'] = conveyor_br['Current RH']-conveyor_br['Last Rpl HR']
    conveyor_br['Utilization %'] = [round(conveyor_br['Utilization (hrs)'][0]/8000*100,0),
                                round(conveyor_br['Utilization (hrs)'][1]/8000*100,0),
                                round(conveyor_br['Utilization (hrs)'][2]/5000*100,0),
                                round(conveyor_br['Utilization (hrs)'][3]/5000*100,0),
                                round(conveyor_br['Utilization (hrs)'][4]/5000*100,0),
                                round(conveyor_br['Utilization (hrs)'][5]/5000*100,0),
                                round(conveyor_br['Utilization (hrs)'][6]/5000*100,0),]

    util = []
    for i in conveyor_br['Utilization %']:
        util.append(min(i, 100))
    conveyor_br['Utilization'] = util
    conveyor_br['Max'] = 100

    # Alternator OH
    alternator_oh = [['GEN01', last['Last Overhaul GEN01'],],
                    ['GEN02', last['Last Overhaul GEN02'],],
                    ['GEN03', last['Last Overhaul GEN03'],],
                    ['GEN04', last['Last Overhaul GEN04'],]]
    alternator_oh = pd.DataFrame(alternator_oh, columns=['Sumatra', 'Last Overhaul (Date)',])

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
    crane = last['Crane']
    df_crane = pd.DataFrame({'names' : ['progress',' '],
                    'values' :  [crane, 100 - crane]})

    conveyor = last['Conveyor']
    df_conveyor = pd.DataFrame({'names' : ['progress',' '],
                    'values' :  [conveyor, 100 - conveyor]})

    engine = last['Engine']
    df_engine = pd.DataFrame({'names' : ['progress',' '],
                    'values' :  [engine, 100 - engine]})


    def shape_cu(value):
        shape=''
        if value == 'Green': shape="img/green check.png"
        elif value == 'Yellow': shape="img/yellow cross.png"
        elif value == 'Red': shape="img/red cross.png"
        elif value == 'Gray': shape="img/gray blank.png"
        return(shape)

    def shape_cs(value):
        shape=''
        if value<=(50) : shape="img/red cross.png"
        elif value<=(80) : shape="img/yellow cross.png"
        elif value<=100 : shape="img/green check.png"
        else : shape="img/gray blank.png"
        return(shape)

    def col_day_delta(value):
        col=''
        if value<=180 : col="img/red calendar.png"
        elif value<=365 : col="img/yellow calendar.png"
        else: col="img/green calendar.png"
        return(col)
    
    path = os.getcwd()  
    parent = os.path.dirname(path)
  
    t_sumatra = last['Next Docking Intermediate Survey (IS)']
    ConMakerShape_sumatra = shape_cu(last['Conveyor-Maker Check-Up'])
    CraneMakerShape_sumatra = shape_cu(last['Crane-Maker Check-Up'])
    EngMakerShape_sumatra = shape_cu(last['Engine-Maker Check-Up'])
    OthersShape_sumatra = shape_cu(last['Others'])
    NextDockingCol_sumatra = col_day_delta((last['Next Docking Intermediate Survey (IS)'] - pd.Timestamp.today()).days)
    ConSpareShape_sumatra = shape_cu(last['Conveyor Belt Spares'])
    CraneSpareShape_sumatra = shape_cu(last['Crane-Wire Rope Spares']) 
    CraneShape_sumatra = shape_cs(last['Crane'])
    ConShape_sumatra = shape_cs(last['Conveyor'])
    EngShape_sumatra = shape_cs(last['Engine'])

    # download_date = html.Strong('Download slides as per ' + str(last['Date'].strftime("%d %b %Y")))
    # download_date ='Text'

    #Chart
    sumatra_hou_c1 = plot_hou1(hou1)
    sumatra_hou_c2= plot_hou2(hou2)
    sumatra_toh = plot_toh(engine_goh['TOH'])
    sumatra_goh = plot_goh(engine_goh['GOH'])
    sumatra_crane_1 = plot_wr1(wire_rope_c1['Utilization'], wire_rope_c1)
    sumatra_crane_2 = plot_wr2(wire_rope_c2['Utilization'], wire_rope_c2)
    sumatra_cbu = plot_cbr(conveyor_br['Utilization'], 'Sumatra', conveyor_br)
    sumatra_au = plot_aoh(alternator_oh['Utilization'], 'Sumatra', alternator_oh)
    sumatra_cs_crane = plot_cs(df_crane)
    sumatra_cs_conveyor = plot_cs(df_conveyor)
    sumatra_cs_engine = plot_cs(df_engine)

    #Save Img 
    pio.write_image(sumatra_hou_c1, 'img/hou_c1 Bulk Sumatra.png')
    pio.write_image(sumatra_hou_c2, 'img/hou_c2 Bulk Sumatra.png')
    pio.write_image(sumatra_toh, 'img/toh Bulk Sumatra.png')
    pio.write_image(sumatra_goh, 'img/goh Bulk Sumatra.png')
    pio.write_image(sumatra_crane_1, 'img/crane_1 Bulk Sumatra.png')
    pio.write_image(sumatra_crane_2, 'img/crane_2 Bulk Sumatra.png')
    pio.write_image(sumatra_cbu, 'img/cbu Bulk Sumatra.png')
    pio.write_image(sumatra_au, 'img/au Bulk Sumatra.png')
    pio.write_image(sumatra_cs_crane, 'img/cs_crane Bulk Sumatra.png')
    pio.write_image(sumatra_cs_conveyor, 'img/cs_conveyor Bulk Sumatra.png')
    pio.write_image(sumatra_cs_engine, 'img/cs_engine Bulk Sumatra.png')

    def to_pptx(bytes_io):
        # Create presentation
        pptx = parent + '//' + 'slide_master.pptx'
        prs = Presentation(pptx)

        # define slidelayouts 
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])
        slide2 = prs.slides.add_slide(prs.slide_layouts[0])
        slide3 = prs.slides.add_slide(prs.slide_layouts[0])
        slide4 = prs.slides.add_slide(prs.slide_layouts[0])
        slide5 = prs.slides.add_slide(prs.slide_layouts[0])
        slide6 = prs.slides.add_slide(prs.slide_layouts[0])
        slide7 = prs.slides.add_slide(prs.slide_layouts[0])
        slide8 = prs.slides.add_slide(prs.slide_layouts[0])
        slide9 = prs.slides.add_slide(prs.slide_layouts[0])

        # slide1 Bulk Sumatra
        slide1.placeholders[10].text = 'EQUIPMENT DASHBOARD \n Bulk Sumatra'
        slide1.shapes.add_picture('img/hou_c1 Bulk Sumatra.png', Cm(0.75), Cm(3.15), Cm(4), Cm(4))
        slide1.shapes.add_picture('img/hou_c2 Bulk Sumatra.png', Cm(0.75), Cm(7.4), Cm(4), Cm(4))
        slide1.shapes.add_picture('img/crane_1 Bulk Sumatra.png', Cm(5), Cm(3.15), Cm(9.5), Cm(4))
        slide1.shapes.add_picture('img/crane_2 Bulk Sumatra.png', Cm(5), Cm(7.4), Cm(9.5), Cm(4))
        slide1.shapes.add_picture('img/toh Bulk Sumatra.png', Cm(14.75), Cm(3.15), Cm(11), Cm(4))
        slide1.shapes.add_picture('img/goh Bulk Sumatra.png', Cm(14.75), Cm(7.4), Cm(11), Cm(4))
        slide1.shapes.add_picture('img/cbu Bulk Sumatra.png', Cm(0.75), Cm(11.65), Cm(13.75), Cm(5.5))
        slide1.shapes.add_picture('img/au Bulk Sumatra.png', Cm(14.75), Cm(11.65), Cm(11), Cm(5.5))
        slide1.shapes.add_picture('img/cs_crane Bulk Sumatra.png', Cm(26.2), Cm(14.2), Cm(2), Cm(2.5))
        slide1.shapes.add_picture('img/cs_conveyor Bulk Sumatra.png', Cm(28.5), Cm(14.2), Cm(2), Cm(2.5))
        slide1.shapes.add_picture('img/cs_engine Bulk Sumatra.png', Cm(30.9), Cm(14.2), Cm(2), Cm(2.5))

        slide1.placeholders[12].text = 'Conveyor-Maker Check-Up'
        slide1.placeholders[13].text = 'Crane-Maker Check-Up'
        slide1.placeholders[14].text = 'Engine-Maker Check-Up'
        slide1.placeholders[15].text = 'Others-Shi ELD'
        slide1.placeholders[16].text = 'Next Docking \n Intermediate \n Survey (IS)'
        slide1.placeholders[17].text = 'Conveyor Belt Spares'
        slide1.placeholders[18].text = 'Crane-Wire Rope Spares'
        slide1.placeholders[19].text = 'Crane'
        slide1.placeholders[20].text = 'Conv.'
        slide1.placeholders[21].text = 'Eng.'

        cal = slide1.shapes.add_picture(NextDockingCol_sumatra, Cm(30), Cm(7.7))

        thn = slide1.shapes.add_textbox(Cm(30.6), Cm(8), 1, 0.5)
        thn = thn.text_frame.paragraphs[0]
        thn.text = t_sumatra.strftime("%Y")
        font = thn.font
        font.bold = True
        font.name = 'Arial Black'
        font.size = Pt(14)

        tgl = slide1.shapes.add_textbox(Cm(30.8), Cm(8.6), 1, 0.5)
        tgl = tgl.text_frame.paragraphs[0]
        tgl.text = t_sumatra.strftime("%d")
        font = tgl.font
        font.bold = True
        font.name = 'Arial Black'
        font.size = Pt(18)

        bln = slide1.shapes.add_textbox(Cm(30.8), Cm(9.35), 1, 0.5)
        bln = bln.text_frame.paragraphs[0]
        bln.text = t_sumatra.strftime("%b")
        font = bln.font
        font.name = 'Arial'
        font.size = Pt(14)

        slide1.shapes.add_picture(ConMakerShape_sumatra, Cm(32.1), Cm(3.3))
        slide1.shapes.add_picture(CraneMakerShape_sumatra, Cm(32.1), Cm(4.4))
        slide1.shapes.add_picture(EngMakerShape_sumatra, Cm(32.1), Cm(5.5))
        slide1.shapes.add_picture(OthersShape_sumatra, Cm(32.1), Cm(6.6))
        slide1.shapes.add_picture(ConSpareShape_sumatra, Cm(32.1), Cm(10.7))
        slide1.shapes.add_picture(CraneSpareShape_sumatra, Cm(32.1), Cm(11.8))
        slide1.shapes.add_picture(CraneShape_sumatra, Cm(27.4), Cm(12.9))
        slide1.shapes.add_picture(ConShape_sumatra, Cm(29.8), Cm(12.9))
        slide1.shapes.add_picture(EngShape_sumatra, Cm(32.1), Cm(12.9))

        prs.save(bytes_io)    

    return dcc.send_bytes(to_pptx, 'Equipment Dashboard.pptx')


######################
# Plot Function
######################
##-----Function
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
    fig.add_annotation(x=0.5, y=0.85, text='<b>Hydrolic Oil Utilization</b><br>Crane no 1<br>(Max 6,000hrs)', 
                       font=dict(size=57), showarrow=False)
    fig.update_layout({'height':700,'width':750,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':3, 'l':3, 'r':3}, "autosize": True})
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
    fig.add_annotation(x=0.5, y=0.85, text='<b>Hydrolic Oil Utilization</b><br>Crane no 2<br>(Max 6,000hrs)', 
                       font=dict(size=57), showarrow=False)
    fig.update_layout({'height':700,'width':750,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':3, 'l':3, 'r':3}, "autosize": True})
    return(fig)


#-- 1. Plot Hydrolic Utilization (Bulk Sumba)
def plot_hou1_s(value):
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
    fig.add_annotation(x=0.5, y=0.7, text='<b>Hydrolic Oil Utilization</b><br>Crane no 1<br>(Max 6,000hrs)', 
                       font=dict(size=57), showarrow=False)
    fig.update_layout({'height':1200,'width':750,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':150, 'l':3, 'r':3}, "autosize": True})
    return(fig)

def plot_hou2_s(value):
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
    fig.add_annotation(x=0.5, y=0.7, text='<b>Hydrolic Oil Utilization</b><br>Crane no 2<br>(Max 6,000hrs)', 
                       font=dict(size=57), showarrow=False)
    fig.update_layout({'height':1200,'width':750,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':150, 'l':3, 'r':3}, "autosize": True})
    return(fig)


#-- 1. Plot Hydrolic Utilization
def plot_hou1_n(value):
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
                       font=dict(size=57), showarrow=False)
    fig.update_layout({'height':2500,'width':750,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':3, 'b':300, 'l':3, 'r':3}, "autosize": True})
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
        title = {'text': "DG01", 'font': {'size': 60}},
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
        title = {'text': "DG02", 'font': {'size': 60}},
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
        title = {'text': "DG03", 'font': {'size': 60}},
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
        title = {'text': "DG04", 'font': {'size': 60}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[3],
            'bar':{'color':color[3]},
            }))
    fig.add_annotation(x=0.5, y=0.90, text='<b>Engine Overhaul</b><br>TOH to be conducted @12,000hrs<br>', 
                       font=dict(size=75), showarrow=False)
    fig.update_layout({'height':700,'width':1900,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
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
        title = {'text': "DG01", 'font': {'size': 60}},
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
        title = {'text': "DG02", 'font': {'size': 60}},
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
        title = {'text': "DG03", 'font': {'size': 60}},
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
        title = {'text': "DG04", 'font': {'size': 60}},
        gauge = {
            'axis': {'range': [None, 100], 'tickwidth': 0, 'tickcolor': "white", 'visible': False},
            'bgcolor': "white",
            'borderwidth': 3,
            'bordercolor': color[3],
            'bar':{'color':color[3]},
            }))
    fig.add_annotation(x=0.5, y=0.90, text='<b>Engine Overhaul</b><br>GOH to be conducted @24,000hrs<br>', 
                       font=dict(size=75), showarrow=False)
    fig.update_layout({'height':700,'width':1900,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 1'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 1</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':700,'width':1700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 2'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 2</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':700,'width':1700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)


#-- 3. Plot Wire Rope (Bulk Derawan & Bulk Karimun)
def plot_wr1_dk(value, wire_rope_c1):
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 1'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 1</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':700,'width':2700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

############################################################################################################################

def plot_wr2_dk(value, wire_rope_c2):
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 2'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 2</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':700,'width':2700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)


#-- 3. Plot Wire Rope (Bulk Sumba)
def plot_wr1_s(value, wire_rope_c1):
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 1'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 1</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':1200,'width':1700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

############################################################################################################################

def plot_wr2_s(value, wire_rope_c2):
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 2'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 2</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':1200,'width':1700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)



#-- 3. Plot Wire Rope (Bulk Sumba)
def plot_wr1_n(value, wire_rope_c1):
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
                         textfont_size=60,
                         textangle=0,
                         hovertemplate='%{y:y}%',
                         marker_color =color,
                         width = anchos, name = 'Crane 1'))
    fig.update_yaxes(visible=False)
    fig.update_layout(title = "<b>Crane 1</b>",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 78,
                      showlegend=False,
                      xaxis = dict(tickfont = dict(size=60)))
    fig.update_layout({'height':2500,'width':1700,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':250, 'b':200, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

#-- 4. Plot Conveyor Belt Replacement
def plot_cbr(value, bulk, conveyor_br):
    anchos = 0.5
    color=[]
    for i in value:
        if i<=20 : color.append('#488f31')
        elif i<=40 : color.append('#88c580')
        elif i<=60 : color.append('#ffe600')
        elif i<=80 : color.append('#f7a258')
        else: color.append('#de425b')

    fig = go.Figure()
    fig.add_trace(go.Bar(x = conveyor_br['Max'][::-1],
                         y = conveyor_br[bulk][::-1],
                         marker_color='#ededed',
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.add_trace(go.Bar(x = value[::-1], 
                         y = conveyor_br[bulk][::-1],
                         text=value[::-1].apply(lambda x: '{0:1.0f}%'.format(x)),
                         textposition='outside',
                         textfont_size=40,
                         marker_color =color[::-1],
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.update_xaxes(visible=False)
    fig.update_layout(title = "<b>Conveyor Belt Utilization</b> <br>FB 8K & CB 5K",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 80,
                      showlegend=False,
                      yaxis = dict(tickfont = dict(size=55)))
    fig.update_layout({'height':1000,'width':2440,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':240,'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)


#-- 5. Plot Alternator OH
def plot_aoh(value, bulk, alternator_oh):
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
                         y = alternator_oh[bulk][::-1],
                         marker_color='#ededed',
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.add_trace(go.Bar(x = value[::-1], 
                         y = alternator_oh[bulk][::-1],
                         text=value[::-1].apply(lambda x: '{0:1.0f}%'.format(x)),
                         textposition=position[::-1],
                         textfont_size=55,
                         marker_color =color[::-1],
                         width = anchos, name = 'Crane 1',
                         orientation='h'))
    fig.update_xaxes(visible=False)
    fig.update_layout(title = "<b>Alternator Utilization</b> <br> To be cleaned and service every 5 years",
                      title_x=0.5,
                      barmode = 'overlay',
                      title_font_size = 80,
                      showlegend=False,
                      yaxis = dict(tickfont = dict(size=55)))
    fig.update_layout({'height':1000,'width':2000,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':240,'b':3, 'l':3, 'r':3}, "autosize": True
                      })

    return(fig)

#-- 6. Plot Critical Spare
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
                      title_font_size = 45,
                      showlegend=False,)
    fig.update_layout({'height':480,'width':400,'plot_bgcolor': '#ffffff', 'paper_bgcolor': '#ffffff',
                       'margin' : {'t':120, 'b':3, 'l':3, 'r':3}, "autosize": True})
    val=str(df['values'][0])
    fig.add_annotation(x=0.5, y=0.5, text=f'<b>{val}%</b>', 
                           font=dict(size=45), showarrow=False)
    return(fig)
