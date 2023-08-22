######-----Import Dash-----#####
import dash
from dash import dcc
from dash import html, callback
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output

import gspread

#####-----Create a Dash app instance-----#####
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.LITERA, dbc.icons.BOOTSTRAP],
    use_pages=True,
    )
app.title = 'Equipment Dashboard'

##-----Navbar
navbar = dbc.NavbarSimple(
    children=[
        dbc.NavLink(
                    [
                    html.Div(page["name"], className="ms-2"),
                    ],
                    href=page["path"],
                    active="exact",
                )
                for page in dash.page_registry.values()
    ],
    brand="Equipment Dashboard",
    brand_href="#",
    color="#2a3f5f",
    dark=True,
)

## -----LAYOUT-----
app.layout = dbc.Container([
    dcc.Interval(id="timer", interval=1000*120, n_intervals=0),
    dcc.Store(id="store", data={}),
    navbar,    
    dbc.Row(
        [
            dbc.Col(
                [
                    dash.page_container
                ])
        ]
    ),
], fluid=True, style={'background-color':'#F1F4F4'})

def load_google_sheets_data():
    cred_file = 'equipment-dashboard-a2d53f750591.json'
    gc = gspread.service_account(cred_file)

    # Open the Google Sheets document
    database = gc.open('Database')

    data = {}
    for sheet in database.worksheets():
        sheet_data = sheet.get_all_records()
        data[sheet.title] = sheet_data

    return data

@callback(
    Output('store', 'data'),
    Input('timer', 'n_intervals')
)
def update_data(n):
    data = load_google_sheets_data()
    return data

######-----Start the Dash server-----#####
if __name__ == "__main__":
    app.run_server(debug=True)