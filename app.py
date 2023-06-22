######-----Import Dash-----#####
import dash
from dash import dcc
from dash import html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output

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
    navbar,

    dbc.Row(
        [
            dbc.Col(
                [
                    dash.page_container
                ])
        ]
    )
], fluid=True, style={'background-color':'#F1F4F4'})

######-----Start the Dash server-----#####
if __name__ == "__main__":
    app.run_server()