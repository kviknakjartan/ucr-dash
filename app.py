from dash import Dash, dcc, html, Input, Output, callback
from get_ucr_data import UCRDataFetcher
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import load_figure_template
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import pandas as pd

load_figure_template('cyborg')

app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])

server = app.server

fetcher = UCRDataFetcher()

app.layout = dbc.Container(fluid=True, children=[

        dbc.Row([
            html.Div(
                className="app-header",
                children=[
                    html.Div('UCR Crime Data Viewer', className="app-header--title")
                ]
            ),
            html.P(["The US ",
                    html.A("Uniform Crime Reporting", 
                        href="https://www.fbi.gov/how-we-can-help-you/more-fbi-services-and-information/ucr", 
                        target="_blank"),
                    ''' Program (UCR) is a nationwide, cooperative, statistical effort of more than 18,000 law enforcement 
                        agencies voluntarily reporting data on crimes brought to their attention. Data from ''',
                    html.A("https://www.fbi.gov/how-we-can-help-you/more-fbi-services-and-information/ucr/publications", 
                        href="https://www.fbi.gov/how-we-can-help-you/more-fbi-services-and-information/ucr/publications", 
                        target="_blank")
            ])
        ]),
        dbc.Row([
            dbc.Col(
                html.Div([
                    "1. Select crime statistic:",
                    dcc.Dropdown(list(fetcher.dataDict.keys()), next(iter(fetcher.dataDict)), id='indicator-dropdown'),
                    dcc.Loading(
                        id="loading-output",
                        type="default", # or "cube", "circle", "dot", "graph", "spinner"
                        children=html.Div(id='output-div')
                    )
                ]), width = 6),
            dbc.Col(
                html.Div([
                    "2. Select variable/relationship:",
                    dcc.Dropdown(id='variable-dropdown'),
                ]), width = 6)
        ]),

        dbc.Row([
            dbc.Col(
                html.Div([
                    "3. Select group/variable:",
                    dcc.Dropdown(id='group-dropdown'),
                ]), width = 6),
            dbc.Col(
                html.Div([
                    "4. Select measure:",
                    dcc.Dropdown(id='measure-dropdown'),
                ]), width = 6)
        ]),

        dbc.Row([
            html.Div([
                "5. Select data series:",
                dcc.Dropdown(id='series-dropdown', multi = True),
            ])
        ]),

        dbc.Row([
            html.Div([
                dcc.Graph(id='graph-with-slider', responsive='auto'),
                dcc.RangeSlider(id='year-slider', 
                                step = 1, 
                                tooltip={"placement": "bottom", "always_visible": True},
                                marks=None)
            ], style={'width': '100%', 'height': '600px', 'margin-bottom': '20px'})
        ]),

        dbc.Row([
            html.Div([
                """NOTE:  Because the number of agencies submitting arrest data varies from year to year, users are 
                cautioned about making direct comparisons between arrest totals and those published in previous years' 
                editions of Crime in the United States. Further, arrest figures may vary widely from state to state because 
                some Part II crimes are not considered crimes in some states."""
            ], style={'font-style': 'italic', 'margin-bottom': '10px'})
        ])
])


@callback(
    [Output('variable-dropdown', 'options'),
    Output('variable-dropdown', 'value'),
    Output('output-div', 'children')],
    [Input('indicator-dropdown', 'value')])
def set_variable_options(selected_indicator):
    fetcher.loadTable(selected_indicator)
    return [{'label': i, 'value': i} for i in fetcher.dataDict[selected_indicator].keys()], \
        next(iter(fetcher.dataDict[selected_indicator])), None

@callback(
    [Output('group-dropdown', 'options'),
    Output('group-dropdown', 'value')],
    [Input('indicator-dropdown', 'value'),
    Input('variable-dropdown', 'value')])
def set_group_options(selected_indicator, selected_variable):
    return [{'label': i, 'value': i} for i in fetcher.dataDict[selected_indicator][selected_variable].keys()], \
        next(iter(fetcher.dataDict[selected_indicator][selected_variable]))

@callback(
    [Output('measure-dropdown', 'options'),
    Output('measure-dropdown', 'value')],
    [Input('indicator-dropdown', 'value'),
    Input('variable-dropdown', 'value'),
    Input('group-dropdown', 'value')])
def set_measure_options(selected_indicator, selected_variable, selected_group):
    return [{'label': i, 'value': i} for i in fetcher.dataDict[selected_indicator][selected_variable][selected_group].keys()], \
        next(iter(fetcher.dataDict[selected_indicator][selected_variable][selected_group]))

@callback(
    Output('series-dropdown', 'options'),
    Input('indicator-dropdown', 'value'),
    Input('variable-dropdown', 'value'),
    Input('group-dropdown', 'value'),
    Input('measure-dropdown', 'value'))
def set_series_options(selected_indicator, selected_variable, selected_group, selected_measure):
    return [{'label': i, 'value': i} for i in \
        fetcher.dataDict[selected_indicator][selected_variable][selected_group][selected_measure].columns]

@callback(
    [Output('year-slider', 'min'),
    Output('year-slider', 'max'),
    Output('year-slider', 'value')],
    [Input('indicator-dropdown', 'value'),
    Input('variable-dropdown', 'value'),
    Input('group-dropdown', 'value'),
    Input('measure-dropdown', 'value')])
def set_slider_minmax_data(selected_indicator, selected_variable, selected_group, selected_measure):
    minimum = fetcher.dataDict[selected_indicator][selected_variable][selected_group][selected_measure].index.min()
    maximum = fetcher.dataDict[selected_indicator][selected_variable][selected_group][selected_measure].index.max()
    return minimum, maximum, (minimum, maximum)

@callback(
    Output('graph-with-slider', 'figure'),
    Input('year-slider', 'value'),
    Input('indicator-dropdown', 'value'),
    Input('variable-dropdown', 'value'),
    Input('group-dropdown', 'value'),
    Input('measure-dropdown', 'value'),
    Input('series-dropdown', 'value'))
def generate_plot(selected_years, selected_indicator, selected_variable, selected_group, selected_measure, selected_series):
    
    if selected_series is None:
        selected_series = []
    df = fetcher.dataDict[selected_indicator][selected_variable][selected_group][selected_measure][selected_series]
    meta = fetcher.metaDict[selected_indicator][selected_variable][selected_group][selected_measure]

    fig = make_subplots()
    for col in df.columns:
        customdata = pd.DataFrame(columns = ['Volume', 'Population', 'Demographic', 'Agencies', 'Notes'])
        for key, df_m in meta.items():
            customdata[key] = df_m[col].fillna('')
        # Add traces
        fig.add_trace(
            go.Scatter(x=df.index,
                y=df[col], 
                name=col,
                customdata = customdata,
                hovertemplate = create_hover_template(df[col], col, meta))
        )
    fig.update_layout(
        title_text=f"<b>{selected_indicator}:</b> <i>{selected_variable} - {selected_group}</i>",
        xaxis=dict(range=selected_years),
        legend=dict(
            x=0.1,  # x-position (0.1 is near left)
            y=0.7,  # y-position (0.9 is near top)
            xref="container",
            yref="container",
            orientation = 'h'
        )
    )
    # Set x-axis title
    fig.update_xaxes(title_text="Year")

    # Set y-axes titles
    fig.update_yaxes(title_text=selected_measure)
    return fig

def create_hover_template(y, col, meta_data):


    if isinstance(y.iloc[0], int):
        template =  '<b>Value: %{y:.0f}'+'<br><i>Year: %{x:.0f}</i></b>'
    elif np.abs(y.iloc[0]) < 1:                   
        template =  '<b>Value: %{y:.3f}'+'<br><i>Year: %{x:.0f}</i></b>'
    else:
        template =  '<b>Value: %{y:.2f}'+'<br><i>Year: %{x:.0f}</i></b>'

    for key, df in meta_data.items():

        if key == 'Volume':
            template += '<br>Volume: %{customdata[0]:,.0f}'
        elif key == 'Population':
            template += '<br>Population: %{customdata[1]:,.0f}'
        elif key == 'Demographic':
            template += '<br>Demographic population: %{customdata[2]:,.0f}'
        elif key == 'Agencies':
            template += '<br>Number of agencies reporting: %{customdata[3]:,.0f}'
        elif key == 'Notes':
            template += '<br><i>%{customdata[4]}</i>'

    return template

if __name__ == '__main__':
    app.run(debug=True)
