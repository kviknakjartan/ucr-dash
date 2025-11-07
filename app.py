from dash import Dash, dcc, html, Input, Output, callback
from get_ucr_data import UCRDataFetcher
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import load_figure_template

load_figure_template('cyborg')

app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])

fetcher = UCRDataFetcher()

app.layout = dbc.Container(fluid=True, children=[
        dbc.Row([
            dbc.Col(
                html.Div([
                    "1. Select crime statistic:",
                    dcc.Dropdown(list(fetcher.dataDict.keys()), next(iter(fetcher.dataDict)), id='indicator-dropdown')
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
                html.Hr()
            ], className="bg-white p-0")
        ]),

        dbc.Row([
            html.Div([
                "5. Select data series:",
                dcc.Dropdown(id='series-dropdown', multi = True),
            ])
        ]),

        dbc.Row([
            html.Div([
                #dcc.Graph(id='graph-with-slider'),
                dcc.RangeSlider(id='year-slider', 
                                step = 1, 
                                tooltip={"placement": "bottom", "always_visible": True},
                                marks=None)
            ])
        ])
])


@callback(
    [Output('variable-dropdown', 'options'),
    Output('variable-dropdown', 'value')],
    [Input('indicator-dropdown', 'value')])
def set_variable_options(selected_indicator):
    fetcher.loadTable(selected_indicator)
    return [{'label': i, 'value': i} for i in fetcher.dataDict[selected_indicator].keys()], \
        next(iter(fetcher.dataDict[selected_indicator]))

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
    print(minimum,maximum)
    return minimum, maximum, (minimum, maximum)

if __name__ == '__main__':
    app.run(debug=True)
