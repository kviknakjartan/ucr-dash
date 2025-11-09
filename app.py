from dash import Dash, dcc, html, Input, Output, callback
from get_ucr_data import UCRDataFetcher
import dash_bootstrap_components as dbc
from dash_bootstrap_templates import load_figure_template
import plotly.graph_objects as go
from plotly.subplots import make_subplots

load_figure_template('cyborg')

app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])

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
                        agencies voluntarily reporting data on crimes brought to their attention.'''
                    ])
        ]),
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
            ], style={'width': '100%', 'height': '600px'})
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
    df = fetcher.dataDict[selected_indicator][selected_variable][selected_group][selected_measure]
    if selected_series is None:
        selected_series = []
    df = df[selected_series]
    fig = make_subplots()
    for col in df.columns:
        # Add traces
        fig.add_trace(
            go.Scatter(x=df.index,
                y=df[col], 
                name=col,
                hovertemplate =
                'Value: %{y:.2f}'+
                '<br>Year: %{x:.0f}')
        )
    fig.update_layout(
        title_text="Graph 1: Global mean surface temperature (instrumental record) 1850 to present",
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

if __name__ == '__main__':
    app.run(debug=True)
