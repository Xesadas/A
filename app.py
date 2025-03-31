import dash
from dash import Dash, html, dcc, page_container
import pandas as pd
import numpy as np
from datetime import datetime
import io
import data_processing

# Inicialização do app
app = Dash(__name__, use_pages=True, suppress_callback_exceptions=True)
server = app.server

# Configuração global do DataFrame
df = data_processing.load_and_process_data()

# =============================================
# LAYOUT PRINCIPAL ATUALIZADO COM NAVEGAÇÃO
# =============================================
app.layout = html.Div(
    style={'backgroundColor': '#111111'},
    children=[
        dcc.Location(id='url'),
        html.Nav(
            [
                dcc.Link('Empréstimos', href='/', style={'color': '#7FDBFF', 'marginRight': '20px'}),
                dcc.Link('Relatório de agentes', href='/agents-analysis', style={'color': '#7FDBFF'})
            ],
            style={'padding': '20px', 'backgroundColor': '#222222'}
        ),
        dash.page_container
    ]
)
if __name__ == "__main__":  
    app.run_server(debug=True)