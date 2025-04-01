# pages/agent_analysis.py (VERSÃO FINAL CORRIGIDA)
from dash import dcc, html, dash_table, Input, Output, callback, register_page, no_update
import pandas as pd
import data_processing
import logging
from datetime import datetime
import numpy as np

logger = logging.getLogger(__name__)
register_page(__name__, path='/agents-analysis')

# Componente para armazenar dados no cliente
dcc.Store(id='agent-data-store')

# Função auxiliar para formatar valores
def format_currency(value):
    try:
        return f'R$ {float(value):,.2f}' if not pd.isna(value) else 'N/A'
    except:
        return 'N/A'

# Layout atualizado com atualização dinâmica
layout = html.Div(
    style={'backgroundColor': '#111111', 'padding': '20px', 'minHeight': '100vh'},
    children=[
        html.H1(
            "Análise de Agentes",
            id='analysis-title',
            style={
                'textAlign': 'center',
                'color': '#7FDBFF',
                'padding': '20px',
                'marginBottom': '30px'
            }
        ),
        
        dcc.Interval(
            id='refresh-interval',
            interval=5*1000,  # Atualiza a cada 30 segundos
            n_intervals=0
        ),
        
        dcc.Loading(
            id="loading-analysis",
            type="circle",
            children=[
                html.Div(id='dynamic-content')
            ]
        )
    ]
)

# Callback principal para carregamento dinâmico
@callback(
    Output('dynamic-content', 'children'),
    Input('refresh-interval', 'n_intervals'),
    prevent_initial_call=True
)
def update_dynamic_content(n):
    try:
        # Recarrega os dados periodicamente
        global df
        df = data_processing.load_and_process_data()
        
        # Atualiza as opções do dropdown
        agents = df['agente'].unique().tolist() if 'agente' in df.columns else []
        
        return [
            dcc.Dropdown(
                id='agent-selector',
                options=[{'label': 'Todos', 'value': 'all'}] + 
                        [{'label': agente, 'value': agente} for agente in agents],
                value='all',
                placeholder="Selecione...",
                style={'width': '300px', 'marginBottom': '20px', 'color': '#111111'}
            ),
            
            dcc.DatePickerRange(
                id="agent-date-picker",
                min_date_allowed=df['data'].min().date(),
                max_date_allowed=df['data'].max().date(),
                start_date=df['data'].min().date(),
                end_date=df['data'].max().date(),
                display_format="DD/MM/YYYY",
                style={'marginBottom': '20px'}
            ),
            
            dash_table.DataTable(
                id='agent-table',
                page_size=15,
                style_table={'overflowX': 'auto'},
                style_cell={
                    'textAlign': 'left',
                    'padding': '8px',
                    'border': '1px solid #7FDBFF',
                    'backgroundColor': '#111111',
                    'color': 'white'
                },
                style_header={
                    'backgroundColor': '#111111',
                    'fontWeight': 'bold',
                    'border': '1px solid #7FDBFF',
                    'color': '#7FDBFF'
                }
            ),
            
            html.Div(
                id="agent-stats",
                style={
                    "fontSize": "18px",
                    "margin": "20px 0",
                    "padding": "15px",
                    "border": '1px solid #7FDBFF',
                    "backgroundColor": '#111111',
                    "color": '#7FDBFF'
                }
            )
        ]
    
    except Exception as e:
        logger.error(f"Erro na atualização dinâmica: {str(e)}")
        return html.Div("Sistema em atualização...", style={'color': '#7FDBFF'})

# Callback para atualização dos dados
@callback(
    [Output('agent-table', 'columns'),
     Output('agent-table', 'data'),
     Output('agent-stats', 'children')],
    [Input('agent-date-picker', 'start_date'),
     Input('agent-date-picker', 'end_date'),
     Input('agent-selector', 'value')]
)
def update_analysis(start_date, end_date, selected_agent):
    try:
        # Verificação de dados
        if df.empty or 'agente' not in df.columns:
            return [], [], html.Div("Dados não disponíveis", style={'color': 'red'})
        
        # Filtragem por datas
        start_date = pd.to_datetime(start_date) if start_date else df['data'].min()
        end_date = pd.to_datetime(end_date) if end_date else df['data'].max()
        filtered_df = df[(df['data'] >= start_date) & (df['data'] <= end_date)].copy()
        
        # Filtragem por agente
        if selected_agent and selected_agent != 'all':
            filtered_df = filtered_df[filtered_df['agente'] == selected_agent]
        
        # Colunas numéricas para cálculos (mantém como float)
        numeric_cols = ['valor_transacionado', 'valor_liberado', 'comissão_agente', 'extra_agente']
        for col in numeric_cols:
            filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce').fillna(0)
        
        # Formatação apenas para exibição
        display_df = filtered_df.copy()
        for col in numeric_cols:
            display_df[col] = display_df[col].apply(format_currency)
        
        # Cálculos seguros
        try:
            totals = {
                'Transacionado': filtered_df['valor_transacionado'].sum(),
                'Liberado': filtered_df['valor_liberado'].sum(),
                'Comissão': filtered_df['comissão_agente'].sum(),
                'Extra': filtered_df['extra_agente'].sum()
            }
        except Exception as e:
            logger.error(f"Erro nos cálculos: {str(e)}")
            totals = {k: 0 for k in numeric_cols}
        
        # Montagem das estatísticas
        stats_content = [
            html.H3("Estatísticas Consolidados" if selected_agent == 'all' else f"Estatísticas de {selected_agent}"),
            html.P(f"Período: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"),
            *[html.P(f"{key}: {format_currency(value)}") for key, value in totals.items()]
        ]
        
        # Colunas da tabela
        columns = [{"name": "Data", "id": "data"}] + [
            {"name": col.replace('_', ' ').title(), "id": col} 
            for col in display_df.columns if col in numeric_cols + ['agente']
        ]
        
        return (
            columns,
            display_df.to_dict('records'),
            stats_content
        )
    
    except Exception as e:
        logger.error(f"Erro crítico: {str(e)}")
        return [], [], html.Div("Erro na atualização dos dados", style={'color': 'red'})