# pages/agent_analysis.py (VERSÃO FINAL CORRIGIDA)
from dash import dcc, html, dash_table, Input, Output, callback, register_page, no_update
import pandas as pd
import data_processing
import logging
from datetime import datetime
import numpy as np

logger = logging.getLogger(__name__)
register_page(__name__, path='/agents-analysis')

# Função auxiliar para limpar e validar dados
def clean_agent_data(df):
    """Garante a integridade dos dados do agente"""
    try:
        # Preenchimento de valores ausentes e tratamento de strings
        df['agente'] = (
            df['agente']
            .fillna('Não Informado')
            .astype(str)
            .str.strip()
            .replace({'': 'Não Informado', 'nan': 'Não Informado', 'None': 'Não Informado'})
        )
        
        # Garantir nomes de colunas consistentes
        df.columns = [col.lower().replace('ç', 'c').replace('ã', 'a').replace('õ', 'o') 
                     for col in df.columns]
        
        return df
    
    except Exception as e:
        logger.error(f"Erro na limpeza de dados: {str(e)}")
        return pd.DataFrame()

# Layout atualizado
layout = html.Div(
    style={'backgroundColor': '#111111', 'padding': '20px', 'minHeight': '100vh'},
    children=[
        html.H1(
            "Análise de Agentes",
            style={'textAlign': 'center', 'color': '#7FDBFF', 'padding': '20px'}
        ),
        
        dcc.Interval(
            id='refresh-interval',
            interval=30*1000,  # 30 segundos
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

# Callback para conteúdo dinâmico
@callback(
    Output('dynamic-content', 'children'),
    Input('refresh-interval', 'n_intervals')
)
def update_dynamic_content(n):
    try:
        global df
        df = clean_agent_data(data_processing.load_and_process_data())
        
        # Gerar opções válidas para dropdown
        valid_agents = [agente for agente in df['agente'].unique() 
                      if agente not in [None, 'Não Informado', '']]
        
        return [
            dcc.Dropdown(
                id='agent-selector',
                options=[{'label': 'Todos', 'value': 'all'}] + 
                        [{'label': agente, 'value': agente} 
                         for agente in sorted(valid_agents)],
                value='all',
                placeholder="Selecione...",
                style={'width': '300px', 'marginBottom': '20px'}
            ),
            
            dcc.DatePickerRange(
                id="agent-date-picker",
                min_date_allowed=df['data'].min(),
                max_date_allowed=df['data'].max(),
                start_date=df['data'].min(),
                end_date=df['data'].max(),
                display_format="DD/MM/YYYY"
            ),
            
            dash_table.DataTable(
                id='agent-table',
                page_size=15,
                style_table={'overflowX': 'auto'},
                style_cell={
                    'backgroundColor': '#111111',
                    'color': 'white',
                    'border': '1px solid #7FDBFF'
                },
                style_header={
                    'backgroundColor': '#111111',
                    'color': '#7FDBFF',
                    'fontWeight': 'bold'
                }
            ),
            
            html.Div(
                id="agent-stats",
                style={
                    "marginTop": "20px",
                    "padding": "15px",
                    "border": "1px solid #7FDBFF",
                    "color": "#7FDBFF"
                }
            )
        ]
    
    except Exception as e:
        logger.error(f"Erro crítico: {str(e)}")
        return html.Div("Sistema em manutenção. Tente novamente em alguns minutos.", 
                      style={'color': '#FF0000'})

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
        # Verificação inicial de dados
        if df.empty or 'agente' not in df.columns:
            return [], [], html.Div("Dados não disponíveis no momento")
        
        # Filtragem por datas
        filtered_df = df[
            (df['data'] >= pd.to_datetime(start_date)) & 
            (df['data'] <= pd.to_datetime(end_date))
        ]
        
        # Filtragem por agente
        if selected_agent and selected_agent != 'all':
            filtered_df = filtered_df[filtered_df['agente'] == selected_agent]
        
        # Colunas numéricas com tratamento de erros
        numeric_cols = {
            'valor_transacionado': 0.0,
            'valor_liberado': 0.0,
            'comissao_agente': 0.0,  # Nome sanitizado
            'extra_agente': 0.0
        }
        
        for col, default in numeric_cols.items():
            filtered_df[col] = pd.to_numeric(filtered_df[col], errors='coerce').fillna(default)
        
        # Formatação para exibição
        display_df = filtered_df.copy()
        for col in numeric_cols:
            display_df[col] = display_df[col].apply(
                lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else 'N/A')
        
        # Cálculos seguros
        stats = {
            'Total Transacionado': filtered_df['valor_transacionado'].sum(),
            'Total Liberado': filtered_df['valor_liberado'].sum(),
            'Comissões': filtered_df['comissao_agente'].sum(),
            'Extras': filtered_df['extra_agente'].sum()
        }
        
        # Montagem do layout
        stats_content = [
            html.H3("Estatísticas Gerais" if selected_agent == 'all' 
                   else f"Estatísticas de {selected_agent}"),
            html.P(f"Período: {pd.to_datetime(start_date).strftime('%d/%m/%Y')} - "
                  f"{pd.to_datetime(end_date).strftime('%d/%m/%Y')}"),
            *[html.P(f"{k}: R$ {v:,.2f}") for k, v in stats.items()]
        ]
        
        # Colunas da tabela
        columns = [{"name": col.replace('_', ' ').title(), "id": col} 
                  for col in display_df.columns if col in list(numeric_cols.keys()) + ['data', 'agente']]
        
        return (
            columns,
            display_df.to_dict('records'),
            stats_content
        )
    
    except Exception as e:
        logger.error(f"Erro na atualização: {str(e)}")
        return [], [], html.Div("Erro temporário. Atualizando dados...")
