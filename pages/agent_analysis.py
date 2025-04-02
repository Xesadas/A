from dash import dcc, html, dash_table, Input, Output, callback, register_page
import pandas as pd
import data_processing
import logging
from datetime import datetime
import numpy as np

logger = logging.getLogger(__name__)
register_page(__name__, path='/agents-analysis')

# Função auxiliar para limpar e validar dados
def clean_agent_data(raw_data):
    """Garante a integridade dos dados com fallbacks robustos"""
    try:
        # Converter para DataFrame se necessário
        if isinstance(raw_data, dict):
            df = pd.concat(raw_data.values(), ignore_index=True)
        else:
            df = raw_data.copy()

        # Criar colunas essenciais se ausentes
        essential_columns = {
            'agente': 'Não Informado',
            'data': pd.to_datetime('2025-01-01'),
            'valor_transacionado': 0.0,
            'valor_liberado': 0.0,
            'comissao_agente': 0.0,
            'extra_agente': 0.0
        }

        for col, default in essential_columns.items():
            if col not in df.columns:
                df[col] = default
                logger.warning(f"Coluna '{col}' criada artificialmente")

        # Tratamento de datas
        df['data'] = pd.to_datetime(df['data'], errors='coerce').fillna(pd.to_datetime('2025-01-01'))
        
        # Tratamento do campo agente
        df['agente'] = (
            df['agente']
            .fillna('Não Informado')
            .astype(str)
            .str.strip()
            .replace({
                '': 'Não Informado', 
                'nan': 'Não Informado', 
                'None': 'Não Informado',
                'null': 'Não Informado',
                np.nan: 'Não Informado'
            })
        )

        # Garantir tipos numéricos
        numeric_cols = ['valor_transacionado', 'valor_liberado', 'comissao_agente', 'extra_agente']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        return df

    except Exception as e:
        logger.error(f"Erro na limpeza de dados: {str(e)}")
        return pd.DataFrame(columns=list(essential_columns.keys()))

# Layout atualizado
layout = html.Div(
    style={
        'backgroundColor': '#111111', 
        'padding': '20px', 
        'minHeight': '100vh',
        'color': '#7FDBFF'
    },
    children=[
        html.H1(
            "Análise de Agentes",
            style={
                'textAlign': 'center', 
                'padding': '20px',
                'marginBottom': '30px'
            }
        ),
        
        dcc.Interval(
            id='refresh-interval',
            interval=30*1000,
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
        # Carregar e processar dados
        raw_data = data_processing.load_and_process_data()
        df = clean_agent_data(raw_data)

        # Configurar datas padrão
        min_date = df['data'].min() if not df.empty else datetime(2025, 1, 1)
        max_date = df['data'].max() if not df.empty else datetime(2025, 12, 31)

        # Gerar opções válidas para dropdown
        valid_agents = [agente for agente in df['agente'].unique() 
                      if agente not in [None, 'Não Informado', '']]

        return [
            html.Div(
                style={'marginBottom': '30px'},
                children=[
                    dcc.Dropdown(
                        id='agent-selector',
                        options=[{'label': 'Todos', 'value': 'all'}] + 
                                [{'label': agente, 'value': agente} 
                                for agente in sorted(valid_agents)],
                        value='all',
                        placeholder="Selecione um agente...",
                        style={'width': '100%', 'maxWidth': '400px'}
                    ),
                ]
            ),
            
            html.Div(
                style={'marginBottom': '30px'},
                children=[
                    dcc.DatePickerRange(
                        id="agent-date-picker",
                        min_date_allowed=min_date,
                        max_date_allowed=max_date,
                        start_date=min_date,
                        end_date=max_date,
                        display_format="DD/MM/YYYY",
                        style={'width': '100%'}
                    )
                ]
            ),
            
            dash_table.DataTable(
                id='agent-table',
                page_size=15,
                style_table={
                    'overflowX': 'auto',
                    'marginBottom': '30px'
                },
                style_cell={
                    'backgroundColor': '#222222',
                    'color': '#7FDBFF',
                    'border': '1px solid #7FDBFF',
                    'padding': '10px'
                },
                style_header={
                    'backgroundColor': '#333333',
                    'fontWeight': 'bold',
                    'fontSize': '16px'
                }
            ),
            
            html.Div(
                id="agent-stats",
                style={
                    'padding': '20px',
                    'border': '2px solid #7FDBFF',
                    'borderRadius': '10px'
                }
            )
        ]
    
    except Exception as e:
        logger.error(f"Erro crítico: {str(e)}")
        return html.Div(
            "Sistema temporariamente indisponível. Tente recarregar a página.",
            style={'color': '#FF5555', 'textAlign': 'center', 'padding': '50px'}
        )

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
        df = clean_agent_data(data_processing.load_and_process_data())
        
        if df.empty:
            return [], [], html.Div("Nenhum dado disponível para análise")

        # Filtrar por datas
        start_date = pd.to_datetime(start_date) if start_date else df['data'].min()
        end_date = pd.to_datetime(end_date) if end_date else df['data'].max()
        
        filtered_df = df[
            (df['data'] >= start_date) & 
            (df['data'] <= end_date)
        ]

        # Filtrar por agente
        if selected_agent and selected_agent != 'all':
            filtered_df = filtered_df[filtered_df['agente'] == selected_agent]

        # Formatar valores
        display_df = filtered_df.copy()
        numeric_cols = ['valor_transacionado', 'valor_liberado', 'comissao_agente', 'extra_agente']
        for col in numeric_cols:
            display_df[col] = display_df[col].apply(
                lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else 'R$ 0,00')

        # Gerar estatísticas
        stats = {
            'Transações Totais': filtered_df['valor_transacionado'].sum(),
            'Valor Liberado Total': filtered_df['valor_liberado'].sum(),
            'Comissões Totais': filtered_df['comissao_agente'].sum(),
            'Extras Totais': filtered_df['extra_agente'].sum()
        }

        # Criar layout das estatísticas
        stats_content = [
            html.H3(
                "Resumo Financeiro",
                style={'marginBottom': '15px'}
            ),
            html.Div(
                style={'display': 'grid', 'gridTemplateColumns': 'repeat(2, 1fr)', 'gap': '15px'},
                children=[
                    html.Div(
                        style={'padding': '15px', 'border': '1px solid #7FDBFF', 'borderRadius': '5px'},
                        children=[
                            html.Div(key, style={'fontWeight': 'bold'}),
                            html.Div(f"R$ {value:,.2f}")
                        ]
                    ) for key, value in stats.items()
                ]
            )
        ]

        # Configurar colunas da tabela
        columns = [{
            "name": col.replace('_', ' ').title(),
            "id": col,
            "type": "numeric" if col in numeric_cols else "text"
        } for col in ['data', 'agente'] + numeric_cols]

        return (
            columns,
            display_df.to_dict('records'),
            stats_content
        )

    except Exception as e:
        logger.error(f"Erro na atualização: {str(e)}")
        return [], [], html.Div(
            "Erro ao carregar dados. Atualizando...",
            style={'color': '#FF5555', 'textAlign': 'center'}
        )