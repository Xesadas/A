# pages/agent_analysis.py
from dash import dcc, html, dash_table, Input, Output, callback
import pandas as pd
from datetime import datetime
from dash import register_page
import data_processing

register_page(__name__, path='/agents-analysis')
df = data_processing.load_and_process_data()
df['agente'] = df['agente'].fillna('').astype(str)

# Layout da página de análise de agentes
layout = html.Div(
    style={'backgroundColor': '#111111', 'padding': '20px'},
    children=[
        html.H1(
            "Análise de Agentes",
            style={
                'textAlign': 'center',
                'color': '#7FDBFF',
                'padding': '20px',
                'marginBottom': '30px'
            }
        ),
        
        dcc.Dropdown(
            id='agent-selector',
            options=[
                {'label': 'Todos os Agentes', 'value': 'all'}
            ] + [
                {'label': str(agente), 'value': str(agente)} 
                for agente in df['agente'].dropna().unique()
                if agente not in [None, 'NaN', 'nan', '']  # Filtra valores inválidos
            ],
            value='all',  # Valor padrão
            placeholder="Selecione um agente...",
            style={
                'width': '300px',
                'marginBottom': '20px',
                'color': '#111111'
            }
        ),
        
        dcc.DatePickerRange(
            start_date=df['data'].min().date(),
            end_date=df['data'].max().date(),
            id="agent-date-picker",
            display_format="DD/MM/YYYY",
            style={'marginBottom': '20px'}
        ),
        
        html.Div(
            id='agent-analysis-table',
            style={'width': '95%', 'margin': '0 auto', 'overflowX': 'auto'},
            children=[
                dash_table.DataTable(
                    id='agent-table',
                    page_size=15,
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
                )
            ]
        ),
        
        html.Div(
            id="agent-stats",
            style={
                "fontSize": "20px",
                "margin": "20px 0",
                "padding": "15px",
                "border": '1px solid #7FDBFF',
                "backgroundColor": '#111111',
                "color": '#7FDBFF'
            }
        )
    ]
)

@callback(
    [Output('agent-table', 'columns'),
     Output('agent-table', 'data'),
     Output('agent-stats', 'children')],
    [Input('agent-date-picker', 'start_date'),
     Input('agent-date-picker', 'end_date'),
     Input('agent-selector', 'value')]
)
def update_agent_analysis(start_date, end_date, selected_agent):
    from data_processing import df
    
    # 1. Pré-processamento de datas
    df['data'] = pd.to_datetime(df['data'], errors='coerce')
    df = df.dropna(subset=['data'])
    
    # 2. Validação das datas
    start_date = pd.to_datetime(start_date) if start_date else df['data'].min()
    end_date = pd.to_datetime(end_date) if end_date else df['data'].max()
    
    # 3. Filtragem por data
    mask = (df['data'] >= start_date) & (df['data'] <= end_date)
    filtered_df = df.loc[mask]
    
    # 4. Lógica de seleção do agente
    if selected_agent:
        if selected_agent == 'all':
            # Agregado de todos os agentes
            columns = [
                {"name": "Métrica", "id": "metric"},
                {"name": "Valor", "id": "value"}
            ]
            
            totals = {
                'Valor Transacionado': filtered_df['valor_transacionado'].sum(),
                'Valor Liberado': filtered_df['valor_liberado'].sum(),
                'Comissão Total': filtered_df['comissão_agente'].sum(),
                'Extra Total': filtered_df['extra_agente'].sum(),
                'Valor Dualcred': filtered_df['valor_dualcred'].sum()
            }
            
            data = [{'metric': k, 'value': f'R$ {v:,.2f}'} for k, v in totals.items()]
            
            stats = html.Div([
                html.H3("Consolidação Geral", style={'color': '#7FDBFF'}),
                html.P(f"Período: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"),
                html.P(f"Total Transacionado: R$ {totals['Valor Transacionado']:,.2f}"),
                html.P(f"Total Liberado: R$ {totals['Valor Liberado']:,.2f}"),
                html.P(f"Comissão Total: R$ {totals['Comissão Total']:,.2f}"),
                html.P(f"Extra Total: R$ {totals['Extra Total']:,.2f}"),
                html.P(f"Valor Dualcred: R$ {totals['Valor Dualcred']:,.2f}")
            ])
            
            return columns, data, stats
        
        else:
            # Detalhamento por agente específico
            filtered_df = filtered_df[filtered_df['agente'] == selected_agent]
            
            columns = [
                {"name": "Cliente", "id": "beneficiário"},
                {"name": "Data", "id": "data"},
                {"name": "Transacionado (R$)", "id": "valor_transacionado"},
                {"name": "Liberado (R$)", "id": "valor_liberado"},
                {"name": "Comissão (R$)", "id": "comissão_agente"},
                {"name": "Extra (R$)", "id": "extra_agente"}
            ]
            
            stats = html.Div([
                html.H3(f"Desempenho de {selected_agent}", style={'color': '#7FDBFF'}),
                html.P(f"Total Clientes: {filtered_df['beneficiário'].nunique()}"),
                html.P(f"Transacionado: R$ {filtered_df['valor_transacionado'].sum():,.2f}"),
                html.P(f"Liberado: R$ {filtered_df['valor_liberado'].sum():,.2f}"),
                html.P(f"Comissão: R$ {filtered_df['comissão_agente'].sum():,.2f}"),
                html.P(f"Extra: R$ {filtered_df['extra_agente'].sum():,.2f}")
            ])
            
            return columns, filtered_df.to_dict('records'), stats
    
    # 5. Visão agregada padrão (sem seleção)
    agent_stats = filtered_df.groupby('agente', as_index=False).agg({
        'valor_transacionado': 'sum',
        'valor_liberado': 'sum',
        'comissão_agente': 'sum',
        'extra_agente': 'sum',
        'valor_dualcred': 'sum'
    })
    
    columns = [{"name": col.upper().replace('_', ' '), "id": col} 
               for col in agent_stats.columns]
    
    stats = html.Div([
        html.H3("Visão Comparativa", style={'color': '#7FDBFF'}),
        html.P(f"Agentes Ativos: {len(agent_stats)}"),
        html.P(f"Maior Transação: R$ {agent_stats['valor_transacionado'].max():,.2f}"),
        html.P(f"Média Liberado/Agente: R$ {agent_stats['valor_liberado'].mean():,.2f}"),
        html.P(f"Comissão Total: R$ {agent_stats['comissão_agente'].sum():,.2f}")
    ])
    
    return columns, agent_stats.to_dict('records'), stats