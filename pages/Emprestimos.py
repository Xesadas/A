import dash
from dash import register_page, html, dcc, dash_table, callback, Input, Output, State
import pandas as pd
import numpy as np
from datetime import datetime
import data_processing

# Registra a página
register_page(__name__, path='/')

# Carrega dados e configurações
processed_sheets = data_processing.load_and_process_data()

# Concatena todas as abas e cria fallback para estrutura vazia
base_columns = [
    'data', 'agente', 'beneficiario', 'chave_pix_cpf', 'valor_transacionado',
    'valor_liberado', 'quantidade_parcelas', 'porcentagem_agente', 'taxa_de_juros',
    'extra_agente', 'comissao_agente', 'valor_dualcred', 'nota_fiscal',
    '%trans', '%liberad'
]

if processed_sheets:
    df = pd.concat(processed_sheets.values(), ignore_index=True)
else:
    df = pd.DataFrame(columns=base_columns)

# Configura datas padrão seguras
min_date = df['data'].min() if not df.empty else pd.to_datetime('2025-01-01')
max_date = df['data'].max() if not df.empty else pd.to_datetime('2025-12-31')


# Configurações da página
colors = {
    'background': '#111111',
    'text': '#7FDBFF'
}

input_columns = [
    'data', 'agente', 'beneficiario', 'chave_pix_cpf',
    'valor_transacionado', 'valor_liberado', 'quantidade_parcelas',
    'porcentagem_agente', 'taxa_de_juros', 'extra_agente',
]

numeric_cols = [
    'valor_transacionado', 'valor_liberado', 'taxa_de_juros',
    'comissao_agente', 'extra_agente', 'porcentagem_agente',
    'nota_fiscal', 'quantidade_parcelas'
]

excluir_colunas = ['%_trans.', '%_liberad.', 'acerto_alessandro', 'retirada_felipe', 'máquina']

# Layout da página
layout = html.Div(
    style={'backgroundColor': colors['background'], 'padding': '20px'},
    children=[
        html.H1(
            "Emprestimos DualBank",
            style={
                'textAlign': 'center',
                'color': colors['text'],
                'padding': '20px',
                'marginBottom': '30px'
            }
        ),
        
        # Container de Inputs
        html.Div([
            html.Div([
                html.Label(
                    col.upper(),
                    style={'fontWeight': 'bold', 'color': colors['text']}
                ),
                dcc.DatePickerSingle(
                    id=f'input-{col}',
                    min_date_allowed=pd.to_datetime('2025-01-01'),
                    date=pd.to_datetime('2025-01-01')
                ) if col == "data" else
                dcc.Dropdown(
                    id=f'input-{col}',
                    options=[{'label': f'{x}X', 'value': x} for x in range(1, 19)],
                    value=1
                ) if col == "quantidade_parcelas" else
                dcc.Input(
                    id=f'input-{col}',
                    type='number',
                    min=0,
                    step=0.01,
                    placeholder='0.00',
                    style={
                        'backgroundColor': colors['background'],
                        'color': colors['text'],
                        'border': f'1px solid {colors["text"]}'
                    }
                ) if col in numeric_cols else
                dcc.Input(
                    id=f'input-{col}',
                    type='text',
                    style={
                        'backgroundColor': colors['background'],
                        'color': colors['text'],
                        'border': f'1px solid {colors["text"]}'
                    }
                )
            ], style={'padding': '10px', 'flex': '1'}) for col in input_columns
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'margin': '20px 0'}),
        
        # Filtro de Data
        dcc.DatePickerRange(
            id="date-picker",
            start_date=min_date,
            end_date=max_date,
            display_format="DD/MM/YYYY"
        ),

        # Tabela
        html.Div(
            style={'width': '95%', 'margin': '0 auto', 'overflowX': 'auto'},
            children=[
                dash_table.DataTable(
                    id="tabela-dados",
                    columns=[
                        {"name": col.upper(), "id": col} 
                        for col in df.columns 
                        if col not in excluir_colunas
                    ],
                    data=df.to_dict("records"),
                    page_size=15,
                    style_cell={
                        'textAlign': 'left',
                        'padding': '8px',
                        'border': f'1px solid {colors["text"]}',
                        'backgroundColor': colors['background'],
                        'color': 'white'
                    },
                    style_header={
                        'backgroundColor': colors['background'],
                        'fontWeight': 'bold',
                        'border': f'1px solid {colors["text"]}',
                        'color': colors['text']
                    },
                    editable=True,
                    row_selectable='single'
                )
            ]
        ),
        
        html.Div(
            id="soma-result",
            style={
                "fontSize": "20px",
                "margin": "20px 0",
                "padding": "15px",
                "border": f'1px solid {colors["text"]}',
                "backgroundColor": colors['background'],
                "color": colors['text']
            }
        ),
        
        html.Div([
            html.Button(
                "Salvar Dados",
                id="salvar-btn",
                n_clicks=0,
                style={
                    'backgroundColor': colors['text'],
                    'color': colors['background'],
                    'margin': '5px',
                    'border': 'none',
                    'padding': '10px 20px',
                    'borderRadius': '5px'
                }
            ),
            html.Button(
                "Exportar Planilha",
                id="exportar-btn",
                n_clicks=0,
                style={
                    'backgroundColor': colors['text'],
                    'color': colors['background'],
                    'margin': '5px',
                    'border': 'none',
                    'padding': '10px 20px',
                    'borderRadius': '5px'
                }
            ),
            html.Button(
                "Apagar Linha Selecionada",
                id="apagar-btn",
                n_clicks=0,
                style={
                    'backgroundColor': '#FF4136',
                    'color': 'white',
                    'margin': '5px',
                    'border': 'none',
                    'padding': '10px 20px',
                    'borderRadius': '5px'
                }
            )
        ], style={'margin': '20px 0'}),
        
        html.Div(id="output-mensagem", style={'color': colors['text']}),
        dcc.Download(id="download-dataframe-xlsx")
    ]
)
# =============================================
# CALLBACKS
# =============================================
@callback(
    Output("tabela-dados", "data"),
    Input("date-picker", "start_date"),
    Input("date-picker", "end_date")
)
def filtrar_dados(start_date, end_date):
    try:
        if df.empty:
            return []
            
        start_date = pd.to_datetime(start_date) if start_date else min_date
        end_date = pd.to_datetime(end_date) if end_date else max_date
        
        mask = (df['data'] >= start_date) & (df['data'] <= end_date)
        df_filtrado = df.loc[mask].copy()
        df_filtrado['nota_fiscal'] = (df_filtrado['valor_transacionado'] * 0.032).round(2)
        
        return df_filtrado.to_dict("records")
    except Exception as e:
        print(f"Erro de filtragem: {str(e)}")
        return df.to_dict("records") if not df.empty else []

@callback(
    Output("soma-result", "children"),
    Input("date-picker", "start_date"),
    Input("date-picker", "end_date")
)
def calcular_soma(start_date, end_date):
    try:
        # Converter para datetime e tratar valores inválidos
        start_dt = pd.to_datetime(start_date, errors='coerce') if start_date else df['data'].min()
        end_dt = pd.to_datetime(end_date, errors='coerce') if end_date else df['data'].max()

        # Garantir que as datas são válidas
        start_str = start_dt.strftime('%d/%m/%Y') if not pd.isna(start_dt) else "N/A"
        end_str = end_dt.strftime('%d/%m/%Y') if not pd.isna(end_dt) else "N/A"
        # Aplicar filtro
        mask = (df['data'] >= start_dt) & (df['data'] <= end_dt)
        df_filtrado = df.loc[mask]
        
        # Cálculos
        soma = {
            'Valor_Transacionado': df_filtrado['valor_transacionado'].sum(),
            'Valor_Liberado': df_filtrado['valor_liberado'].sum(),
            'Comissao_Agente': df_filtrado['comissao_agente'].sum(),
            'Valor_DualCred': df_filtrado['valor_dualcred'].sum(),
            'Extra_Agente': df_filtrado['extra_agente'].sum(),
            'nota_fiscal': df_filtrado['nota_fiscal'].sum()
        }
        
        return html.Pre(
            f"RELATÓRIO DE VALORES\n"
            f"──────────────────────\n"
            f"Período: {start_str} - {end_str}\n\n"
            f"Valor Transacionado: R$ {soma['Valor_Transacionado']:,.2f}\n"
            f"Valor Liberado:      R$ {soma['Valor_Liberado']:,.2f}\n"
            f"Comissao Agente:     R$ {soma['Comissao_Agente']:,.2f}\n"  # Key corrected here
            f"Valor Dualcred:      R$ {soma['Valor_DualCred']:,.2f}\n"
            f"Extra Agente:        R$ {soma['Extra_Agente']:,.2f}\n"
            f"Nota Fiscal:         R$ {soma['nota_fiscal']:,.2f}"
        )
    except Exception as e:
        return html.Pre(f"Erro no cálculo: {str(e)}")
    

@callback(
    Output("output-mensagem", "children"),      # Mensagem de status
    Output("download-dataframe-xlsx", "data"),  # Dados para download
    Output("tabela-dados", "data", allow_duplicate=True),  # Dados da tabela
    Output("tabela-dados", "selected_rows"),    # Linhas selecionadas
    [
        Input(f"input-{col}", "value") if col != "data" else 
        Input(f"input-{col}", "date") for col in input_columns
    ],  # Todos os inputs do formulário
    Input("salvar-btn", "n_clicks"),     # Botão Salvar
    Input("exportar-btn", "n_clicks"),   # Botão Exportar
    Input("apagar-btn", "n_clicks"),     # Botão Apagar
    Input("date-picker", "start_date"),  # Filtro data inicial
    Input("date-picker", "end_date"),    # Filtro data final
    State("tabela-dados", "selected_rows"),  # Linhas selecionadas (estado)
    prevent_initial_call=True
)
def gerenciar_dados(*args):
    global processed_sheets
    #processed_sheets = data_processing.load_and_process_data() 
    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else None

    try:
        # 1. Dividir os argumentos corretamente
        num_form_inputs = len(input_columns)
        form_inputs = args[:num_form_inputs]
        button_clicks = args[num_form_inputs:num_form_inputs+3]
        start_date, end_date = args[num_form_inputs+3:num_form_inputs+5]
        selected_rows = args[-1] if len(args) > num_form_inputs+5 else []

        # 2. Converter datas para o formato correto
        start_date = pd.to_datetime(start_date, errors='coerce') or df['data'].min()
        end_date = pd.to_datetime(end_date, errors='coerce') or df['data'].max()
        
        # 3. Aplicar filtro inicial
        mask = (df['data'] >= start_date) & (df['data'] <= end_date)
        filtered_df = df.loc[mask].copy()
    except Exception as e:
        print(f"Erro no pré-processamento: {str(e)}")
        return dash.no_update, dash.no_update, df.to_dict("records"), []

    # 4. Determinar ação do usuário
    try:
        if triggered_id == "salvar-btn":
            return salvar_dados(form_inputs, filtered_df, start_date, end_date)
            
        elif triggered_id == "exportar-btn":
            export_data = data_processing.exportar_dados(processed_sheets)  # Dicionário de abas
            return (
            "✅ Planilha exportada com sucesso!",  # Mensagem
            export_data,                          # Dados download
            dash.no_update,                       # Mantém tabela
            []                                    # Limpa seleção
        )
        
            
        elif triggered_id == "apagar-btn":
            return apagar_linha(selected_rows, start_date, end_date)
            
    except Exception as e:
        print(f"Erro na ação: {str(e)}")
        return f"Erro: {str(e)}", None, dash.no_update, []

    return dash.no_update, dash.no_update, filtered_df.to_dict("records"), []

def salvar_dados(form_inputs, filtered_df, start_date, end_date):
    try:
        # Coletar dados do formulário
        novos_dados = {}
        for col, val in zip(input_columns, form_inputs):
            if col == 'data':
                # Converter e tratar datas inválidas
                dt = pd.to_datetime(val, errors='coerce', dayfirst=False)
                novos_dados[col] = dt if not pd.isna(dt) else pd.Timestamp('2025-01-01')
                
            elif col in numeric_cols:
                novos_dados[col] = round(float(val or 0), 2)
            else:
                novos_dados[col] = str(val).strip() if val else ''

        # 2. Garantir substituição de NaT residual
        if pd.isna(novos_dados['data']):
            novos_dados['data'] = pd.Timestamp('2025-01-01')
            
        # 2. Cálculos automáticos
        novos_dados['comissao_agente'] = round(
            novos_dados['valor_liberado'] * (novos_dados['porcentagem_agente'] / 100), 2
        )
        novos_dados['valor_dualcred'] = (
            novos_dados['valor_transacionado'] 
            - novos_dados['valor_liberado'] 
            - novos_dados['taxa_de_juros'] 
            - novos_dados['comissao_agente'] 
            - novos_dados['extra_agente']
        )
        novos_dados['%trans'] = round(
            (novos_dados['valor_dualcred'] / novos_dados['valor_transacionado'] * 100), 2
        ) if novos_dados['valor_transacionado'] else 0
        novos_dados['%liberad'] = round(
            (novos_dados['valor_dualcred'] / novos_dados['valor_liberado'] * 100), 2
        ) if novos_dados['valor_liberado'] else 0
        novos_dados['nota_fiscal'] = round(novos_dados['valor_transacionado'] * 0.032, 2)

        # 3. Atualizar DataFrame global
        global df
        df = pd.concat([df, pd.DataFrame([novos_dados])], ignore_index=True)
        data_processing.salvar_no_excel(df) 

        # 4. Reaplicar filtro após atualização
        mask = (df['data'] >= start_date) & (df['data'] <= end_date)
        filtered_df = df.loc[mask]
        
        return (
            "✅ Dados salvos com sucesso!", 
            None, 
            filtered_df.to_dict("records"), 
            []
        )
    except Exception as e:
        print(f"Erro ao salvar: {str(e)}")
        return f"❌ Erro ao salvar: {str(e)}", None, dash.no_update, []

def apagar_linha(selected_rows, start_date, end_date):
    global df
    try:
        if not selected_rows:
            return "⚠️ Selecione uma linha antes de apagar!", None, dash.no_update, []
        
        # 1. Obter índices reais no DataFrame global
        mask = (df['data'] >= pd.to_datetime(start_date)) & (df['data'] <= pd.to_datetime(end_date))
        filtered_indices = df[mask].index.tolist()
        
        # 2. Mapear índices filtrados para índices globais
        global_indices = [filtered_indices[i] for i in selected_rows]
        
        # 3. Remover linhas
        df = df.drop(global_indices)
        data_processing.salvar_no_excel(df)
        
        # 4. Atualizar DataFrame filtrado
        mask = (df['data'] >= pd.to_datetime(start_date)) & (df['data'] <= pd.to_datetime(end_date))
        filtered_df = df.loc[mask]
        
        return (
            "✅ Linha apagada com sucesso!", 
            None, 
            filtered_df.to_dict("records"), 
            []
        )
    except Exception as e:
        print(f"Erro ao apagar: {str(e)}")
        return f"❌ Erro ao apagar linha: {str(e)}", None, dash.no_update, []