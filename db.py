import dash
from dash import dcc, html, dash_table, Input, Output, State
import pandas as pd
import numpy as np
from datetime import datetime
import io
from dash.exceptions import PreventUpdate

# Configuração de cores
colors = {
    'background': '#111111',
    'text': '#7FDBFF'
}
meses = {
    1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR',
    5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO',
    9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
}
def sanitize_column_name(col):
    return (
        str(col)
        .strip()
        .lower()
        .replace(" ", "_")
        .replace("(", "")
        .replace(")", "")
        .replace("?", "")
    )

# Carregar dados do Excel
file_path = 'a.xlsx'
sheets_to_read = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
# Ler cada planilha e combinar em um único DataFrame
df_list = []
for sheet in sheets_to_read:
    df_sheet = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
    # Sanitizar colunas antes de concatenar
    df_sheet.columns = [sanitize_column_name(col) for col in df_sheet.columns]
    df_list.append(df_sheet)
df = pd.concat(df_list, ignore_index=True)

app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server

# =============================================
# FUNÇÕES DE PRÉ-PROCESSAMENTO (ATUALIZADAS)
# =============================================

def salvar_no_excel(df):
    try:
        with pd.ExcelWriter('a.xlsx', engine='openpyxl') as writer:
            for month_num, sheet_name in meses.items():
                month_df = df[df['data'].dt.month == month_num].copy()
                month_df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"Erro ao salvar: {str(e)}")

def exportar_dados(filtered_df):
    try:
        # Criar buffer de memória
        buffer = io.BytesIO()
        
        # Criar ExcelWriter
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Para cada mês no DataFrame filtrado
            for month_num, sheet_name in meses.items():
                month_df = filtered_df[filtered_df['data'].dt.month == month_num].copy()
                month_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        buffer.seek(0)  # Rebobinar o buffer
        return (
            None,
            dcc.send_bytes(buffer.getvalue(), filename="dados_exportados.xlsx"),
            dash.no_update,
            []
        )
    except Exception as e:
        print(f"Erro na exportação: {str(e)}")
        return f"❌ Erro na exportação: {str(e)}", None, dash.no_update, []



def calcular_valor_dualcred(row):
    """Calcula o Valor_DualCred conforme o Excel"""
    return (
        row['valor_transacionado']
        - row['valor_liberado']
        - row['taxa_de_juros']
        - row['comissão_alessandro']
        - row['extra_alessandro']
    )

def atualizar_porcentagens(df):
    """Atualiza %trans e %liberad com base no valor_dualcred"""
    df['%trans'] = np.where(
        df['valor_transacionado'] != 0,
        (df['valor_dualcred'] / df['valor_transacionado']) * 100,
        0
    ).round(2)
    
    df['%liberad'] = np.where(
        df['valor_liberado'] != 0,
        (df['valor_dualcred'] / df['valor_liberado']) * 100,
        0
    ).round(2)
    return df

def calcular_nota_fiscal(df):
    """Calcula 3.2% do Valor_Transacionado para a coluna Nota_fiscal"""
    df['nota_fiscal'] = (df['valor_transacionado'] * 0.032).round(2)
    return df  # ← Adicione esta função


# Pré-processamento inicial
df.columns = [sanitize_column_name(col) for col in df.columns]
df['data'] = pd.to_datetime(
    df['data'],
    dayfirst=True,
    errors='coerce'
).fillna(pd.to_datetime('2025-01-01'))


numeric_cols = [
    'valor_transacionado', 'valor_liberado', 'taxa_de_juros',
    'comissão_alessandro', 'extra_alessandro', 'qtd_parcelas',
    'porcentagem_alessandro', 'nota_fiscal'
]

for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)

# Calcular valor_dualcred e porcentagens
df['valor_dualcred'] = df.apply(calcular_valor_dualcred, axis=1).round(2)
df = atualizar_porcentagens(df)
df = calcular_nota_fiscal(df) 
# =============================================
# LAYOUT (COM INPUTS NUMÉRICOS CORRIGIDOS)
# =============================================
input_columns = [
    'data',
    'beneficiário',
    'chave_pix_cpf',
    'valor_transacionado',
    'valor_liberado',
    'quantidade_parcelas',
    'porcentagem_alessandro',
    'taxa_de_juros',
    'extra_alessandro'
]

app.layout = html.Div(
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
            start_date=df['data'].min(),
            end_date=df['data'].max(),
            display_format="DD/MM/YYYY"
        ),

        # Tabela
        html.Div(
            style={'width': '95%', 'margin': '0 auto', 'overflowX': 'auto'},
            children=[
                dash_table.DataTable(
                    id="tabela-dados",
                    columns=[{"name": col.upper(), "id": col} for col in df.columns],
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
@app.callback(
    Output("tabela-dados", "data"),
    Input("date-picker", "start_date"),
    Input("date-picker", "end_date")
)
def filtrar_dados(start_date, end_date):
    try:
        start_date = pd.to_datetime(start_date) if start_date else df['data'].min()
        end_date = pd.to_datetime(end_date) if end_date else df['data'].max()
        
        mask = (df['data'] >= start_date) & (df['data'] <= end_date)
        df_filtrado = df.loc[mask].copy()
        df_filtrado = calcular_nota_fiscal(df_filtrado)
        
        return df_filtrado.to_dict("records")
    except Exception as e:
        print(f"Erro de filtragem: {str(e)}")
        return df.to_dict("records")

@app.callback(
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
            'Comissão_Alessandro': df_filtrado['comissão_alessandro'].sum(),
            'Valor_DualCred': df_filtrado['valor_dualcred'].sum(),
            'Extra_Alessandro': df_filtrado['extra_alessandro'].sum()
        }
        
        return html.Pre(
            f"RELATÓRIO CONSOLIDADO\n"
            f"──────────────────────\n"
            f"Período: {start_str} - {end_str}\n\n"  # Usar strings tratadas
            f"Valor Transacionado: R$ {soma['Valor_Transacionado']:,.2f}\n"
            f"Valor Liberado:      R$ {soma['Valor_Liberado']:,.2f}\n"
            f"Comissão Alessandro: R$ {soma['Comissão_Alessandro']:,.2f}\n"
            f"Valor Dualcred:      R$ {soma['Valor_DualCred']:,.2f}\n"
            f"Extra Alessandro:    R$ {soma['Extra_Alessandro']:,.2f}"
        )
    except Exception as e:
        return html.Pre(f"Erro no cálculo: {str(e)}")

    

@app.callback(
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
    global df
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
            return exportar_dados(filtered_df)
            
        elif triggered_id == "apagar-btn":
            return apagar_linha(selected_rows, start_date, end_date)
            
    except Exception as e:
        print(f"Erro na ação: {str(e)}")
        return f"Erro: {str(e)}", None, dash.no_update, []

    return dash.no_update, dash.no_update, filtered_df.to_dict("records"), []

def salvar_dados(form_inputs, filtered_df, start_date, end_date):
    try:
        # 1. Coletar dados do formulário
        novos_dados = {}
        for col, val in zip(input_columns, form_inputs):
            if col == 'data':
                # Converter e tratar datas inválidas
                dt = pd.to_datetime(val, errors='coerce', dayfirst=True)
                novos_dados[col] = dt if not pd.isna(dt) else pd.Timestamp('2025-01-01')
                
            elif col in numeric_cols:
                novos_dados[col] = round(float(val or 0), 2)
            else:
                novos_dados[col] = str(val).strip() if val else ''

        # 2. Garantir substituição de NaT residual
        if pd.isna(novos_dados['data']):
            novos_dados['data'] = pd.Timestamp('2025-01-01')
            
        # 2. Cálculos automáticos
        novos_dados['comissão_alessandro'] = round(
            novos_dados['valor_liberado'] * (novos_dados['porcentagem_alessandro'] / 100), 2
        )
        novos_dados['valor_dualcred'] = (
            novos_dados['valor_transacionado'] 
            - novos_dados['valor_liberado'] 
            - novos_dados['taxa_de_juros'] 
            - novos_dados['comissão_alessandro'] 
            - novos_dados['extra_alessandro']
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
        salvar_no_excel(df)
        
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
        salvar_no_excel(df)
        
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



if __name__ == "__main__":
    app.run_server(debug=True)