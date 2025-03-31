import dash
from dash import dcc
import pandas as pd
import numpy as np
from datetime import datetime
import io
global df

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

def load_and_process_data():
    """Função única para carregar e pré-processar todos os dados"""
    # 1. Carregar dados brutos
    file_path = 'b.xlsx'
    sheets_to_read = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
                     'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
    
    df_list = []
    for sheet in sheets_to_read:
        df_sheet = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
        df_sheet = df_sheet.dropna(axis=1, how='all')
        df_sheet.columns = [sanitize_column_name(col) for col in df_sheet.columns]
        df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]
        df_list.append(df_sheet)
    
    df = pd.concat(df_list, ignore_index=True)

    # 2. Pré-processamento básico
    if 'agente' not in df.columns:
        df['agente'] = 'Alessandro'

    if 'qtd_parcelas' in df.columns:
        if 'quantidade_parcelas' in df.columns:
            df['quantidade_parcelas'] = df['quantidade_parcelas'].combine_first(df['qtd_parcelas'])
            df.drop('qtd_parcelas', axis=1, inplace=True)
        else:
            df.rename(columns={'qtd_parcelas': 'quantidade_parcelas'}, inplace=True)

    # 3. Limpeza de colunas
    excluir_inicial = [
        '%trans', '%liberad', 'acerto_alessandro', 
        'retirada_felipe', 'máquina', 'acerto_alesandro'
    ]
    df = df.drop(columns=excluir_inicial, errors='ignore')

    # 4. Renomeação de colunas
    renomear = {
        'comissão_alessandro': 'comissão_agente',
        'extra_alessandro': 'extra_agente',
        'porcentagem_alessandro': 'porcentagem_agente'
    }
    df.rename(columns=renomear, inplace=True)

    # 5. Tipos de dados
    df['data'] = pd.to_datetime(df['data'], errors='coerce').fillna(pd.to_datetime('2025-01-01'))
    
    numeric_cols = [
        'valor_transacionado', 'valor_liberado', 'taxa_de_juros',
        'comissão_agente', 'extra_agente', 'porcentagem_agente',
        'nota_fiscal', 'quantidade_parcelas'
    ]
    
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)

    # 6. Cálculos derivados
    def calcular_valor_dualcred(row):
        return (
            row['valor_transacionado']
            - row['valor_liberado']
            - row['taxa_de_juros']
            - row['comissão_agente']
            - row['extra_agente']
        )
    
    df['valor_dualcred'] = df.apply(calcular_valor_dualcred, axis=1).round(2)
    
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
    
    df['nota_fiscal'] = (df['valor_transacionado'] * 0.032).round(2)

    # 7. Exclusão final de colunas
    excluir_final = ['%_trans.', '%_liberad.', 'acerto_alessandro', 'retirada_felipe', 'máquina']
    df = df.drop(columns=excluir_final, errors='ignore')

    return df

# Funções auxiliares separadas
def salvar_no_excel(df):
    meses = {1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN', 
            7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'}
    try:
        with pd.ExcelWriter('b.xlsx', engine='openpyxl') as writer:
            for month_num, sheet_name in meses.items():
                month_df = df[df['data'].dt.month == month_num].copy()
                month_df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"Erro ao salvar: {str(e)}")

def exportar_dados(filtered_df):
    meses = {1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR', 5: 'MAI', 6: 'JUN', 
            7: 'JUL', 8: 'AGO', 9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'}
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            for month_num, sheet_name in meses.items():
                month_df = filtered_df[filtered_df['data'].dt.month == month_num].copy()
                month_df.to_excel(writer, sheet_name=sheet_name, index=False)
            buffer.seek(0)
        return dcc.send_bytes(
            buffer.getvalue(), 
            filename="Nome_diferente.xlsx",
            type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        print(f"Erro na exportação: {str(e)}")
        return None
df = load_and_process_data()