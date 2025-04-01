import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from dash import dcc
import io

# Configuração de logging detalhada
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuração de caminhos dinâmica
MOUNT_PATH = '/data' if os.environ.get('RENDER') else os.path.join(os.getcwd(), 'data')
EXCEL_PATH = os.path.join(MOUNT_PATH, 'b.xlsx')

def setup_persistent_environment():
    """Configuração robusta do ambiente persistente"""
    try:
        # 1. Cria diretório se não existir
        os.makedirs(MOUNT_PATH, exist_ok=True)
        logger.info(f"Diretório verificado: {MOUNT_PATH}")

        # 2. Cria arquivo Excel inicial com estrutura completa
        if not os.path.exists(EXCEL_PATH):
            logger.info("Criando novo arquivo Excel com estrutura completa...")
            
            wb = Workbook()
            meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN',
                    'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
            
            # Cabeçalho completo com todas colunas necessárias
            colunas_obrigatorias = [
                'data', 'agente', 'valor_transacionado', 'valor_liberado',
                'taxa_de_juros', 'comissão_agente', 'extra_agente',
                'quantidade_parcelas'
            ]
            
            # Remove sheet padrão
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            
            # Cria abas com estrutura completa
            for mes in meses:
                ws = wb.create_sheet(mes)
                ws.append(colunas_obrigatorias)
            
            wb.save(EXCEL_PATH)
            logger.info(f"Arquivo inicial criado: {EXCEL_PATH}")

        # 3. Verifica permissões
        if not os.access(MOUNT_PATH, os.W_OK):
            logger.error(f"Sem permissão de escrita em: {MOUNT_PATH}")
            raise PermissionError("Erro de permissão no diretório persistente")

    except Exception as e:
        logger.error(f"Falha na configuração inicial: {str(e)}")
        raise

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
    """Carrega e processa dados com validação completa"""
    try:
        setup_persistent_environment()
        logger.info("Iniciando processamento de dados...")

        # 1. Carregar dados brutos
        sheets_to_read = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN',
                         'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
        
        df_list = []
        for sheet in sheets_to_read:
            try:
                df_sheet = pd.read_excel(
                    EXCEL_PATH,
                    sheet_name=sheet,
                    engine='openpyxl',
                    dtype=str
                )
                
                # Processamento inicial
                df_sheet = df_sheet.dropna(axis=1, how='all')
                df_sheet.columns = [sanitize_column_name(col) for col in df_sheet.columns]
                df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]
                
                df_list.append(df_sheet)
                logger.debug(f"Aba {sheet} carregada: {len(df_sheet)} registros")
                
            except Exception as e:
                logger.warning(f"Erro na aba {sheet}: {str(e)}")
                continue

        df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

        # 2. Validação de colunas obrigatórias
        colunas_obrigatorias = [
            'data', 'agente', 'valor_transacionado', 'valor_liberado',
            'taxa_de_juros', 'comissão_agente', 'extra_agente'
        ]
        
        for col in colunas_obrigatorias:
            if col not in df.columns:
                logger.warning(f"Coluna {col} não encontrada. Criando com valores padrão.")
                df[col] = None if col == 'data' else 0.0
                if col == 'agente':
                    df[col] = 'Alessandro'

        # 3. Processamento de datas
        df['data'] = pd.to_datetime(
            df['data'],
            errors='coerce',
            dayfirst=True,
            format='mixed'
        ).fillna(pd.to_datetime('2025-01-01'))  # Data padrão segura

        # 4. Conversão numérica segura
        numeric_cols = [
            'valor_transacionado', 'valor_liberado', 'taxa_de_juros',
            'comissão_agente', 'extra_agente', 'quantidade_parcelas'
        ]
        
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)
            else:
                df[col] = 0.0

        # 5. Cálculos derivados com verificação
        df['valor_dualcred'] = (
            df['valor_transacionado'] 
            - df['valor_liberado'] 
            - df['taxa_de_juros'] 
            - df['comissão_agente'] 
            - df['extra_agente']
        ).round(2)
        
        df['%trans'] = np.where(
            df['valor_transacionado'] > 0,
            (df['valor_dualcred'] / df['valor_transacionado']) * 100,
            0
        ).round(2)
        
        df['%liberad'] = np.where(
            df['valor_liberado'] > 0,
            (df['valor_dualcred'] / df['valor_liberado']) * 100,
            0
        ).round(2)
        
        df['nota_fiscal'] = (df['valor_transacionado'] * 0.032).round(2)

        logger.info("Processamento concluído com sucesso")
        return df

    except Exception as e:
        logger.error(f"Erro crítico no processamento: {str(e)}", exc_info=True)
        return pd.DataFrame()


def salvar_no_excel(df):
    """Salvamento otimizado para ambiente persistente"""
    try:
        logger.info("Iniciando salvamento persistente...")
        
        setup_persistent_environment()
        
        meses = {
            1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR',
            5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO',
            9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
        }
        
        # Modo 'a' para append mantendo a estrutura existente
        with pd.ExcelWriter(
            EXCEL_PATH,
            engine='openpyxl',
            mode='a',
            if_sheet_exists='replace'
        ) as writer:
            
            for month_num, sheet_name in meses.items():
                month_df = df[df['data'].dt.month == month_num].copy()
                if not month_df.empty:
                    month_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False,
                        header=True
                    )
                    logger.debug(f"Aba {sheet_name} atualizada com {len(month_df)} registros")
        
        logger.info("Dados salvos com sucesso no armazenamento persistente")
        return True
    
    except Exception as e:
        logger.error(f"Erro no salvamento: {str(e)}")
        return False

def exportar_dados(filtered_df):
    """Exportação para download sem afetar o arquivo persistente"""
    try:
        meses = {
            1: 'JAN', 2: 'FEV', 3: 'MAR', 4: 'ABR',
            5: 'MAI', 6: 'JUN', 7: 'JUL', 8: 'AGO',
            9: 'SET', 10: 'OUT', 11: 'NOV', 12: 'DEZ'
        }
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            for month_num, sheet_name in meses.items():
                month_df = filtered_df[filtered_df['data'].dt.month == month_num].copy()
                month_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        buffer.seek(0)
        return dcc.send_bytes(
            buffer.getvalue(),
            filename="Dados_Exportados.xlsx",
            type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as e:
        logger.error(f"Erro na exportação: {str(e)}")
        return None

# Inicialização segura
try:
    df = load_and_process_data()
    if df.empty:
        logger.warning("DataFrame inicial vazio - possíveis dados ausentes")
except Exception as e:
    logger.error(f"Falha na inicialização: {str(e)}")
    df = pd.DataFrame()