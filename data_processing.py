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
        os.makedirs(MOUNT_PATH, exist_ok=True)
        logger.info(f"Diretório verificado: {MOUNT_PATH}")

        # Verifica se o arquivo existe sem criar um novo
        if not os.path.exists(EXCEL_PATH):
            logger.warning("Arquivo Excel não encontrado. Criando novo...")
            wb = Workbook()
            del wb['Sheet']
            wb.save(EXCEL_PATH)
        
        # Verifica permissões
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
        .replace("ç", "c")
        .replace("ã", "a")
        .replace("õ", "o")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("à", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("ú", "u")
        .replace("%", "porcento")
        .replace("(", "")
        .replace(")", "")
    )

def load_and_process_data():
    """Carrega e processa dados mantendo a estrutura original"""
    try:
        setup_persistent_environment()
        logger.info("Iniciando processamento de dados...")

        # Mapeamento de colunas esperadas
        column_mapping = {
            'data': 'data',
            'agente': 'agente',
            'beneficiario': 'beneficiario',
            'valor_transacionado': 'valor_transacionado',
            'valor_liberado': 'valor_liberado',
            'taxa_de_juros': 'taxa_de_juros',
            'comissão_agente': 'comissao_agente',
            'extra_agente': 'extra_agente',
            'valor_dualcred': 'valor_dualcred',
            'nota_fiscal': 'nota_fiscal',
            'porcentagem_agente': 'porcentagem_agente',
            'quantidade_parcelas': 'quantidade_parcelas'
        }

        # Carregar todas as abas
        sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, engine='openpyxl')
        df_list = []

        for sheet_name, df in sheets.items():
            try:
                # Sanitizar colunas
                df.columns = [sanitize_column_name(col) for col in df.columns]
                
                # Renomear colunas
                df.rename(columns=column_mapping, inplace=True, errors='ignore')
                
                # Adicionar colunas faltantes
                for col in column_mapping.values():
                    if col not in df.columns:
                        df[col] = np.nan if col == 'data' else 0.0
                
                df_list.append(df)
                logger.debug(f"Aba {sheet_name} processada com {len(df)} registros")

            except Exception as e:
                logger.warning(f"Erro na aba {sheet_name}: {str(e)}")
                continue

        df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

        # Processamento de datas
        df['data'] = pd.to_datetime(
            df['data'],
            errors='coerce',
            dayfirst=True
        ).fillna(pd.to_datetime('2025-01-01'))

        # Conversão numérica
        numeric_cols = [
            'valor_transacionado', 'valor_liberado', 'taxa_de_juros',
            'comissao_agente', 'extra_agente', 'valor_dualcred',
            'nota_fiscal', 'porcentagem_agente', 'quantidade_parcelas'
        ]
        
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)

        # Cálculos condicionais
        if 'valor_dualcred' not in df.columns:
            df['valor_dualcred'] = (
                df['valor_transacionado'] 
                - df['valor_liberado'] 
                - df['taxa_de_juros'] 
                - df['comissao_agente'] 
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

        if 'nota_fiscal' not in df.columns:
            df['nota_fiscal'] = (df['valor_transacionado'] * 0.032).round(2)

        logger.info(f"Dados carregados: {len(df)} registros")
        return df

    except Exception as e:
        logger.error(f"Erro crítico no processamento: {str(e)}", exc_info=True)
        return pd.DataFrame()

def salvar_no_excel(df):
    """Salvamento incremental mantendo dados existentes"""
    try:
        logger.info("Iniciando salvamento persistente...")
        setup_persistent_environment()

        # Carregar dados existentes
        existing_data = pd.read_excel(EXCEL_PATH, sheet_name=None, engine='openpyxl')
        
        with pd.ExcelWriter(
            EXCEL_PATH,
            engine='openpyxl',
            mode='w'  # Modo de escrita completo para evitar corrupção
        ) as writer:
            for sheet_name in existing_data.keys():
                # Combinar dados existentes com novos
                sheet_df = existing_data[sheet_name]
                new_data = df[df['data'].dt.strftime('%b').str.upper() == sheet_name]
                combined_df = pd.concat([sheet_df, new_data]).drop_duplicates()
                
                # Manter ordem original das colunas
                cols_order = [
                    'data', 'beneficiario', 'valor_transacionado', 'valor_liberado',
                    'taxa_de_juros', 'comissao_agente', 'extra_agente', 'valor_dualcred',
                    'nota_fiscal', 'porcentagem_agente', 'quantidade_parcelas', 'agente',
                    '%trans', '%liberad'
                ]
                
                combined_df = combined_df.reindex(columns=cols_order)
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                logger.debug(f"Aba {sheet_name} salva com {len(combined_df)} registros")

        logger.info("Dados salvos com sucesso")
        return True

    except Exception as e:
        logger.error(f"Erro no salvamento: {str(e)}")
        return False

def exportar_dados(filtered_df):
    """Exportação para download mantendo estrutura original"""
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            for sheet_name in ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN',
                              'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']:
                month_df = filtered_df[filtered_df['data'].dt.strftime('%b').str.upper() == sheet_name]
                if not month_df.empty:
                    month_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False,
                        columns=[
                            'data', 'beneficiario', 'valor_transacionado', 'valor_liberado',
                            'taxa_de_juros', 'comissao_agente', 'extra_agente', 'valor_dualcred',
                            'nota_fiscal', 'quantidade_parcelas', 'agente', '%trans', '%liberad'
                        ]
                    )
        
        buffer.seek(0)
        return dcc.send_bytes(
            buffer.getvalue(),
            filename="Dados_Atualizados.xlsx",
            type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as e:
        logger.error(f"Erro na exportação: {str(e)}")
        return None

# Inicialização segura
try:
    df = load_and_process_data()
    if df.empty:
        logger.warning("DataFrame inicial vazio - verifique o arquivo fonte")
    else:
        logger.info(f"Dados iniciais carregados: {df.shape}")
except Exception as e:
    logger.error(f"Falha crítica na inicialização: {str(e)}")
    df = pd.DataFrame()