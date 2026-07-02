import unicodedata
import re
import shutil
import warnings
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import openpyxl

# Ignora avisos inofensivos do Pandas sobre formatações do Excel
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def normalizar_coluna(nome: str) -> str:
    if pd.isna(nome):
        return ""
        
    nome = str(nome).strip().lower()
    nome = nome.replace('\n', ' ').replace('ç', 'c')
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    nome = re.sub(r'[^a-z0-9]+', '_', nome)
    return nome.strip('_')

def _extrair_tabela_siga(df_bruto: pd.DataFrame) -> pd.DataFrame:
    linha_cabecalho = -1
    
    for idx, row in df_bruto.iterrows():
        linha_texto = ' '.join(row.fillna('').astype(str)).lower()
        if 'ordem de serviço' in linha_texto and 'status da atividade' in linha_texto:
            linha_cabecalho = idx
            break

    if linha_cabecalho == -1:
        return pd.DataFrame()

    df_tabela = df_bruto.iloc[linha_cabecalho:].copy()
    colunas_originais = df_tabela.iloc[0].astype(str).tolist()
    df_tabela.columns = [normalizar_coluna(col) for col in colunas_originais]
    
    df_tabela = df_tabela.iloc[1:].copy()
    df_tabela.dropna(axis=1, how='all', inplace=True)
    df_tabela.dropna(axis=0, how='all', inplace=True)
    
    return df_tabela

def _limpar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    colunas_validas = [
        col for col in df.columns 
        if not str(col).lower().startswith('unnamed') 
        and str(col).lower() not in ['nan', 'none', '', '<na>']
    ]
    df = df[colunas_validas]
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def processar_boletins(pasta_origem_input: str, pasta_destino_input: str, pasta_backup_input: str) -> None:
    print("\n" + "="*60)
    print("INICIANDO PROCESSAMENTO DE DADOS - SIGA (SEGMENTADO)".center(60))
    print("="*60)

    pasta_origem = Path(pasta_origem_input)
    pasta_destino = Path(pasta_destino_input)
    pasta_backup = Path(pasta_backup_input)

    pasta_destino.mkdir(parents=True, exist_ok=True)
    pasta_backup.mkdir(parents=True, exist_ok=True)

    arquivos = [
        f for f in pasta_origem.glob('*.xlsx') 
        if not f.name.startswith('~$') and "siga_consolidado_" not in f.name
    ]

    if not arquivos:
        print("Nenhum ficheiro pendente de processamento encontrado.")
        return

    print(f"Encontrados {len(arquivos)} ficheiro(s) SIGA válido(s). A extrair dados...\n")

    dataframes_limpos = []
    total_linhas_origem = 0
    arquivos_processados_com_sucesso = []

    # ---------------------------------------------------------
    # FASE 1: LEITURA E EXTRAÇÃO
    # ---------------------------------------------------------
    for caminho_arquivo in arquivos:
        nome_arquivo = caminho_arquivo.name
        teve_sucesso = False
        
        try:
            with open(caminho_arquivo, "rb") as f:
                xls = pd.ExcelFile(f, engine='openpyxl')
                abas = xls.sheet_names

            for aba in abas:
                # DEFESA 1: Garante que a aba existe no workbook antes de a procurar
                wb = openpyxl.load_workbook(caminho_arquivo, read_only=True)
                is_visible = False
                if aba in wb.sheetnames:
                    is_visible = wb[aba].sheet_state == 'visible'
                wb.close()

                if not is_visible:
                    continue

                with open(caminho_arquivo, "rb") as f:
                    df_bruto = pd.read_excel(f, sheet_name=aba, header=None, dtype=str, engine='openpyxl')

                df_tabela = _extrair_tabela_siga(df_bruto)
                
                if df_tabela.empty:
                    continue

                df_tabela = _limpar_dataframe(df_tabela)
                df_tabela['arquivo_origem'] = nome_arquivo

                total_linhas_origem += len(df_tabela)
                dataframes_limpos.append(df_tabela)
                teve_sucesso = True

            if teve_sucesso:
                arquivos_processados_com_sucesso.append(caminho_arquivo)
                print(f" -> LIDO: {nome_arquivo}")

        except Exception as e:
            print(f"[!] ERRO ao processar '{nome_arquivo}': {e}")

    if not dataframes_limpos:
        print("\n[!] Nenhum dado válido foi encontrado no interior dos ficheiros.")
        return

    # ---------------------------------------------------------
    # FASE 2: CONSOLIDAÇÃO E LIMPEZA FINAL
    # ---------------------------------------------------------
    print("\nA unir e normalizar toda a base de dados...")
    df_final = pd.concat(dataframes_limpos, ignore_index=True)
    df_final = _limpar_dataframe(df_final)

    colunas_obrigatorias = [
        'arquivo_origem', 'data', 'ordem_de_servico', 'origem', 'id_da_atividade', 
        'status_da_atividade', 'numero_da_nota', 'numero_ocorrencia', 'numero_da_conta', 
        'agrupamento_id', 'code', 'cidade', 'bairro', 'latitude', 'longitude'
    ]
    
    col_presentes = [col for col in colunas_obrigatorias if col in df_final.columns]
    col_extras = [col for col in df_final.columns if col not in colunas_obrigatorias]
    
    df_final = df_final[col_presentes + col_extras]

    # ---------------------------------------------------------
    # FASE 3: SEGMENTAÇÃO POR MÊS/ANO E EXPORTAÇÃO (CSV)
    # ---------------------------------------------------------
    print("A analisar datas e a criar ficheiros segmentados...")
    
    if 'data' in df_final.columns:
        df_final['Mes_Ano'] = pd.to_datetime(df_final['data'], format='mixed', dayfirst=True, errors='coerce').dt.strftime('%m-%Y')
        df_final['Mes_Ano'] = df_final['Mes_Ano'].fillna('Sem_Data')
    else:
        df_final['Mes_Ano'] = 'Sem_Data'

    meses_unicos = df_final['Mes_Ano'].unique()
    data_hoje = datetime.now().strftime('%Y_%m_%d')

    for mes in meses_unicos:
        df_mes = df_final[df_final['Mes_Ano'] == mes].drop(columns=['Mes_Ano'])
        
        # Alterado para extensão .csv
        nome_arquivo_saida = f"siga_consolidado_{mes}_{data_hoje}.csv"
        caminho_saida = pasta_destino / nome_arquivo_saida
        
        print(f" -> A guardar {nome_arquivo_saida} ({len(df_mes)} linhas)...")
        
        # DEFESA 2: Apaga o ficheiro antigo se ele existir para evitar bugs de sobrescrita
        if caminho_saida.exists():
            try:
                caminho_saida.unlink()
            except PermissionError:
                print(f"\n[ERRO] O ficheiro {nome_arquivo_saida} está aberto. Feche-o!")
                continue

        # Exporta como CSV utilizando to_csv em vez de to_excel
        try:
            df_mes.to_csv(caminho_saida, index=False, sep=';', encoding='utf-8-sig')
        except Exception as e:
            print(f"[!] Erro crítico ao guardar {nome_arquivo_saida}: {e}")

    # ---------------------------------------------------------
    # FASE 4: BACKUP E AUDITORIA (COMPACTAÇÃO ZIP)
    # ---------------------------------------------------------
    print("\nA compactar e mover ficheiros processados para Backup...")
    if arquivos_processados_com_sucesso:
        # Formato YYYYMMDDHHMMSS para garantir nome único
        timestamp_zip = datetime.now().strftime('%Y%m%d%H%M%S')
        nome_zip = f"arquivo-processado_{timestamp_zip}.zip"
        caminho_zip = pasta_backup / nome_zip

        try:
            # Abre/Cria o arquivo ZIP no modo de compressão
            with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for caminho in arquivos_processados_com_sucesso:
                    # Adiciona o arquivo dentro do ZIP usando apenas o nome dele (arcname)
                    zipf.write(caminho, arcname=caminho.name)
                    # Exclui o arquivo original da pasta source após compactá-lo com sucesso
                    caminho.unlink()
            print(f" -> ✅ Arquivos compactados com sucesso em: {nome_zip}")
        except Exception as e:
            print(f" -> [!] Erro ao criar o arquivo ZIP de backup: {e}")

    total_linhas_destino = len(df_final)
    
    print("\n" + "="*60)
    print("RESUMO DA CARGA INCREMENTAL".center(60))
    print("="*60)
    print(f"FICHEIROS DE ORIGEM (.xlsx):")
    print(f" -> Lidos na pasta origem:  {len(arquivos)}")
    print(f" -> Compactados p/ backup:  {len(arquivos_processados_com_sucesso)}")
    
    print(f"\nAUDITORIA DE DADOS:")
    print(f" -> Colunas Mapeadas:       {len(df_final.columns) - 1} campos únicos encontrados")
    print(f" -> Linhas Extraídas:       {total_linhas_origem}")
    print(f" -> Linhas Processadas:     {total_linhas_destino}")
    
    if total_linhas_origem == total_linhas_destino:
        print(" -> Status Integridade:     ✅ SUCESSO! 100% de match nas linhas.")
    else:
        print(f" -> Status Integridade:     ⚠️ ATENÇÃO! Diferença de {total_linhas_origem - total_linhas_destino} linhas.")
    print("="*60)