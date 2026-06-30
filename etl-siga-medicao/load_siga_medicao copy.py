import pandas as pd
import openpyxl
import glob
import os
import shutil
import unicodedata
import re

def normalizar_coluna(nome):
    """
    Limpa o nome da coluna seguindo as regras:
    - Tudo em minúsculo
    - Sem acentos
    - Troca 'ç' por 'c'
    - Substitui espaços e símbolos especiais por '_'
    """
    nome = str(nome).strip().lower()
    nome = nome.replace('\n', ' ')
    nome = nome.replace('ç', 'c')
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    nome = re.sub(r'[^a-z0-9]+', '_', nome)
    return nome.strip('_')

def processar_boletins(pasta_origem_input, pasta_destino_input, pasta_backup_input):
    """
    Consolida arquivos do SIGA, normaliza cabeçalhos, segmenta por Mês/Ano,
    audita a volumetria (linhas e colunas) e move origens para Backup.
    """
    pasta_origem = os.path.join(pasta_origem_input, '*.xlsx')
    
    os.makedirs(pasta_destino_input, exist_ok=True)
    os.makedirs(pasta_backup_input, exist_ok=True)
    
    dataframes_limpos = []
    total_linhas_origem = 0
    arquivos_processados_com_sucesso = []

    arquivos = glob.glob(pasta_origem)
    quantidade_arquivos_origem = len(arquivos)
    
    if quantidade_arquivos_origem == 0:
        print(f"\n✅ Nenhum arquivo novo para processar na pasta origem.")
        return

    print(f"Encontrados {quantidade_arquivos_origem} ficheiros SIGA novos. Iniciando a extração:\n")

    # 1. LEITURA E MAPEAMENTO COM NORMALIZAÇÃO
    for indice, caminho_arquivo in enumerate(arquivos, start=1):
        nome_arquivo_atual = os.path.basename(caminho_arquivo)
        
        print(f" [{indice:02d}/{quantidade_arquivos_origem:02d}] Lendo: {nome_arquivo_atual}...", end=" ", flush=True)
        
        try:
            df_bruto = pd.read_excel(caminho_arquivo, header=None, dtype=str, engine='openpyxl')
            
            contagem_nao_nulos = df_bruto.notna().sum(axis=1)
            linha_cabecalho = contagem_nao_nulos.idxmax()
            
            df_tabela = df_bruto.iloc[linha_cabecalho + 1:].copy()
            
            cabecalhos_sujos = df_bruto.iloc[linha_cabecalho]
            df_tabela.columns = [normalizar_coluna(c) for c in cabecalhos_sujos]
            
            df_tabela.dropna(axis=0, how='all', inplace=True)
            
            colunas_validas = [c for c in df_tabela.columns if c not in ['nan', '', 'none', 'na']]
            df_tabela = df_tabela[colunas_validas]
            
            df_tabela = df_tabela.loc[:, ~df_tabela.columns.duplicated(keep='first')]
            
            linhas_neste_arquivo = len(df_tabela)
            total_linhas_origem += linhas_neste_arquivo
            
            df_tabela['arquivo_origem'] = nome_arquivo_atual
            
            dataframes_limpos.append(df_tabela)
            arquivos_processados_com_sucesso.append(caminho_arquivo)
            
            print(f"OK! ({linhas_neste_arquivo} linhas)")
            
        except Exception as e:
            print(f"ERRO! Detalhes: {e}")

    # 2. CONSOLIDAÇÃO E PARTICIONAMENTO (CSV)
    if dataframes_limpos:
        print("\nUnificando os lotes extraídos...")
        df_final = pd.concat(dataframes_limpos, ignore_index=True)
        
        colunas = ['arquivo_origem'] + [c for c in df_final.columns if c != 'arquivo_origem']
        df_final = df_final[colunas]
        
        if 'data' not in df_final.columns:
            df_final['data'] = pd.NA
            
        df_final['_data_dt'] = pd.to_datetime(df_final['data'], dayfirst=True, errors='coerce')
        df_final['_ano'] = df_final['_data_dt'].dt.strftime('%Y')
        df_final['_mes'] = df_final['_data_dt'].dt.strftime('%m')
        
        df_final['_ano'] = df_final['_ano'].fillna('DATA')
        df_final['_mes'] = df_final['_mes'].fillna('INDETERMINADA')
        
        grupos = df_final.groupby(['_ano', '_mes'])
        arquivos_gerados_resumo = []
        
        print("\nAlocando as informações nos consolidados (CSV)...")
        for (_ano, _mes), grupo_dados in grupos:
            df_exportar = grupo_dados.drop(columns=['_data_dt', '_ano', '_mes'])
            
            if _ano == 'DATA' and _mes == 'INDETERMINADA':
                nome_arquivo = 'equipes_stc_siga_consolidado_data_nao_identificada.csv'
            else:
                nome_arquivo = f'equipes_stc_siga_consolidado_{_mes}_{_ano}.csv'
                
            caminho_saida = os.path.join(pasta_destino_input, nome_arquivo)
            
            if os.path.exists(caminho_saida):
                print(f" -> Atualizando base existente: {nome_arquivo}")
                df_existente = pd.read_csv(caminho_saida, sep=';', dtype=str, encoding='utf-8-sig')
                df_salvar = pd.concat([df_existente, df_exportar], ignore_index=True)
            else:
                print(f" -> Criando nova base: {nome_arquivo}")
                df_salvar = df_exportar
            
            df_salvar.to_csv(caminho_saida, sep=';', index=False, encoding='utf-8-sig')
            arquivos_gerados_resumo.append((nome_arquivo, len(df_exportar)))

        # 3. MOVER ARQUIVOS PARA O BACKUP
        print("\nMovendo planilhas de origem para a pasta de backup...")
        for arquivo_path in arquivos_processados_com_sucesso:
            nome_arq = os.path.basename(arquivo_path)
            destino_backup = os.path.join(pasta_backup_input, nome_arq)
            
            if os.path.exists(destino_backup):
                os.remove(destino_backup)
                
            shutil.move(arquivo_path, destino_backup)
            print(f" -> Movido: {nome_arq}")

        # 4. VALIDAÇÃO DE VOLUMETRIA E RESUMO FINAL
        total_linhas_destino = len(df_final)
        total_colunas_destino = len(df_final.columns)

        print("\n" + "="*60)
        print("                 RESUMO DA CARGA INCREMENTAL")
        print("="*60)
        print(f"FICHEIROS DE ORIGEM (.xlsx):")
        print(f" -> Lidos na pasta origem:  {quantidade_arquivos_origem}")
        print(f" -> Movidos para o backup:  {len(arquivos_processados_com_sucesso)}")
        
        print(f"\nAUDITORIA DE DADOS (NOVA CARGA):")
        print(f" -> Colunas Mapeadas:       {total_colunas_destino} campos únicos encontrados")
        print(f" -> Linhas de Origem:       {total_linhas_origem} linhas extraídas")
        print(f" -> Linhas de Destino:      {total_linhas_destino} linhas processadas")
        
        if total_linhas_origem == total_linhas_destino:
            print(" -> Status Integridade:     ✅ SUCESSO! 100% de match nas linhas.")
        else:
            print(f" -> Status Integridade:     ⚠️ ATENÇÃO! Diferença de {total_linhas_origem - total_linhas_destino} linhas.")

        print(f"\nARQUIVOS ATUALIZADOS/CRIADOS NO DESTINO (.csv):")
        for nome_arq, qtd_lin in arquivos_gerados_resumo:
            print(f" -> {nome_arq} (+{qtd_lin} novas linhas)")
        print("="*60)
        
    else:
        print("\n❌ Nenhum dado válido foi extraído das planilhas de origem.")