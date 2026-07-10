"""
load_medicao.py
----------------
ETL para leitura, tratamento e consolidação de múltiplas planilhas de
medição (.xlsx) de uma pasta de origem, gerando um CSV consolidado e
fazendo backup (zip) dos arquivos lidos com sucesso.

Autor: Engenharia de Dados
"""

import re
import unicodedata
import warnings
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import openpyxl

# Ignora avisos inofensivos do openpyxl (estilos/extensões não suportadas etc.)
warnings.filterwarnings("ignore", module="openpyxl")

# =====================================================================
# CONFIGURAÇÃO DE MAPEAMENTO DE COLUNAS (DE-PARA)
# =====================================================================
# Chave = nome da coluna já normalizado (minúsculo, sem acento, "_" no
# lugar de espaço/pontuação). Valor = nome padrão no dataframe final.
MAPEAMENTO_COLUNAS = {
    'mes_ano': 'data',                     # MÊS/ANO
    'data': 'data',                        # DATA
    'data_envio': 'data',                  # DATA ENVIO
    'tipo_de_servico': 'servico',          # TIPO DE SERVIÇO
    'servico': 'servico',                  # SERVIÇO
    'descricao_servico': 'servico',        # DESCRIÇÃO SERVIÇO
    'os': 'ordem_servico',                 # OS
    'ordem_de_servico': 'ordem_servico',   # ORDEM DE SERVIÇO
    'ordem_servico': 'ordem_servico',      # ORDEM SERVIÇO
    'n_da_nota': 'numero_nota',            # Nº DA NOTA
    'nota': 'numero_nota',                 # NOTA
    'numero_nota': 'numero_nota',          # NÚMERO NOTA
    'equipe': 'equipe',                    # EQUIPE
    'municipio': 'cidade',                 # MUNICÍPIO
    'localidade': 'cidade',                # LOCALIDADE
    'cidade': 'cidade',                    # CIDADE
    'qtd': 'quantidade',                   # QTD
    'quantidade': 'quantidade',            # QUANTIDADE
    'valor_r_': 'valor',                   # VALOR (R$)
    'valor': 'valor',                      # VALOR
    'preco_unitario': 'valor_unitario',    # PREÇO UNITÁRIO
    'status_da_atividade': 'status',       # STATUS DA ATIVIDADE
    'status': 'status',                    # STATUS
    'observacoes': 'observacao',           # OBSERVAÇÕES
    'observacao': 'observacao',            # OBSERVAÇÃO
    'fornecedor': 'fornecedor',            # FORNECEDOR
    'responsavel': 'responsavel',          # RESPONSÁVEL
    'responsavel_validacao': 'responsavel',
}

# Colunas mapeadas que devem ficar na frente do dataframe final, nessa ordem.
ORDEM_PADRAO = [
    'arquivo_origem',
    'data',
    'ordem_servico',
    'numero_nota',
    'equipe',
    'cidade',
    'fornecedor',
    'responsavel',
    'servico',
    'quantidade',
    'valor_unitario',
    'valor',
    'status',
    'observacao',
]

# Colunas numéricas conhecidas: convertidas para float com 2 casas decimais
# para evitar artefatos de ponto flutuante do Excel (ex: "47256.299999999996")
# quando o valor é serializado como texto no CSV.
COLUNAS_NUMERICAS = ['quantidade', 'valor_unitario', 'valor']

# Palavras-chave usadas para localizar dinamicamente a linha de cabeçalho
# real da tabela, ignorando logos/textos soltos no topo da planilha.
PALAVRAS_CHAVE_CABECALHO = [
    'nota', 'ordem', 'data', 'equipe', 'valor', 'status',
    'municipio', 'localidade', 'servico',
]

MAX_LINHAS_VARREDURA_CABECALHO = 30  # segurança: não varre a planilha inteira


# =====================================================================
# NORMALIZAÇÃO
# =====================================================================

def normalizar_coluna(nome) -> str:
    """Normaliza nome de coluna: minúsculo, sem acento, sem caractere especial."""
    if pd.isna(nome):
        return ""
    nome = str(nome).strip().lower()
    nome = nome.replace('\n', ' ').replace('ç', 'c')
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    nome = re.sub(r'[^a-z0-9]+', '_', nome)
    return nome.strip('_')


# =====================================================================
# EXTRAÇÃO DA TABELA (CABEÇALHO DINÂMICO + REMOÇÃO DE TOTAL)
# =====================================================================

def _localizar_linha_cabecalho(df_bruto: pd.DataFrame) -> int:
    """Varre as primeiras linhas procurando por palavras-chave que indiquem
    o cabeçalho real da tabela (ignora logos/textos soltos no topo)."""
    limite = min(len(df_bruto), MAX_LINHAS_VARREDURA_CABECALHO)

    for idx in range(limite):
        row = df_bruto.iloc[idx]
        linha_normalizada = normalizar_coluna(' '.join(row.fillna('').astype(str)))
        matches = sum(1 for kw in PALAVRAS_CHAVE_CABECALHO if kw in linha_normalizada)
        if matches >= 2:
            return idx

    # Fallback: primeira linha com um número razoável de células preenchidas
    for idx in range(limite):
        if df_bruto.iloc[idx].notna().sum() > 4:
            return idx

    return -1


def _remover_linha_de_total(df_tabela: pd.DataFrame) -> pd.DataFrame:
    """Remove a última linha se parecer ser um total (texto 'total' presente
    ou maioria das células em branco)."""
    if df_tabela.empty:
        return df_tabela

    ultima_linha = df_tabela.iloc[-1]
    texto_ultima_linha = ' '.join(ultima_linha.fillna('').astype(str)).lower()

    e_total_textual = 'total' in texto_ultima_linha
    e_maioria_vazia = ultima_linha.isna().sum() > (len(df_tabela.columns) / 2)

    if e_total_textual or e_maioria_vazia:
        return df_tabela.iloc[:-1]
    return df_tabela


def _consolidar_colunas_duplicadas(colunas_normalizadas: list, contexto: str, log_avisos: list) -> list:
    """Se duas colunas NOMEADAS normalizarem para o mesmo texto (ex.: duas
    colunas 'Processo' na planilha), renomeia a partir da segunda ocorrência
    em vez de deixar o pandas colidir/descartar dado silenciosamente.
    Colunas sem nome (célula de cabeçalho vazia) não geram aviso aqui —
    são descartadas depois, em _extrair_tabela."""
    vistos = {}
    resultado = []
    for col in colunas_normalizadas:
        if not col:
            resultado.append("")
            continue
        if col in vistos:
            vistos[col] += 1
            novo_nome = f"{col}_dup{vistos[col]}"
            log_avisos.append(f"[AVISO] {contexto}: coluna '{col}' duplicada — renomeada para '{novo_nome}'.")
            resultado.append(novo_nome)
        else:
            vistos[col] = 0
            resultado.append(col)
    return resultado


def _extrair_tabela(df_bruto: pd.DataFrame, contexto: str, log_avisos: list) -> pd.DataFrame:
    """Extrai a tabela de dados de um dataframe bruto (sem cabeçalho definido),
    localizando o cabeçalho real e removendo a linha de total."""
    linha_cabecalho = _localizar_linha_cabecalho(df_bruto)
    if linha_cabecalho == -1:
        log_avisos.append(f"[SKIP] {contexto}: cabeçalho não identificado.")
        return pd.DataFrame()

    df_tabela = df_bruto.iloc[linha_cabecalho:].reset_index(drop=True)
    colunas_originais = df_tabela.iloc[0].tolist()
    colunas_normalizadas = [normalizar_coluna(c) for c in colunas_originais]
    colunas_normalizadas = _consolidar_colunas_duplicadas(colunas_normalizadas, contexto, log_avisos)

    df_tabela = df_tabela.iloc[1:].copy()
    df_tabela.columns = colunas_normalizadas

    # Descarta colunas sem cabeçalho (célula em branco no topo). Seleção
    # POSICIONAL (iloc), nunca por nome: se algum nome ainda estiver
    # duplicado, `df[[nome, nome]]` faz o pandas devolver colunas em
    # dobro (expansão combinatória) em vez de selecionar 1:1.
    posicoes_validas = [i for i, c in enumerate(df_tabela.columns) if c != ""]
    df_tabela = df_tabela.iloc[:, posicoes_validas]

    df_tabela = df_tabela.dropna(axis=1, how='all')
    df_tabela = df_tabela.dropna(axis=0, how='all')

    df_tabela = _remover_linha_de_total(df_tabela)

    return df_tabela


# =====================================================================
# TIPAGEM
# =====================================================================

def _tratar_tipos(df: pd.DataFrame) -> pd.DataFrame:
    """Converte colunas numéricas conhecidas para float (2 casas decimais),
    evitando gravar artefato de ponto flutuante como string no CSV."""
    for col in COLUNAS_NUMERICAS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(2)
    return df


# =====================================================================
# LEITURA DE ARQUIVO (1 ÚNICA ABERTURA DE WORKBOOK)
# =====================================================================

def _ler_abas_validas(caminho_arquivo: Path):
    """Abre o workbook uma única vez e retorna apenas a PRIMEIRA aba visível
    do arquivo — ex.: ENLISA (Nordeste BT): pega 'demandas mensais' e ignora
    'Planilha1'; Normalização Centro: pega 'ELINSA.NORMALIZAÇÃO JUNHO -2026'
    e ignora tanto a aba oculta 'FEVEREIRO - 26' quanto 'JUN - 26'.
    Abas ocultas nunca contam como "primeira aba" — só concorrem entre si
    as abas com sheet_state == 'visible'."""
    wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
    resultado = []
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible':
                continue
            valores = list(ws.values)
            if not valores:
                continue
            df_bruto = pd.DataFrame(valores)
            resultado.append((ws.title, df_bruto))
            break  # só a primeira aba visível encontrada
    finally:
        wb.close()
    return resultado


# =====================================================================
# PIPELINE PRINCIPAL
# =====================================================================

def processar_boletins(pasta_origem_input: str, pasta_destino_input: str, pasta_backup_input: str) -> None:
    print("\n" + "=" * 60)
    print("INICIANDO PROCESSAMENTO DE DADOS - MEDIÇÃO".center(60))
    print("=" * 60)

    pasta_origem = Path(pasta_origem_input)
    pasta_destino = Path(pasta_destino_input)
    pasta_backup = Path(pasta_backup_input)

    pasta_destino.mkdir(parents=True, exist_ok=True)
    pasta_backup.mkdir(parents=True, exist_ok=True)

    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: pasta de origem não encontrada:\n{pasta_origem}")
        return

    arquivos = [
        f for f in pasta_origem.glob('*.xlsx')
        if not f.name.startswith('~$') and "medicao_consolidada_" not in f.name
    ]

    if not arquivos:
        print("Nenhum ficheiro pendente de processamento encontrado na origem.")
        return

    print(f"Encontrados {len(arquivos)} ficheiro(s) válido(s). A extrair dados...\n")

    dataframes_limpos = []
    total_linhas_origem = 0
    arquivos_processados_com_sucesso = []
    log_avisos = []

    # ---------------------------------------------------------
    # FASE 1: LEITURA E EXTRAÇÃO
    # ---------------------------------------------------------
    for caminho_arquivo in arquivos:
        nome_arquivo = caminho_arquivo.name
        teve_sucesso = False

        try:
            abas = _ler_abas_validas(caminho_arquivo)

            for nome_aba, df_bruto in abas:
                contexto = f"'{nome_arquivo}' / aba '{nome_aba}'"
                try:
                    df_tabela = _extrair_tabela(df_bruto, contexto, log_avisos)
                except Exception as e:
                    log_avisos.append(f"[ERRO] {contexto}: falha ao extrair tabela ({e}).")
                    continue

                if df_tabela.empty:
                    continue

                df_tabela = df_tabela.copy()
                df_tabela['arquivo_origem'] = nome_arquivo

                total_linhas_origem += len(df_tabela)
                dataframes_limpos.append(df_tabela)
                teve_sucesso = True

            if teve_sucesso:
                arquivos_processados_com_sucesso.append(caminho_arquivo)
                print(f" -> LIDO: {nome_arquivo}")
            else:
                print(f" -> [!] Nenhuma tabela válida encontrada em: {nome_arquivo}")

        except Exception as e:
            print(f"[!] ERRO ao processar '{nome_arquivo}': {e}")

    if log_avisos:
        print("\nAvisos durante a extração:")
        for aviso in log_avisos:
            print(f"  {aviso}")

    if not dataframes_limpos:
        print("\n[!] Nenhum dado válido foi encontrado no interior dos ficheiros.")
        return

    # ---------------------------------------------------------
    # FASE 2: CONSOLIDAÇÃO E MAPEAMENTO (DE-PARA E ORDENAÇÃO)
    # ---------------------------------------------------------
    print("\nA unir e mapear colunas da base de dados...")

    for i in range(len(dataframes_limpos)):
        dataframes_limpos[i] = dataframes_limpos[i].rename(columns=MAPEAMENTO_COLUNAS)

    df_final = pd.concat(dataframes_limpos, ignore_index=True, sort=False)
    df_final = _tratar_tipos(df_final)

    colunas_presentes = set(df_final.columns)
    colunas_finais = [col for col in ORDEM_PADRAO if col in colunas_presentes]
    colunas_extras = sorted(colunas_presentes - set(colunas_finais))
    colunas_finais.extend(colunas_extras)

    if 'arquivo_origem' in colunas_finais:
        colunas_finais.remove('arquivo_origem')
        colunas_finais.insert(0, 'arquivo_origem')

    df_final = df_final[colunas_finais]

    # ---------------------------------------------------------
    # FASE 3: EXPORTAÇÃO UNIFICADA (CSV)
    # ---------------------------------------------------------
    print("A exportar base de dados unificada...")

    timestamp_exportacao = datetime.now().strftime('%Y%m%d%H%M%S')
    nome_arquivo_saida = f"medicao_consolidada_geral_{timestamp_exportacao}.csv"
    caminho_saida = pasta_destino / nome_arquivo_saida

    print(f" -> A guardar arquivo único: {nome_arquivo_saida} ({len(df_final)} linhas)...")

    if caminho_saida.exists():
        try:
            caminho_saida.unlink()
        except PermissionError:
            print(f"\n[ERRO] O ficheiro {nome_arquivo_saida} está aberto. Feche-o!")
            return

    try:
        df_final.to_csv(caminho_saida, index=False, sep=';', encoding='utf-8-sig')
    except Exception as e:
        print(f"[!] Erro crítico ao guardar {nome_arquivo_saida}: {e}")
        return

    # ---------------------------------------------------------
    # FASE 4: BACKUP (COMPACTAÇÃO ZIP) E LIMPEZA
    # ---------------------------------------------------------
    print("\nA compactar e mover ficheiros processados para Backup...")
    if arquivos_processados_com_sucesso:
        timestamp_zip = datetime.now().strftime('%Y%m%d%H%M%S')
        nome_zip = f"arquivo-processado_medicao_{timestamp_zip}.zip"
        caminho_zip = pasta_backup / nome_zip

        try:
            with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for caminho in arquivos_processados_com_sucesso:
                    zipf.write(caminho, arcname=caminho.name)
            for caminho in arquivos_processados_com_sucesso:
                try:
                    caminho.unlink()
                except Exception as e:
                    print(f" -> [!] Não foi possível excluir '{caminho.name}': {e}")
            print(f" -> ✅ Arquivos compactados com sucesso em: {nome_zip}")
        except Exception as e:
            print(f" -> [!] Erro ao criar o arquivo ZIP de backup: {e}")

    total_linhas_destino = len(df_final)

    print("\n" + "=" * 60)
    print("RESUMO DA CARGA INCREMENTAL".center(60))
    print("=" * 60)
    print("FICHEIROS DE ORIGEM (.xlsx):")
    print(f" -> Lidos na pasta origem:  {len(arquivos)}")
    print(f" -> Compactados p/ backup:  {len(arquivos_processados_com_sucesso)}")

    print("\nAUDITORIA DE MAPEAMENTO:")
    print(f" -> Colunas Padronizadas:   {sum(1 for col in ORDEM_PADRAO if col in colunas_presentes)}")
    print(f" -> Colunas Extras (Fim):   {len(colunas_extras)}")
    print(f" -> Total de Colunas:       {len(colunas_finais)}")

    print("\nAUDITORIA DE DADOS (LINHAS):")
    print(f" -> Linhas Extraídas:       {total_linhas_origem}")
    print(f" -> Linhas Processadas:     {total_linhas_destino}")

    if total_linhas_origem == total_linhas_destino:
        print(" -> Status Integridade:     ✅ SUCESSO! 100% de match nas linhas.")
    else:
        print(f" -> Status Integridade:     ⚠️ ATENÇÃO! Diferença de {total_linhas_origem - total_linhas_destino} linhas.")
    print("=" * 60)