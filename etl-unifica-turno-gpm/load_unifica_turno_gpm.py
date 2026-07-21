"""
load_unifica_turno_gpm.py
--------------------------
ETL para consolidação de arquivos de turnos GPM (.xlsx e .csv, inclusive em
subpastas) de uma pasta de origem. Os nomes de coluna são normalizados
(minúsculo, sem acento, "_" no lugar de espaço/pontuação) e depois mapeados
para um nome canônico através de um dicionário fixo de-para (MAPEAMENTO_COLUNAS),
igual ao padrão usado nos demais pipelines do projeto. Colunas mapeadas seguem
a ordem de ORDEM_PADRAO; colunas não mapeadas (fora do de-para) são anexadas
ao final, em ordem alfabética. Considera apenas a primeira aba visível de cada
arquivo .xlsx.

Autor: Engenharia de Dados
"""

import logging
import random
import re
import shutil
import unicodedata
import warnings
from pathlib import Path
from typing import List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException

warnings.filterwarnings("ignore", module="openpyxl")

logger = logging.getLogger(__name__)

MAX_LINHAS_VARREDURA_CABECALHO = 30  # segurança: não varre o arquivo inteiro
MIN_CELULAS_CABECALHO = 3            # nº mínimo de células preenchidas para considerar a linha um cabeçalho

# =====================================================================
# CONFIGURAÇÃO DE MAPEAMENTO DE COLUNAS (DE-PARA)
# =====================================================================
# Chave = nome da coluna já normalizado (minúsculo, sem acento, "_" no lugar
# de espaço/pontuação — ver normalizar_coluna()). Valor = nome padrão no
# dataframe final. Preencher com base no levantamento de campos das
# planilhas de origem.
MAPEAMENTO_COLUNAS = {
}

# Colunas mapeadas que devem ficar na frente do dataframe final, nessa ordem.
# Colunas não listadas aqui (mas presentes em MAPEAMENTO_COLUNAS ou não) são
# anexadas ao final, em ordem alfabética.
ORDEM_PADRAO = [
]


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
# EXTRAÇÃO DA TABELA (CABEÇALHO DINÂMICO + LIMPEZA)
# =====================================================================

def _localizar_linha_cabecalho(df_bruto: pd.DataFrame) -> int:
    """Retorna o índice da primeira linha com um número razoável de células
    preenchidas — logos/textos soltos no topo do arquivo costumam ocupar
    só 1 ou 2 células e ficam antes do cabeçalho real da tabela."""
    limite = min(len(df_bruto), MAX_LINHAS_VARREDURA_CABECALHO)
    for idx in range(limite):
        if df_bruto.iloc[idx].notna().sum() >= MIN_CELULAS_CABECALHO:
            return idx
    return -1


def _consolidar_colunas_duplicadas(colunas_normalizadas: list, contexto: str, log_avisos: list) -> list:
    """Se duas colunas nomeadas normalizarem para o mesmo texto, renomeia a
    partir da segunda ocorrência em vez de deixar o pandas colidir/descartar
    dado silenciosamente. Colunas sem nome (célula de cabeçalho vazia) não
    geram aviso aqui — são descartadas depois, em _extrair_tabela."""
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
    localizando o cabeçalho real e limpando linhas/colunas vazias."""
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
    # POSICIONAL (iloc), nunca por nome: evita expansão combinatória caso
    # algum nome ainda esteja duplicado.
    posicoes_validas = [i for i, c in enumerate(df_tabela.columns) if c != ""]
    df_tabela = df_tabela.iloc[:, posicoes_validas]

    df_tabela = df_tabela.dropna(axis=1, how='all')
    df_tabela = df_tabela.dropna(axis=0, how='all')

    return df_tabela


# =====================================================================
# LEITURA DE ARQUIVO (XLSX: 1ª ABA VISÍVEL / CSV)
# =====================================================================

def _ler_primeira_aba_visivel(caminho_arquivo: Path) -> Tuple[Optional[str], Optional[pd.DataFrame]]:
    """Abre o workbook e retorna apenas a PRIMEIRA aba visível encontrada
    (abas ocultas nunca concorrem)."""
    wb = openpyxl.load_workbook(caminho_arquivo, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible':
                continue
            valores = list(ws.values)
            if not valores:
                continue
            return ws.title, pd.DataFrame(valores)
    finally:
        wb.close()
    return None, None


def _ler_csv_bruto(caminho_arquivo: Path) -> Optional[pd.DataFrame]:
    """Lê um CSV sem cabeçalho pré-definido, detectando separador e encoding,
    para que passe pelo mesmo pipeline de localização de cabeçalho do XLSX."""
    for encoding in ('utf-8-sig', 'latin1'):
        try:
            return pd.read_csv(caminho_arquivo, header=None, sep=None, engine='python', encoding=encoding)
        except UnicodeDecodeError:
            continue
        except Exception as e:
            logger.error(f"Erro ao ler CSV '{caminho_arquivo.name}': {e}")
            return None
    return None


def _processar_arquivo(caminho: Path, log_avisos: list) -> List[pd.DataFrame]:
    """Processa um único arquivo (.xlsx ou .csv) e retorna a tabela extraída
    (já com a coluna 'arquivo_origem'), ou lista vazia se não houver dado válido."""
    nome_arquivo = caminho.name
    sufixo = caminho.suffix.lower()

    if sufixo == '.csv':
        df_bruto = _ler_csv_bruto(caminho)
        if df_bruto is None or df_bruto.empty:
            log_avisos.append(f"[SKIP] '{nome_arquivo}': CSV vazio ou ilegível.")
            return []
        contexto = f"'{nome_arquivo}'"
    else:
        try:
            nome_aba, df_bruto = _ler_primeira_aba_visivel(caminho)
        except (InvalidFileException, PermissionError, OSError) as e:
            log_avisos.append(f"[ERRO] '{nome_arquivo}': falha ao abrir arquivo ({e}).")
            return []

        if df_bruto is None:
            log_avisos.append(f"[SKIP] '{nome_arquivo}': nenhuma aba visível com dados.")
            return []
        contexto = f"'{nome_arquivo}' / aba '{nome_aba}'"

    df_tabela = _extrair_tabela(df_bruto, contexto, log_avisos)
    if df_tabela.empty:
        return []

    df_tabela = df_tabela.copy()
    df_tabela['arquivo_origem'] = nome_arquivo
    return [df_tabela]


# =====================================================================
# PIPELINE PRINCIPAL
# =====================================================================

def processar_boletins(
    pasta_origem: Path,
    pasta_destino: Path,
    pasta_processados: Path,
) -> None:
    pasta_origem = Path(pasta_origem)
    pasta_destino = Path(pasta_destino)
    pasta_processados = Path(pasta_processados)

    pasta_destino.mkdir(parents=True, exist_ok=True)
    pasta_processados.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 60)
    logger.info("CONSOLIDAÇÃO DE TURNOS GPM (XLSX + CSV)")
    logger.info(f"Origem : {pasta_origem}")
    logger.info(f"Destino: {pasta_destino}")
    logger.info("=" * 60)

    if not pasta_origem.exists():
        logger.error(f"Pasta de origem não encontrada: {pasta_origem}")
        return

    arquivos = [
        f for f in pasta_origem.rglob('*')
        if f.is_file()
        and f.suffix.lower() in ('.xlsx', '.csv')
        and not f.name.startswith('~$')
        and 'turnos-gpm-consolidado_' not in f.name
    ]

    total_arquivos = len(arquivos)
    if total_arquivos == 0:
        logger.warning("Nenhum arquivo .xlsx/.csv encontrado na origem.")
        return

    logger.info(f"Encontrados {total_arquivos} arquivo(s). Iniciando extração...\n")

    todos_dfs = []
    arquivos_lidos = []
    arquivos_erro = []
    log_avisos = []
    total_linhas = 0

    # ---------------------------------------------------------
    # FASE 1: LEITURA E EXTRAÇÃO
    # ---------------------------------------------------------
    for idx, caminho in enumerate(arquivos, start=1):
        logger.info(f"[{idx}/{total_arquivos}] Processando: {caminho.name} ...")
        dfs = _processar_arquivo(caminho, log_avisos)
        if dfs:
            linhas = sum(len(df) for df in dfs)
            total_linhas += linhas
            todos_dfs.extend(dfs)
            arquivos_lidos.append(caminho)
            logger.info(f"  -> OK: {linhas} linha(s).")
        else:
            arquivos_erro.append(caminho)
            logger.warning("  -> Nenhum dado válido encontrado. Mantido na origem.")

    if log_avisos:
        logger.info("\nAvisos durante a extração:")
        for aviso in log_avisos:
            logger.info(f"  {aviso}")

    if not todos_dfs:
        logger.error("Nenhum dado extraído de nenhum arquivo. Operação interrompida.")
        return

    # ---------------------------------------------------------
    # FASE 2: MAPEAMENTO (DE-PARA) E ORDENAÇÃO
    # ---------------------------------------------------------
    logger.info("\nA unir e mapear colunas da base de dados...")

    for i in range(len(todos_dfs)):
        todos_dfs[i] = todos_dfs[i].rename(columns=MAPEAMENTO_COLUNAS)

    df_final = pd.concat(todos_dfs, ignore_index=True, sort=False)

    colunas_presentes = set(df_final.columns)
    colunas_finais = [col for col in ORDEM_PADRAO if col in colunas_presentes]
    colunas_extras = sorted(colunas_presentes - set(colunas_finais) - {'arquivo_origem'})
    colunas_finais.extend(colunas_extras)

    if 'arquivo_origem' in colunas_presentes:
        colunas_finais.insert(0, 'arquivo_origem')

    df_final = df_final[colunas_finais]

    # ---------------------------------------------------------
    # FASE 3: EXPORTAÇÃO (CSV)
    # ---------------------------------------------------------
    timestamp = pd.Timestamp.now().strftime('%Y%m%d%H%M%S')
    caminho_consolidado = pasta_destino / f'turnos-gpm-consolidado_{timestamp}.csv'

    logger.info(f"Gravando arquivo consolidado: {caminho_consolidado.name} ({len(df_final)} linhas)...")
    try:
        df_final.to_csv(caminho_consolidado, index=False, sep=';', encoding='utf-8-sig')
    except PermissionError:
        logger.error("ERRO: O arquivo destino está aberto em outro programa. Feche-o e tente novamente.")
        return

    # ---------------------------------------------------------
    # FASE 4: MOVER ARQUIVOS PROCESSADOS
    # ---------------------------------------------------------
    logger.info(f"\nMovendo {len(arquivos_lidos)} arquivo(s) para '{pasta_processados}' ...")
    movidos = 0
    for caminho_origem in arquivos_lidos:
        destino_arq = pasta_processados / caminho_origem.name
        if destino_arq.exists():
            prefixo = random.randint(1000, 9999)
            destino_arq = pasta_processados / f"{prefixo} - {caminho_origem.name}"
        try:
            shutil.move(str(caminho_origem), str(destino_arq))
            movidos += 1
        except Exception as e:
            logger.error(f"Erro ao mover {caminho_origem.name}: {e}")

    # ---------------------------------------------------------
    # RESUMO
    # ---------------------------------------------------------
    logger.info("\n" + "=" * 60)
    logger.info("RESUMO DA CONSOLIDAÇÃO")
    logger.info("=" * 60)
    logger.info(f"Arquivos na origem       : {total_arquivos}")
    logger.info(f"Arquivos processados     : {len(arquivos_lidos)}")
    logger.info(f"Arquivos com erro/vazios : {len(arquivos_erro)}")
    logger.info(f"Arquivos movidos         : {movidos}")
    logger.info(f"Colunas padronizadas     : {sum(1 for col in ORDEM_PADRAO if col in colunas_presentes)}")
    logger.info(f"Colunas extras (fim)     : {len(colunas_extras)}")
    logger.info(f"Linhas extraídas         : {total_linhas}")
    logger.info(f"Linhas salvas            : {len(df_final)}")
    logger.info("-" * 60)

    if arquivos_erro:
        logger.warning("Arquivos NÃO processados (permanecem na origem):")
        for arq in arquivos_erro:
            logger.warning(f"  - {arq.name}")

    logger.info("=" * 60)
    logger.info("PROCESSAMENTO CONCLUÍDO!")
