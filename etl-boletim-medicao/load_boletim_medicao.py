import calendar
import logging
import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

import random
import pandas as pd
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

logger = logging.getLogger(__name__)


# ------------------------------------------------------------------
# MAPEAMENTO COMPLETO (todas as variações → padrão)
# ------------------------------------------------------------------
MAPEAMENTO = {
    'distribuidora': 'distribuidora',
    'regional': 'regional',
    'tipodemedicao_cust_invest_ativ': 'tipo_medicao',
    'tipo_de_medicao_cust_invest_ativ': 'tipo_medicao',
    'tipomedicao': 'tipo_medicao',
    'estrutura_prod_disp': 'estrutura',
    'medicao_ciclo_final_pend': 'medicao',
    'periododemedicao': 'periodo_medicao',
    'periodo_de_medicao': 'periodo_medicao',
    'competencia': 'periodo_medicao',
    'competencia_servico': 'periodo_medicao',
    'parceiro': 'parceiro',
    'municipio': 'municipio',
    'municipiodolocaldaprestacaodeservico': 'municipio',
    'equipe': 'equipe',
    'descricaodanotafiscal': 'desc_nota_fiscal',
    'descricao_da_nota_fiscal': 'desc_nota_fiscal',
    'descricaonotafiscal': 'desc_nota_fiscal',
    'folhaderegistro': 'boletim',
    'boletim': 'boletim',
    'boletimdemedicao': 'boletim',
    'boletim_de_medicao': 'boletim',
    'valor': 'valor_bm',
    'datadeenviodoboletim': 'data_envio_bm',
    'data_de_envio_do_boletim': 'data_envio_bm',
    'cm': 'cp',
    'centro': 'cp',
    'cp': 'cp',
    'iva': 'iva',
    'nonotafiscal': 'nota_fiscal',
    'nf': 'nota_fiscal',
    'domiciliofiscal': 'domicilio_fiscal',
    'codigotarifafiscal': 'codigo_tarifa_fiscal',
    'cod_tarifafiscal': 'codigo_tarifa_fiscal',
    'identificadormedicao': 'identificador_medicao',
    'identificadoragrupamento': 'identificador_agrupamento',
    'contrato': 'contrato',
    'processo': 'processo',
    'textoboletim': 'texto_boletim',
    'coletorcusto': 'origem_lancamento',
    'origemdelancamento_pep_diagr_ord_cc': 'origem_lancamento',
    'origem_de_lancamento_pep_diagr_ord_cc': 'origem_lancamento',
    'notaproj': 'nota_prol',
}

# Colunas removidas do modelo final (não fazem mais parte do consolidado)
COLUNAS_REMOVIDAS = ['nota_fiscal', 'nota_prol']

# ------------------------------------------------------------------
# ORDEM FINAL DAS COLUNAS (a coluna "centro" não está nesta lista
# e será adicionada automaticamente no final)
# ------------------------------------------------------------------
ORDEM_PADRAO = [
    'distribuidora',
    'regional',
    'tipo_medicao',
    'estrutura',
    'medicao',
    'periodo_medicao',
    'parceiro',
    'municipio',
    'equipe',
    'desc_nota_fiscal',
    'boletim',
    'valor_bm',
    'data_envio_bm',
    'origem_lancamento',
    'cp',
    'iva',
    'domicilio_fiscal',
    'codigo_tarifa_fiscal',
    'identificador_medicao',
    'identificador_agrupamento',
    'contrato',
    'processo',
    'texto_boletim',
]


# ------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------------------------------------------

def _normalizar_cabecalho(texto) -> str:
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'[^a-z0-9]+', '_', texto)
    return texto.strip('_')


MESES_PT = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
}

_RE_INTERVALO_PERIODO = re.compile(r'^(\d{2}\.\d{2}\.\d{4})\s*\S*\s*(\d{2}\.\d{2}\.\d{4})$')


def _normalizar_periodo_medicao(valor, hoje: Optional[datetime] = None):
    if not isinstance(valor, str) or not valor.strip():
        return valor
    texto = valor.strip()

    m = _RE_INTERVALO_PERIODO.match(texto)
    if m:
        return f"{m.group(1)} a {m.group(2)}"

    mes_num = MESES_PT.get(texto.lower())
    if mes_num:
        hoje = hoje or datetime.now()
        ano = hoje.year if mes_num <= hoje.month else hoje.year - 1
        ultimo_dia = calendar.monthrange(ano, mes_num)[1]
        return f"01.{mes_num:02d}.{ano} a {ultimo_dia:02d}.{mes_num:02d}.{ano}"

    logger.warning(f"periodo_medicao em formato não reconhecido: {texto!r} — mantido sem alteração.")
    return texto


def _localizar_linha_cabecalho(df_bruto: pd.DataFrame) -> int:
    for i, row in df_bruto.iterrows():
        linha_texto = ' '.join(row.fillna('').astype(str)).lower()
        if (
            'distribuidora' in linha_texto
            and 'regional' in linha_texto
            and 'valor' in linha_texto
        ):
            return i
    return -1


def _remover_linha_somatorio(df: pd.DataFrame) -> pd.DataFrame:
    if 'valor_bm' not in df.columns:
        return df
    outras_colunas = [c for c in df.columns if c != 'valor_bm']
    e_somatorio = df[outras_colunas].isna().all(axis=1) & df['valor_bm'].notna()
    return df[~e_somatorio]


def _processar_aba(sheet, nome_arquivo: str) -> Optional[pd.DataFrame]:
    dados_aba = list(sheet.iter_rows(values_only=True))
    if not dados_aba:
        return None

    df_bruto = pd.DataFrame(dados_aba)
    linha_cabecalho = _localizar_linha_cabecalho(df_bruto)
    if linha_cabecalho == -1:
        logger.debug(f"Aba '{sheet.title}' ignorada: cabeçalho não encontrado.")
        return None

    df_tabela = df_bruto.iloc[linha_cabecalho:].copy()
    df_tabela.columns = df_tabela.iloc[0].map(_normalizar_cabecalho)
    df_tabela = df_tabela.iloc[1:].copy()

    df_tabela.dropna(axis=1, how='all', inplace=True)
    df_tabela.dropna(axis=0, how='all', inplace=True)

    if df_tabela.empty:
        return None

    df_tabela = df_tabela.loc[:, ~df_tabela.columns.duplicated()]

    df_tabela.rename(columns=MAPEAMENTO, inplace=True)
    df_tabela.drop(columns=COLUNAS_REMOVIDAS, errors='ignore', inplace=True)

    if 'periodo_medicao' in df_tabela.columns:
        df_tabela['periodo_medicao'] = df_tabela['periodo_medicao'].apply(_normalizar_periodo_medicao)

    df_tabela = _remover_linha_somatorio(df_tabela)

    df_tabela['arquivo_origem'] = nome_arquivo

    colunas_validas = [
        col for col in df_tabela.columns
        if not str(col).lower().startswith('unnamed')
        and str(col).lower() not in ['nan', 'none', '', '<na>']
    ]
    df_tabela = df_tabela[colunas_validas]
    df_tabela = df_tabela.loc[:, ~df_tabela.columns.duplicated()]

    return df_tabela


def _processar_arquivo(caminho: Path, nome_saida: Optional[str] = None) -> Tuple[List[pd.DataFrame], bool]:
    dfs_extraidos = []
    arquivo_processado = False

    try:
        workbook = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    except (InvalidFileException, PermissionError, OSError) as e:
        logger.error(f"Erro ao abrir {caminho.name}: {e}")
        return [], False

    try:
        for sheet in workbook.worksheets:
            if sheet.sheet_state != 'visible':
                continue
            df_aba = _processar_aba(sheet, nome_saida or caminho.name)
            if df_aba is not None and not df_aba.empty:
                dfs_extraidos.append(df_aba)
                arquivo_processado = True
    except Exception as e:
        logger.error(f"Erro inesperado em {caminho.name}: {e}", exc_info=True)
        return [], False
    finally:
        workbook.close()

    return dfs_extraidos, arquivo_processado


# ------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ------------------------------------------------------------------

def processar_boletins(
    pasta_origem: Path,
    pasta_destino: Path,
    pasta_processados: Optional[Path] = None
) -> None:
    pasta_origem = Path(pasta_origem)
    pasta_destino = Path(pasta_destino)
    if pasta_processados is None:
        pasta_processados = pasta_origem / 'arquivos-processados'
    else:
        pasta_processados = Path(pasta_processados)

    pasta_destino.mkdir(parents=True, exist_ok=True)
    pasta_processados.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 60)
    logger.info("CONSOLIDAÇÃO DE BOLETINS DE MEDIÇÃO")
    logger.info(f"Origem : {pasta_origem}")
    logger.info(f"Destino: {pasta_destino}")
    logger.info("=" * 60)

    arquivos = list(pasta_origem.glob("*.xlsx"))
    arquivos = [f for f in arquivos if not f.name.startswith('~$')]

    # Evita reprocessar consolidados existentes (CSV)
    consolidados = list(pasta_destino.glob("boletim-medicao-consolidado_*.xlsx"))
    nomes_consolidados = {f.name for f in consolidados}
    arquivos = [f for f in arquivos if f.name not in nomes_consolidados]

    total_arquivos = len(arquivos)
    if total_arquivos == 0:
        logger.warning("Nenhum arquivo .xlsx válido encontrado.")
        return

    logger.info(f"Encontrados {total_arquivos} arquivo(s).\n")

    todos_dfs = []
    arquivos_lidos = []
    arquivos_erro = []
    total_linhas = 0

    for idx, caminho in enumerate(arquivos, start=1):
        # Determina o nome de saída antes de processar:
        # se já existe arquivo igual em processados, o novo receberá prefixo aleatório
        destino_previsto = pasta_processados / caminho.name
        if destino_previsto.exists():
            prefixo = random.randint(1000, 9999)
            nome_saida = f"{prefixo} - {caminho.name}"
        else:
            nome_saida = caminho.name

        logger.info(f"[{idx}/{total_arquivos}] Processando: {caminho.name} ...")
        dfs, sucesso = _processar_arquivo(caminho, nome_saida=nome_saida)
        if sucesso and dfs:
            total_linhas += sum(len(df) for df in dfs)
            todos_dfs.extend(dfs)
            arquivos_lidos.append((caminho, nome_saida))
            logger.info(f"  -> OK: {len(dfs)} aba(s), {sum(len(df) for df in dfs)} linhas.")
        else:
            arquivos_erro.append(caminho)
            logger.warning(f"  -> Nenhum dado válido. Arquivo mantido na origem.")

    if not todos_dfs:
        logger.error("Nenhum dado extraído. Abortando.")
        return

    logger.info("\nConcatenando dados...")
    df_final = pd.concat(todos_dfs, ignore_index=True)

    # --- Ordenação final: todas as colunas padrão (mesmo ausentes no lote), depois extras, arquivo_origem sempre por último ---
    colunas_extras = [
        col for col in df_final.columns
        if col not in ORDEM_PADRAO and col != 'arquivo_origem'
    ]
    colunas_finais = ORDEM_PADRAO + colunas_extras
    if 'arquivo_origem' in df_final.columns:
        colunas_finais.append('arquivo_origem')
    df_final = df_final.reindex(columns=colunas_finais)

    # --- Nome do arquivo com data e hora (yyyymmddhhmmss) ---
    timestamp = pd.Timestamp.now().strftime('%Y%m%d%H%M%S')
    caminho_consolidado = pasta_destino / f'boletim-medicao-consolidado_{timestamp}.xlsx'

    logger.info(f"Salvando consolidado em: {caminho_consolidado} ...")
    try:
        with pd.ExcelWriter(caminho_consolidado, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Boletins')
    except PermissionError:
        logger.error("ERRO: O arquivo destino está aberto em outro programa. Feche-o e tente novamente.")
        return

    total_salvas = len(df_final)

    # --- Relatório ---
    logger.info("\n" + "=" * 60)
    logger.info("RESUMO DA CONSOLIDAÇÃO")
    logger.info("=" * 60)
    logger.info(f"Arquivos na origem          : {total_arquivos}")
    logger.info(f"Arquivos processados        : {len(arquivos_lidos)}")
    logger.info(f"Arquivos com erro/sem dados : {len(arquivos_erro)}")
    logger.info(f"Linhas extraídas            : {total_linhas}")
    logger.info(f"Linhas salvas               : {total_salvas}")
    logger.info("-" * 60)

    if arquivos_erro:
        logger.warning("Arquivos NÃO processados (permanecem na origem):")
        for arq in arquivos_erro:
            logger.warning(f"  - {arq.name}")

    if arquivos_lidos:
        logger.info(f"\nMovendo {len(arquivos_lidos)} arquivo(s) para '{pasta_processados}' ...")
        movidos = 0
        for caminho_origem, nome_saida in arquivos_lidos:
            destino_arq = pasta_processados / nome_saida
            try:
                if nome_saida != caminho_origem.name:
                    logger.warning(
                        f"  ⚠️  Arquivo duplicado encontrado: '{caminho_origem.name}' "
                        f"já existia em processados. Movido como '{nome_saida}'."
                    )
                shutil.move(str(caminho_origem), str(destino_arq))
                movidos += 1
            except Exception as e:
                logger.error(f"Erro ao mover {caminho_origem.name}: {e}")
        logger.info(f"✅ {movidos} arquivo(s) movidos com sucesso.")
    else:
        logger.info("Nenhum arquivo para mover.")

    if total_linhas != total_salvas:
        logger.warning(
            f"⚠️ Diferença de linhas: extraídas={total_linhas}, salvas={total_salvas}. "
            "Possível eliminação de duplicatas ou linhas vazias."
        )
    else:
        logger.info("✅ Contagem de linhas consistente.")

    logger.info("=" * 60)
    logger.info("PROCESSAMENTO CONCLUÍDO!")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    ORIGEM = Path(r'D:\One-Drive\Amper Elinsa\Fabiano Braz da Silva - Boletins de Medição\PARÁ')
    DESTINO = ORIGEM / 'base-consolidado'
    processar_boletins(ORIGEM, DESTINO)