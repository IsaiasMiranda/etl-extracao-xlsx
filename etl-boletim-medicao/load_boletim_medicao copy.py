import logging
import shutil
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

logger = logging.getLogger(__name__)


# ------------------------------------------------------------------
# MAPEAMENTO COMPLETO (todas as variações → padrão)
# ------------------------------------------------------------------
MAPEAMENTO = {
    # ------------------------------------------------------------------
    # Boletim
    # ------------------------------------------------------------------
    'boletim de medição': 'boletim',
    'boletim de medicao': 'boletim',
    'folha de registro': 'boletim',
    'boletim': 'boletim',

    # ------------------------------------------------------------------
    # Nota fiscal
    # ------------------------------------------------------------------
    'nº nota fiscal': 'nota_fiscal',
    'n° nota fiscal': 'nota_fiscal',
    'no nota fiscal': 'nota_fiscal',
    'numero nota fiscal': 'nota_fiscal',
    'número nota fiscal': 'nota_fiscal',
    'nota fiscal': 'nota_fiscal',
    'nota_fiscal': 'nota_fiscal',
    'nf': 'nota_fiscal',

    # ------------------------------------------------------------------
    # Valor
    # ------------------------------------------------------------------
    'valor total': 'valor',
    'valor bruto': 'valor',
    'valor': 'valor',

    # ------------------------------------------------------------------
    # Data de envio
    # ------------------------------------------------------------------
    'data de envio do boletim': 'data_envio',
    'data de envio': 'data_envio',
    'data envio': 'data_envio',
    'data_envio': 'data_envio',

    # ------------------------------------------------------------------
    # Origem de lançamento
    # ------------------------------------------------------------------
    'coletor custo': 'origem_lancamento',
    'coletorcusto': 'origem_lancamento',
    'origem de lançamento (pep, diagr, ord, cc)': 'origem_lancamento',
    'origem de lancamento (pep, diagr, ord, cc)': 'origem_lancamento',
    'origem de lançamento (pep,diagr,ord,cc)': 'origem_lancamento',
    'origem de lancamento (pep,diagr,ord,cc)': 'origem_lancamento',
    'origem de lançamento\n(pep, diagr, ord, cc)': 'origem_lancamento',
    'origem de lançamento': 'origem_lancamento',
    'origem de lancamento': 'origem_lancamento',
    'origem_lancamento': 'origem_lancamento',

    # ------------------------------------------------------------------
    # Município
    # ------------------------------------------------------------------
    'município': 'municipio',
    'municipio': 'municipio',

    # ------------------------------------------------------------------
    # Competência / Período
    # ------------------------------------------------------------------
    'período de medição': 'competencia',
    'periodo de medição': 'competencia',
    'período de medicao': 'competencia',
    'periodo de medicao': 'competencia',
    'período': 'competencia',
    'periodo': 'competencia',
    'competência': 'competencia',
    'competencia': 'competencia',

    # ------------------------------------------------------------------
    # Descrição da nota fiscal
    # ------------------------------------------------------------------
    'descrição da nota fiscal': 'descricaonotafiscal',
    'descricao da nota fiscal': 'descricaonotafiscal',
    'descrição nota fiscal': 'descricaonotafiscal',
    'descricao nota fiscal': 'descricaonotafiscal',
    'descricaonotafiscal': 'descricaonotafiscal',

    # ------------------------------------------------------------------
    # CP / CM
    # ------------------------------------------------------------------
    'cm': 'cp',
    'cp': 'cp',
    'centro': 'cp',

    # ------------------------------------------------------------------
    # Tipo de medição – todas as variações identificadas
    # Cobre: acentuação, erro tipográfico "medicão", parênteses opcionais
    # ------------------------------------------------------------------
    'tipo de medição (cust/invest/ativ)': 'tipomedicao',
    'tipo de medição (cust/invest/atividade)': 'tipomedicao',
    'tipo de medicão (cust/invest/ativ)': 'tipomedicao',
    'tipo de medicão (cust/invest/atividade)': 'tipomedicao',
    'tipo de medicao (cust/invest/ativ)': 'tipomedicao',
    'tipo de medicao (cust/invest/atividade)': 'tipomedicao',
    'tipo de medição\n(cust/invest/ativ)': 'tipomedicao',
    'tipo de medição': 'tipomedicao',
    'tipo de medicão': 'tipomedicao',
    'tipo de medicao': 'tipomedicao',
    'tipomedicao': 'tipomedicao',

    # ------------------------------------------------------------------
    # Estrutura
    # ------------------------------------------------------------------
    'estrutura (prod/disp)': 'estrutura',
    'estrutura(prod/disp)': 'estrutura',
    'estrutura (prod / disp)': 'estrutura',
    'estrutura\n(prod/disp)': 'estrutura',
    'estrutura': 'estrutura',

    # ------------------------------------------------------------------
    # Medição (ciclo) – todas as variações
    # Cobre: acentuação, espaço extra no "fin al", parênteses opcionais
    # ------------------------------------------------------------------
    'medição (ciclo/final/pend)': 'medicao',
    'medição (ciclo/fin al/pend)': 'medicao',
    'medicão (ciclo/final/pend)': 'medicao',
    'medicao (ciclo/final/pend)': 'medicao',
    'medicão (ciclo/fin al/pend)': 'medicao',
    'medicao (ciclo/fin al/pend)': 'medicao',
    'medição\n(ciclo/final/pend)': 'medicao',
    'medição': 'medicao',
    'medicão': 'medicao',
    'medicao': 'medicao',

    # ------------------------------------------------------------------
    # Campos já padronizados — entradas defensivas incluídas
    # ------------------------------------------------------------------
    'distribuidora': 'distribuidora',
    'regional': 'regional',
    'parceiro': 'parceiro',
    'equipe': 'equipe',
    'iva': 'iva',
    'domicilio fiscal': 'domiciliofiscal',
    'domicílio fiscal': 'domiciliofiscal',
    'domiciliofiscal': 'domiciliofiscal',
    'codigo tarifafiscal': 'codigotarifafiscal',
    'código tarifa fiscal': 'codigotarifafiscal',
    'codigotarifafiscal': 'codigotarifafiscal',
}

# ------------------------------------------------------------------
# ORDEM FINAL DAS COLUNAS (as colunas "centro" e "nota proj"
# não estão nesta lista e serão adicionadas automaticamente no final)
# ------------------------------------------------------------------
ORDEM_PADRAO = [
    'arquivo_origem',
    'distribuidora',
    'regional',
    'tipomedicao',
    'estrutura',
    'medicao',          # corrigido: era 'medição' (com acento), mas o MAPEAMENTO padroniza para 'medicao'
    'competencia',
    'parceiro',
    'municipio',
    'equipe',
    'descricaonotafiscal',
    'boletim',
    'valor',
    'data_envio',
    'origem_lancamento',
    'cp',
    'iva',
    'nota_fiscal',
    'domiciliofiscal',
    'codigotarifafiscal'
]


# ------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------------------------------------------

def _localizar_linha_cabecalho(df_bruto: pd.DataFrame) -> int:
    for i, row in df_bruto.iterrows():
        linha_texto = ' '.join(row.fillna('').astype(str)).lower()
        if 'distribuidora' in linha_texto and 'regional' in linha_texto:
            return i
    return -1


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
    df_tabela.columns = df_tabela.iloc[0].astype(str).str.strip().str.lower()
    df_tabela = df_tabela.iloc[1:].copy()

    df_tabela.dropna(axis=1, how='all', inplace=True)
    df_tabela.dropna(axis=0, how='all', inplace=True)

    if df_tabela.empty:
        return None

    df_tabela = df_tabela.loc[:, ~df_tabela.columns.duplicated()]

    df_tabela.rename(columns=MAPEAMENTO, inplace=True)

    df_tabela['arquivo_origem'] = nome_arquivo

    colunas_validas = [
        col for col in df_tabela.columns
        if not str(col).lower().startswith('unnamed')
        and str(col).lower() not in ['nan', 'none', '', '<na>']
    ]
    df_tabela = df_tabela[colunas_validas]
    df_tabela = df_tabela.loc[:, ~df_tabela.columns.duplicated()]

    return df_tabela


def _processar_arquivo(caminho: Path) -> Tuple[List[pd.DataFrame], bool]:
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
            df_aba = _processar_aba(sheet, caminho.name)
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
        logger.info(f"[{idx}/{total_arquivos}] Processando: {caminho.name} ...")
        dfs, sucesso = _processar_arquivo(caminho)
        if sucesso and dfs:
            total_linhas += sum(len(df) for df in dfs)
            todos_dfs.extend(dfs)
            arquivos_lidos.append(caminho)
            logger.info(f"  -> OK: {len(dfs)} aba(s), {sum(len(df) for df in dfs)} linhas.")
        else:
            arquivos_erro.append(caminho)
            logger.warning(f"  -> Nenhum dado válido. Arquivo mantido na origem.")

    if not todos_dfs:
        logger.error("Nenhum dado extraído. Abortando.")
        return

    logger.info("\nConcatenando dados...")
    df_final = pd.concat(todos_dfs, ignore_index=True)

    # --- Ordenação final: colunas padrão primeiro, extras depois ---
    colunas_presentes = [col for col in ORDEM_PADRAO if col in df_final.columns]
    colunas_extras = [col for col in df_final.columns if col not in ORDEM_PADRAO]
    df_final = df_final[colunas_presentes + colunas_extras]

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
        for caminho_origem in arquivos_lidos:
            destino_arq = pasta_processados / caminho_origem.name
            try:
                if destino_arq.exists():
                    destino_arq.unlink()
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