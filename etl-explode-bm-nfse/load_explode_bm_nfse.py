import logging
from pathlib import Path
import pandas as pd

logger = logging.getLogger(__name__)

def quebrar_valores_em_linhas(df: pd.DataFrame, coluna: str) -> pd.DataFrame:
    """
    Recebe um DataFrame e expande a coluna informada quebrando por delimitadores.
    """
    if coluna not in df.columns:
        logger.warning(f"A coluna '{coluna}' não foi encontrada no DataFrame.")
        return df

    df_transformado = df.copy()
    df_transformado[coluna] = df_transformado[coluna].astype(str)

    # Divide as strings por vírgula, ponto e vírgula, barras ou quebras de linha
    df_transformado[coluna] = df_transformado[coluna].str.split(r'[;,/|\n]+')

    # Desfaz o agrupamento (explode) transformando listas em novas linhas
    df_transformado = df_transformado.explode(coluna)

    # Limpeza de strings e remoção de valores nulos/vazios resultantes
    df_transformado[coluna] = df_transformado[coluna].str.strip()
    df_transformado = df_transformado[
        (df_transformado[coluna] != '') & 
        (df_transformado[coluna].str.lower() != 'nan') &
        (df_transformado[coluna].str.lower() != 'none')
    ]

    return df_transformado

def processar_arquivos_explode(pasta_origem: Path, pasta_destino: Path, coluna_alvo: str = 'boletim') -> None:
    """
    Mapeia arquivos Excel na pasta informada, aplica o explode e exporta como CSV.
    """
    pasta_origem = Path(pasta_origem)
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # Busca especificamente por planilhas Excel (.xlsx e .xls) para não conflitar com os CSVs de saída
    arquivos = list(pasta_origem.glob("*.xlsx")) + list(pasta_origem.glob("*.xls"))
    
    # Filtra arquivos temporários que o Excel cria quando a planilha está aberta
    arquivos = [f for f in arquivos if not f.name.startswith('~$')]

    if not arquivos:
        logger.warning(f"Nenhum arquivo Excel (.xlsx / .xls) encontrado em: {pasta_origem}")
        return

    logger.info(f"Foram encontrados {len(arquivos)} arquivo(s) para processamento.")

    for arquivo in arquivos:
        logger.info(f"Processando planilha: {arquivo.name}")
        try:
            # Carrega a planilha Excel de forma dinâmica
            df = pd.read_excel(arquivo)

            # Normaliza os cabeçalhos (sem espaços e em minúsculas)
            df.columns = df.columns.astype(str).str.strip().str.lower()
            linhas_antes = len(df)

            # Executa a quebra dos boletins
            df_final = quebrar_valores_em_linhas(df, coluna_alvo.lower())
            linhas_depois = len(df_final)

            # Define o nome do arquivo final estruturado
            nome_saida = f"{arquivo.stem}_explodido.csv"
            caminho_saida = pasta_destino / nome_saida

            # Exporta para CSV estruturado em UTF-8
            df_final.to_csv(caminho_saida, index=False, encoding='utf-8')
            logger.info(f"  -> Sucesso: {linhas_antes} linhas estruturadas em {linhas_depois} registros pós-explode.")

        except Exception as e:
            logger.error(f"  -> Erro crítico ao processar o arquivo {arquivo.name}: {e}", exc_info=True)

    logger.info("Pipeline de processamento local concluído.")