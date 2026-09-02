import logging
from pathlib import Path

from load_boletim_medicao import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    origem = Path(r'D:\base-geral\base-boletim-medicao\source')
    destino = Path(r'D:\base-geral\base-boletim-medicao\base-consolidado')
    processados = Path(r'D:\base-geral\base-boletim-medicao\arquivos-processados')
    auditoria = Path(r'D:\base-geral\base-boletim-medicao\auditoria\boletim-medicao-normalizacao.csv')

    logger = logging.getLogger(__name__)
    logger.info("Iniciando o processamento dos boletins de medição...")

    # Regras de limpeza promovidas de homologacao/ em 2026-09-02, depois
    # de validadas ponta a ponta contra o histórico bruto real de
    # produção (ver CLAUDE.md/AGENTS.md, seção 8) -- mesmos parâmetros
    # que homologacao/run_boletim_medicao_homologacao.py já usa, agora
    # apontando pras pastas reais de produção.
    processar_boletins(
        pasta_origem=origem,
        pasta_destino=destino,
        pasta_processados=processados,
        normalizar_deslocamentos=True,
        caminho_auditoria=auditoria,
        remover_sem_boletim=True,
        excluir_boletim_nfse_duplicado=True,
        remover_boletim_sem_valor_duplicado=True,
    )

    logger.info("Script finalizado.")