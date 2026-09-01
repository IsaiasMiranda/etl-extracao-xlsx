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

    logger = logging.getLogger(__name__)
    logger.info("Iniciando o processamento dos boletins de medição...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=destino,
        pasta_processados=processados
    )

    logger.info("Script finalizado.")