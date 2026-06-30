import logging
from pathlib import Path

from load_boletim_medicao import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    origem = Path(r'D:\One-Drive\Amper Elinsa\Fabiano Braz da Silva - Boletins de Medição\PARÁ')
    destino = origem / 'base-consolidado'
    processados = origem / 'arquivos-processados'

    logger = logging.getLogger(__name__)
    logger.info("Iniciando o processamento dos boletins de medição...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=destino,
        pasta_processados=processados
    )

    logger.info("Script finalizado.")