import logging
from pathlib import Path
import os

# Raiz dos dados (B4, 2026-09-19): variavel de ambiente BASE_GERAL ou padrao local.
BASE_GERAL = Path(os.environ.get('BASE_GERAL', 'D:/base-geral'))

from load_unifica_turno_gpm import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    origem = BASE_GERAL / 'base-turnos-gpm/source/TurnosGPM'
    destino = BASE_GERAL / 'base-turnos-gpm/base-consolidado'
    processados = BASE_GERAL / 'base-turnos-gpm/arquivos-processados'

    logger = logging.getLogger(__name__)
    logger.info("Iniciando o processamento dos turnos GPM...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=destino,
        pasta_processados=processados,
    )

    logger.info("Script finalizado.")
