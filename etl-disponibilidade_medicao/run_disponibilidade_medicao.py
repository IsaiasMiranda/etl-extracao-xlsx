import logging
from pathlib import Path
import os

# Raiz dos dados (B4, 2026-09-19): variavel de ambiente BASE_GERAL ou padrao local.
BASE_GERAL = Path(os.environ.get('BASE_GERAL', 'D:/base-geral'))
from load_disponibilidade_medicao import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    # Definição dos novos caminhos apontados para Medição de Disponibilidades
    origem = BASE_GERAL / 'base-disponibilidade-medicao/source/MEDICAO_DISPONIBILIDADES - Documentos'
    arquivo_consolidado = BASE_GERAL / 'base-disponibilidade-medicao/arquivo-consolidado'
    arquivo_processado = BASE_GERAL / 'base-disponibilidade-medicao/arquivo-processado'

    logger = logging.getLogger(__name__)
    logger.info("Iniciando rotina de consolidação e unpivot de Disponibilidades...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=arquivo_consolidado,
        pasta_processados=arquivo_processado
    )

    logger.info("Rotina finalizada.")