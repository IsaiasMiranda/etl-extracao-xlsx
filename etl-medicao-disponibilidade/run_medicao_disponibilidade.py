import logging
from pathlib import Path
from load_medicao_disponibilidade import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    # Definição dos novos caminhos apontados para Medição de Disponibilidades
    origem = Path(r'D:\One-Drive\Amper Elinsa\MEDICAO_DISPONIBILIDADES - Documentos')
    destino = origem / 'arquivo-consolidado'
    processados = origem / 'arquivo-processado'

    logger = logging.getLogger(__name__)
    logger.info("Iniciando rotina de consolidação e unpivot de Disponibilidades...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=destino,
        pasta_processados=processados
    )

    logger.info("Rotina finalizada.")