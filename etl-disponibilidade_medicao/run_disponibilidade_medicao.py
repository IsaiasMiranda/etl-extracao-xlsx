import logging
from pathlib import Path
from load_disponibilidade_medicao import processar_boletins

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

if __name__ == '__main__':
    # Definição dos novos caminhos apontados para Medição de Disponibilidades
    origem = Path(r'D:\base-geral\base-disponibilidade-medicao\source\MEDICAO_DISPONIBILIDADES - Documentos')
    arquivo_consolidado = Path(r'D:\base-geral\base-disponibilidade-medicao\arquivo-consolidado')
    arquivo_processado = Path(r'D:\base-geral\base-disponibilidade-medicao\arquivo-processado')

    logger = logging.getLogger(__name__)
    logger.info("Iniciando rotina de consolidação e unpivot de Disponibilidades...")

    processar_boletins(
        pasta_origem=origem,
        pasta_destino=arquivo_consolidado,
        pasta_processados=arquivo_processado
    )

    logger.info("Rotina finalizada.")