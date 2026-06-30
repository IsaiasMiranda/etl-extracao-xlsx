import logging
from pathlib import Path

# Importa a função principal do módulo de carga/transformação
from load_explode_bm_nfse import processar_arquivos_explode

# Configuração global de logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)
logger = logging.getLogger(__name__)

if __name__ == '__main__':
    # Mapeia dinamicamente a pasta exata onde este script está executando
    DIRETORIO_ATUAL = Path(__file__).parent.resolve()
    
    # Origem e destino apontam para a mesma pasta do script
    PASTA_ORIGEM = DIRETORIO_ATUAL
    PASTA_DESTINO = DIRETORIO_ATUAL
    
    # Coluna que será fragmentada
    COLUNA_PARA_EXPLODIR = 'boletim'

    logger.info("=" * 60)
    logger.info("Iniciando rotina de separação de boletins...")
    logger.info(f"Diretório de busca (Origem): {PASTA_ORIGEM}")
    logger.info(f"Diretório de gravação (Destino): {PASTA_DESTINO}")
    logger.info("=" * 60)

    # Executa o pipeline de tratamento
    processar_arquivos_explode(
        pasta_origem=PASTA_ORIGEM,
        pasta_destino=PASTA_DESTINO,
        coluna_alvo=COLUNA_PARA_EXPLODIR
    )

    logger.info("Script run_explode_bm_nfse finalizado com sucesso.")