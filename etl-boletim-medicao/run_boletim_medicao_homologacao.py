"""Executa somente a saída normalizada do piloto de boletim."""

import logging
from pathlib import Path

from load_boletim_medicao import processar_boletins


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)


if __name__ == '__main__':
    raiz = Path(r'D:\base-geral\homologacao\boletim-medicao')
    processar_boletins(
        pasta_origem=raiz / 'source',
        pasta_destino=raiz / 'base-normalizada',
        pasta_processados=raiz / 'arquivos-processados',
        normalizar_deslocamentos=True,
        caminho_auditoria=raiz / 'auditoria' / 'boletim-medicao-normalizacao.csv',
        remover_sem_boletim=True,
        excluir_boletim_nfse_duplicado=True,
        remover_boletim_sem_valor_duplicado=True,
    )
