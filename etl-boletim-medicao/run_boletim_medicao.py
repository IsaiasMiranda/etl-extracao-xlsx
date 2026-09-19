"""Ponto de entrada unico do ETL de boletim de medicao (producao e homologacao).

Uso:
    uv run etl-boletim-medicao/run_boletim_medicao.py --ambiente homologacao
    uv run etl-boletim-medicao/run_boletim_medicao.py --ambiente producao
    uv run etl-boletim-medicao/run_boletim_medicao.py --ambiente homologacao --raiz D:\outra\pasta

Ate 2026-09-19 producao/ e homologacao/ eram copias de diretorio que
diferiam em 1 caminho; desde entao ha um unico modulo e o ambiente e um
argumento (B3 da proposta de evolucao). A raiz de dados pode vir de
--raiz ou da variavel de ambiente BASE_GERAL (padrao D:\base-geral).
"""
from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path

from load_boletim_medicao import processar_boletins

BASE_GERAL = Path(os.environ.get("BASE_GERAL", "D:/base-geral"))
RAIZ_POR_AMBIENTE = {
    "producao":    BASE_GERAL / "base-boletim-medicao",
    "homologacao": BASE_GERAL / "homologacao" / "boletim-medicao",
}


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ambiente", choices=RAIZ_POR_AMBIENTE, required=True)
    ap.add_argument("--raiz", type=Path, help="raiz do dominio (sobrepoe o padrao do ambiente)")
    a = ap.parse_args(argv)
    raiz = a.raiz or RAIZ_POR_AMBIENTE[a.ambiente]

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    log = logging.getLogger(__name__)
    log.info("Ambiente %s | raiz %s", a.ambiente, raiz)

    # Mesmos parametros nos 2 ambientes desde a promocao de 2026-09-02.
    processar_boletins(
        pasta_origem=raiz / "source",
        pasta_destino=raiz / "base-consolidado",
        pasta_processados=raiz / "arquivos-processados",
        normalizar_deslocamentos=True,
        caminho_auditoria=raiz / "auditoria" / "boletim-medicao-normalizacao.csv",
        remover_sem_boletim=True,
        excluir_boletim_nfse_duplicado=True,
        remover_boletim_sem_valor_duplicado=True,
    )
    log.info("Script finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
