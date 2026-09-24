r"""Ponto de entrada do ETL de medicao de disponibilidades (producao e homologacao).

Uso:
    uv run etl-disponibilidade_medicao/run_disponibilidade_medicao.py --ambiente homologacao
    uv run etl-disponibilidade_medicao/run_disponibilidade_medicao.py --ambiente producao
    uv run etl-disponibilidade_medicao/run_disponibilidade_medicao.py --ambiente homologacao --raiz D:\outra\pasta

Mesmo padrao do etl-boletim-medicao: o ambiente e argumento e so muda a
raiz de dados; as subpastas e os parametros sao os mesmos nos dois.
Homologacao fica em BASE_GERAL\homologacao\<pipeline> (BASE_GERAL vem da
variavel de ambiente, padrao D:\base-geral).
"""
from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path

from load_disponibilidade_medicao import processar_boletins

BASE_GERAL = Path(os.environ.get("BASE_GERAL", "D:/base-geral"))
RAIZ_POR_AMBIENTE = {
    "producao":    BASE_GERAL / "base-disponibilidade-medicao",
    "homologacao": BASE_GERAL / "homologacao" / "disponibilidade-medicao",
}


def _args(argv):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ambiente", choices=RAIZ_POR_AMBIENTE, required=True)
    ap.add_argument("--raiz", type=Path, help="raiz do dominio (sobrepoe o padrao do ambiente)")
    a = ap.parse_args(argv)
    return a.ambiente, a.raiz or RAIZ_POR_AMBIENTE[a.ambiente]


def main(argv: list[str] | None = None) -> int:
    ambiente, raiz = _args(argv)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    log = logging.getLogger(__name__)
    log.info("Ambiente %s | raiz %s", ambiente, raiz)

    origem = raiz / "source" / "MEDICAO_DISPONIBILIDADES - Documentos"
    arquivo_consolidado = raiz / "arquivo-consolidado"
    arquivo_processado = raiz / "arquivo-processado"

    log.info("Iniciando rotina de consolidação e unpivot de Disponibilidades...")
    processar_boletins(
        pasta_origem=origem,
        pasta_destino=arquivo_consolidado,
        pasta_processados=arquivo_processado,
    )
    log.info("Script finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
