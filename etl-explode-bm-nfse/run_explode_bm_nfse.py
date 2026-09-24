r"""Ponto de entrada do ETL que explode a coluna boletim de BM x NFS-e (producao e homologacao).

Uso:
    uv run etl-explode-bm-nfse/run_explode_bm_nfse.py --ambiente homologacao
    uv run etl-explode-bm-nfse/run_explode_bm_nfse.py --ambiente producao
    uv run etl-explode-bm-nfse/run_explode_bm_nfse.py --ambiente homologacao --raiz D:\outra\pasta

Mesmo padrao do etl-boletim-medicao: o ambiente e argumento e so muda a
raiz de dados; as subpastas e os parametros sao os mesmos nos dois.
Homologacao fica em BASE_GERAL\homologacao\<pipeline> (BASE_GERAL vem da
variavel de ambiente, padrao D:\base-geral).

Excecao: em producao a raiz e a propria pasta do script (comportamento
anterior mantido), nao BASE_GERAL.
"""
from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path

from load_explode_bm_nfse import processar_arquivos_explode

BASE_GERAL = Path(os.environ.get("BASE_GERAL", "D:/base-geral"))
RAIZ_POR_AMBIENTE = {
    "producao":    Path(__file__).parent.resolve(),
    "homologacao": BASE_GERAL / "homologacao" / "explode-bm-nfse",
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

    # Origem e destino sao a mesma pasta (em producao, a do proprio script).
    pasta_origem = raiz
    pasta_destino = raiz
    coluna_para_explodir = "boletim"

    log.info("Iniciando rotina de separação de boletins...")
    processar_arquivos_explode(
        pasta_origem=pasta_origem,
        pasta_destino=pasta_destino,
        coluna_alvo=coluna_para_explodir,
    )
    log.info("Script finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
