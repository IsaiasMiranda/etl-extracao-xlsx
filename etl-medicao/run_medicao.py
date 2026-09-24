r"""Ponto de entrada do ETL de medicao (producao e homologacao).

Uso:
    uv run etl-medicao/run_medicao.py --ambiente homologacao
    uv run etl-medicao/run_medicao.py --ambiente producao
    uv run etl-medicao/run_medicao.py --ambiente homologacao --raiz D:\outra\pasta

Mesmo padrao do etl-boletim-medicao: o ambiente e argumento e so muda a
raiz de dados; as subpastas e os parametros sao os mesmos nos dois.
Homologacao fica em BASE_GERAL\homologacao\<pipeline> (BASE_GERAL vem da
variavel de ambiente, padrao D:\base-geral).
"""
from __future__ import annotations

import argparse
import os
import time
import traceback
from pathlib import Path

from load_medicao import processar_boletins

BASE_GERAL = Path(os.environ.get("BASE_GERAL", "D:/base-geral"))
RAIZ_POR_AMBIENTE = {
    "producao":    BASE_GERAL / "base-medicao",
    "homologacao": BASE_GERAL / "homologacao" / "medicao",
}


def _args(argv):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--ambiente", choices=RAIZ_POR_AMBIENTE, required=True)
    ap.add_argument("--raiz", type=Path, help="raiz do dominio (sobrepoe o padrao do ambiente)")
    a = ap.parse_args(argv)
    return a.ambiente, a.raiz or RAIZ_POR_AMBIENTE[a.ambiente]


def main(argv: list[str] | None = None) -> int:
    ambiente, raiz = _args(argv)
    inicio_tempo = time.time()

    print("=" * 60)
    print("EXECUTOR - CONSOLIDAÇÃO MEDIÇÃO (INCREMENTAL)".center(60))
    print(f"Ambiente: {ambiente}".center(60))
    print("=" * 60)

    # 1. Caminhos base (subpastas iguais nos dois ambientes)
    pasta_origem = raiz / "source"
    pasta_destino = raiz / "base-consolidada"
    pasta_processada = raiz / "arquivo-processado"

    # 2. Validação básica de diretórios
    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return 1

    print(f"Origem  -> {pasta_origem}")
    print(f"Destino -> {pasta_destino}")
    print(f"Processada  -> {pasta_processada}")
    print("\nIniciando o processamento dos dados... Por favor, aguarde.")

    # 3. Executa o motor de processamento
    status = 0
    try:
        processar_boletins(
            pasta_origem_input=str(pasta_origem),
            pasta_destino_input=str(pasta_destino),
            pasta_backup_input=str(pasta_processada),
        )
    except Exception:
        status = 1
        print("\n" + "!" * 60)
        print("[❌] Ocorreu um erro fatal durante a execução do script:")
        print("Detalhes técnicos do erro:")
        print(traceback.format_exc())
        print("!" * 60)

    # 4. Finalização
    tempo_total = time.time() - inicio_tempo
    minutos, segundos = divmod(tempo_total, 60)

    print(f"\n🚀 Pipeline finalizado em {int(minutos)}m {segundos:.1f}s.")
    print("=" * 60)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
