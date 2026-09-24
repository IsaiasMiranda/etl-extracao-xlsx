r"""Ponto de entrada da remocao de XLSX duplicados (producao e homologacao).

Uso:
    uv run etl-dedup-xlsx/run_dedup_xlsx.py --ambiente homologacao
    uv run etl-dedup-xlsx/run_dedup_xlsx.py --ambiente producao
    uv run etl-dedup-xlsx/run_dedup_xlsx.py --ambiente homologacao --raiz D:\outra\pasta

Mesmo padrao do etl-boletim-medicao: o ambiente e argumento e so muda a
raiz de dados; as subpastas e os parametros sao os mesmos nos dois.
Homologacao fica em BASE_GERAL\homologacao\<pipeline> (BASE_GERAL vem da
variavel de ambiente, padrao D:\base-geral).

Excecao: em producao a raiz continua sendo a pasta da area de trabalho
(comportamento anterior mantido), nao BASE_GERAL.
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import traceback
from pathlib import Path

from load_dedup_xlsx import remover_duplicados_xlsx

# Evita UnicodeEncodeError ao imprimir emojis/acentos quando o console usa
# um codepage legado (ex.: cp1252) em vez de UTF-8 (ex.: saída redirecionada).
sys.stdout.reconfigure(errors="replace")

BASE_GERAL = Path(os.environ.get("BASE_GERAL", "D:/base-geral"))
RAIZ_POR_AMBIENTE = {
    "producao":    Path(r"C:\Users\isaias.miranda\Desktop\xmls-dedup"),
    "homologacao": BASE_GERAL / "homologacao" / "dedup-xlsx",
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
    print("EXECUTOR - REMOÇÃO DE XLSX DUPLICADOS".center(60))
    print(f"Ambiente: {ambiente}".center(60))
    print("=" * 60)

    # 1. Caminhos base (subpastas iguais nos dois ambientes)
    pasta_origem = raiz
    pasta_duplicados = raiz / "arquivos-duplicados"
    pasta_unicos = raiz / "xmlx-unicos"

    # 2. Validação básica de diretórios
    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return 1

    print(f"Origem      -> {pasta_origem}")
    print(f"Duplicados  -> {pasta_duplicados}")
    print(f"Únicos      -> {pasta_unicos}")
    print("\nIniciando a análise de duplicados... Por favor, aguarde.")

    # 3. Executa o motor de processamento
    status = 0
    try:
        remover_duplicados_xlsx(
            pasta_origem_input=str(pasta_origem),
            pasta_duplicados_input=str(pasta_duplicados),
            pasta_unicos_input=str(pasta_unicos),
            limiar_similaridade_nome=0.85,
            dry_run=False,
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

    print(f"\n🚀 Script finalizado em {int(minutos)}m {segundos:.1f}s.")
    print("=" * 60)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
