import sys
import time
import traceback
from pathlib import Path

from load_dedup_xlsx import remover_duplicados_xlsx

# Evita UnicodeEncodeError ao imprimir emojis/acentos quando o console usa
# um codepage legado (ex.: cp1252) em vez de UTF-8 (ex.: saída redirecionada).
sys.stdout.reconfigure(errors='replace')


def main():
    inicio_tempo = time.time()

    print("=" * 60)
    print("EXECUTOR - REMOÇÃO DE XLSX DUPLICADOS".center(60))
    print("=" * 60)

    # 1. Caminhos Base
    pasta_origem = Path(r'C:\Users\isaias.miranda\Desktop\xmls-dedup')
    pasta_duplicados = pasta_origem / 'arquivos-duplicados'
    pasta_unicos = pasta_origem / 'xmlx-unicos'

    # 2. Validação básica de diretórios
    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return

    print(f"Origem      -> {pasta_origem}")
    print(f"Duplicados  -> {pasta_duplicados}")
    print(f"Únicos      -> {pasta_unicos}")
    print("\nIniciando a análise de duplicados... Por favor, aguarde.")

    # 3. Executa o motor de processamento
    try:
        remover_duplicados_xlsx(
            pasta_origem_input=str(pasta_origem),
            pasta_duplicados_input=str(pasta_duplicados),
            pasta_unicos_input=str(pasta_unicos),
            limiar_similaridade_nome=0.85,
            dry_run=False,
        )
    except Exception:
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


if __name__ == "__main__":
    main()
