import os
import time
import traceback
from pathlib import Path
from load_medicao import processar_boletins

def main():
    inicio_tempo = time.time()
    
    print("="*60)
    print("EXECUTOR - CONSOLIDAÇÃO MEDIÇÃO (INCREMENTAL)".center(60))
    print("="*60)

    # 1. Caminhos Base (Ajustados para o novo projeto 'medicao')
    base_dir = Path(r'D:\base-geral\base-medicao\source')
    
    pasta_origem  = base_dir
    pasta_destino = Path(r'D:\base-geral\base-medicao\base-consolidada')
    pasta_processada  = Path(r'D:\base-geral\base-medicao\arquivo-processado')

    # 2. Validação básica de diretórios
    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return

    print(f"Origem  -> {pasta_origem}")
    print(f"Destino -> {pasta_destino}")
    print(f"Processada  -> {pasta_processada}")
    print("\nIniciando o processamento dos dados... Por favor, aguarde.")
    
    # 3. Executa o motor de processamento
    try:
        processar_boletins(
            pasta_origem_input=str(pasta_origem),
            pasta_destino_input=str(pasta_destino),
            pasta_backup_input=str(pasta_processada)
        )
    except Exception as e:
        print("\n" + "!"*60)
        print("[❌] Ocorreu um erro fatal durante a execução do script:")
        print("Detalhes técnicos do erro:")
        print(traceback.format_exc())
        print("!"*60)

    # 4. Finalização
    tempo_total = time.time() - inicio_tempo
    minutos, segundos = divmod(tempo_total, 60)
    
    print(f"\n🚀 Pipeline finalizado em {int(minutos)}m {segundos:.1f}s.")
    print("="*60)

if __name__ == "__main__":
    main()