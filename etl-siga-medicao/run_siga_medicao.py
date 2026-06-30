import os
import time
import traceback
from pathlib import Path
from load_siga_medicao import processar_boletins

def main():
    inicio_tempo = time.time()
    
    print("="*60)
    print("EXECUTOR - CONSOLIDAÇÃO SIGA (INCREMENTAL)".center(60))
    print("="*60)

    # 1. Caminhos Base (Utilizando pathlib para maior segurança)
    base_dir = Path(r'D:\One-Drive\Amper Elinsa\Bases Siga - Documentos\Medições')
    
    pasta_origem  = base_dir
    pasta_destino = base_dir / 'base-consolidada'
    pasta_backup  = base_dir / 'arquivo-processado'

    # 2. Validação básica de diretórios
    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return

    print(f"Origem  -> {pasta_origem}")
    print(f"Destino -> {pasta_destino}")
    print(f"Backup  -> {pasta_backup}")
    print("\nIniciando o processamento dos dados... Por favor, aguarde.")
    
    # 3. Executa o motor de processamento, capturando eventuais quebras críticas com detalhes
    try:
        processar_boletins(
            pasta_origem_input=str(pasta_origem),
            pasta_destino_input=str(pasta_destino),
            pasta_backup_input=str(pasta_backup)
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