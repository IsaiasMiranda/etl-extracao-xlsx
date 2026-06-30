import os
from load_siga_medicao import processar_boletins

def main():
    print("="*60)
    print("      EXECUTOR - CONSOLIDAÇÃO SIGA (INCREMENTAL)")
    print("="*60)

    # 1. Caminhos "chumbados" (Hardcoded)
    # ATENÇÃO: Altere os 3 diretórios abaixo para o seu ambiente real
    pasta_origem  = r'D:\One-Drive\Amper Elinsa\Bases Siga - Documentos\Medições'
    pasta_destino = r'D:\One-Drive\Amper Elinsa\Bases Siga - Documentos\Medições\base-consolidada'
    pasta_backup  = r'D:\One-Drive\Amper Elinsa\Bases Siga - Documentos\Medições\arquivo-processado'

    # 2. Validação básica de diretórios
    if not os.path.exists(pasta_origem):
        print(f"\n❌ Erro crítico: A pasta de origem não foi encontrada:\n{pasta_origem}")
        return

    print(f"Origem  -> {pasta_origem}")
    print(f"Destino -> {pasta_destino}")
    print(f"Backup  -> {pasta_backup}")
    print("\nIniciando o processamento dos dados... Por favor, aguarde.")
    print("-" * 60)
    
    # 3. Executa o motor de processamento incremental
    processar_boletins(pasta_origem, pasta_destino, pasta_backup)
    
    print("\nProcesso finalizado com sucesso.")

if __name__ == "__main__":
    main()