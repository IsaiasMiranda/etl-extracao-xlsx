# etl-siga-medicao

Consolida planilhas exportadas do sistema SIGA, padroniza colunas e **segmenta a saída por mês/ano** (um CSV por período), com backup compactado (`.zip`) dos arquivos de origem processados, excluindo-os da origem.

## Arquivos

- [load_siga_medicao.py](load_siga_medicao.py) — lógica de extração, limpeza, segmentação e consolidação.
- [run_siga_medicao.py](run_siga_medicao.py) — ponto de entrada; define as pastas de origem/destino/backup e chama `processar_boletins`, com tratamento de erro fatal e medição de tempo de execução.

## Como executar

```bash
uv run etl-siga-medicao/run_siga_medicao.py
```

Ajuste `base_dir`, `pasta_destino` e `pasta_processada` no início de [run_siga_medicao.py](run_siga_medicao.py) antes de rodar.

## Funcionalidades

### `normalizar_coluna(nome) -> str`

Normaliza um nome de coluna: minúsculas, remove acentuação, substitui `ç` por `c`, troca sequências de caracteres não alfanuméricos por `_`, remove `_` nas pontas. Retorna string vazia se o valor for `NaN`. (Implementação idêntica à de [etl-medicao](../etl-medicao/README.md).)

### `_extrair_tabela_siga(df_bruto) -> DataFrame`

Localiza a linha de cabeçalho procurando, linha a linha, a presença simultânea dos textos `"ordem de serviço"` e `"status da atividade"`. A partir dessa linha, normaliza os nomes de coluna, descarta a linha de cabeçalho original dos dados, e remove linhas/colunas totalmente vazias. Retorna um DataFrame vazio se o cabeçalho não for encontrado.

### `_limpar_dataframe(df) -> DataFrame`

Descarta colunas sem nome útil (`unnamed*`, vazia, `"nan"`/`"none"`/`"<na>"`) e remove colunas duplicadas (mantendo a primeira ocorrência).

### `processar_boletins(pasta_origem_input, pasta_destino_input, pasta_backup_input) -> None`

Função principal, dividida em 4 fases:

1. **Fase 1 — Leitura e extração**: lista `*.xlsx` em `pasta_origem` (ignorando `~$*` e arquivos já consolidados, identificados pelo padrão `siga_consolidado_` no nome). Para cada arquivo, obtém a lista de abas via `pd.ExcelFile`, e para **cada aba** verifica sua visibilidade abrindo o workbook separadamente com `openpyxl` (`sheet_state == 'visible'`) antes de ler os dados com `pd.read_excel(..., header=None, dtype=str)`. Extrai a tabela via `_extrair_tabela_siga`, limpa com `_limpar_dataframe` e adiciona `arquivo_origem`.
2. **Fase 2 — Consolidação e limpeza final**: concatena todos os DataFrames, aplica `_limpar_dataframe` novamente sobre o resultado consolidado, e reordena as colunas — prioriza a lista fixa `colunas_obrigatorias` (`arquivo_origem`, `data`, `ordem_de_servico`, `origem`, `id_da_atividade`, `status_da_atividade`, `numero_da_nota`, `numero_ocorrencia`, `numero_da_conta`, `agrupamento_id`, `code`, `cidade`, `bairro`, `latitude`, `longitude`), com as demais colunas anexadas ao final na ordem em que aparecem.
3. **Fase 3 — Segmentação e exportação**: converte a coluna `data` para `datetime` (`format='mixed', dayfirst=True`, erros viram `NaT`) e deriva `Mes_Ano` no formato `MM-YYYY` (linhas sem data válida caem no grupo `"Sem_Data"`). Para cada valor único de `Mes_Ano`, grava um arquivo separado `siga_consolidado_<Mes_Ano>_<data_de_hoje>.csv` (separador `;`, `utf-8-sig`), removendo um arquivo pré-existente de mesmo nome antes de gravar (aborta esse arquivo específico, sem interromper os demais, se estiver aberto em outro programa).
4. **Fase 4 — Backup e auditoria**: compacta todos os arquivos processados com sucesso num único `.zip` (`arquivo-processado_<timestamp>.zip`) em `pasta_backup` e **exclui cada arquivo original imediatamente após adicioná-lo ao zip** (diferente de `etl-medicao`, que só exclui depois de fechar o zip inteiro).

Ao final, imprime um resumo (arquivos lidos/compactados, colunas mapeadas, linhas extraídas vs. processadas, com alerta caso não batam).

## Formato de saída

Um arquivo `.csv` **por mês/ano** presente nos dados (`siga_consolidado_MM-YYYY_YYYY_MM_DD.csv`, separador `;`, `utf-8-sig`), mais um `.zip` de backup dos arquivos de origem processados (`arquivo-processado_YYYYMMDDHHMMSS.zip`).
