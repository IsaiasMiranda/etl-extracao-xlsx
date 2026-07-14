# etl-medicao

Pipeline genérico de consolidação de medições: lê múltiplas planilhas `.xlsx` com layouts variados, localiza o cabeçalho dinamicamente por palavras-chave, padroniza colunas e tipos, exporta um CSV único e faz backup compactado (`.zip`) dos arquivos de origem processados, excluindo-os da origem.

## Arquivos

- [load_medicao.py](load_medicao.py) — lógica de extração, limpeza, tipagem e consolidação.
- [run_medicao.py](run_medicao.py) — ponto de entrada; define as pastas de origem/destino/backup e chama `processar_boletins`, com tratamento de erro fatal e medição de tempo de execução.

## Como executar

```bash
uv run etl-medicao/run_medicao.py
```

Ajuste `base_dir`, `pasta_destino` e `pasta_processada` no início de [run_medicao.py](run_medicao.py) antes de rodar.

## Funcionalidades

### `MAPEAMENTO_COLUNAS` (dicionário)

Tabela "de-para" aplicada **após** a normalização do nome de coluna (chaves já em minúsculas, sem acento, `_` no lugar de espaço/pontuação). Unifica variações como `mes_ano`/`data`/`data_envio` → `data`, `tipo_de_servico`/`servico`/`descricao_servico` → `servico`, `os`/`ordem_de_servico` → `ordem_servico`, `n_da_nota`/`nota` → `numero_nota`, `municipio`/`localidade` → `cidade`, `qtd` → `quantidade`, `valor_r_` → `valor`, `preco_unitario` → `valor_unitario`, `status_da_atividade` → `status`, `observacoes` → `observacao`, `responsavel_validacao` → `responsavel`.

### `ORDEM_PADRAO` (lista)

Ordem final das colunas no CSV consolidado: `arquivo_origem`, `data`, `ordem_servico`, `numero_nota`, `equipe`, `cidade`, `fornecedor`, `responsavel`, `servico`, `quantidade`, `valor_unitario`, `valor`, `status`, `observacao`. Colunas extras (não mapeadas) são ordenadas alfabeticamente e anexadas ao final.

### `COLUNAS_NUMERICAS` / `PALAVRAS_CHAVE_CABECALHO` / `MAX_LINHAS_VARREDURA_CABECALHO`

- `COLUNAS_NUMERICAS`: `quantidade`, `valor_unitario`, `valor` — convertidas para `float` com 2 casas decimais para evitar artefatos de ponto flutuante ao gravar como texto no CSV.
- `PALAVRAS_CHAVE_CABECALHO`: `nota`, `ordem`, `data`, `equipe`, `valor`, `status`, `municipio`, `localidade`, `servico` — usadas para localizar a linha de cabeçalho real.
- `MAX_LINHAS_VARREDURA_CABECALHO = 30`: limite de segurança para não varrer a planilha inteira procurando cabeçalho.

### `normalizar_coluna(nome) -> str`

Normaliza um nome de coluna: minúsculas, remove acentuação (via `unicodedata.normalize('NFD', ...)`), substitui `ç` por `c`, troca qualquer sequência de caracteres não alfanuméricos por `_`, remove `_` nas pontas. Retorna string vazia se o valor for `NaN`.

### `_localizar_linha_cabecalho(df_bruto) -> int`

Varre até `MAX_LINHAS_VARREDURA_CABECALHO` linhas do DataFrame bruto, normaliza cada linha inteira como texto e conta quantas `PALAVRAS_CHAVE_CABECALHO` aparecem nela. A primeira linha com **2 ou mais** palavras-chave é considerada o cabeçalho. Se nenhuma linha atingir esse critério, usa como *fallback* a primeira linha com mais de 4 células preenchidas. Retorna `-1` se nada for encontrado.

### `_remover_linha_de_total(df_tabela) -> DataFrame`

Remove a última linha da tabela se ela parecer ser uma linha de total: contém a palavra `"total"` em algum texto, **ou** mais da metade das células estão vazias.

### `_consolidar_colunas_duplicadas(colunas_normalizadas, contexto, log_avisos) -> list`

Evita colisão silenciosa de colunas: se duas colunas nomeadas normalizarem para o mesmo texto (ex.: duas colunas "Processo"), renomeia a partir da segunda ocorrência para `<nome>_dup1`, `<nome>_dup2`, etc., e registra um aviso em `log_avisos`. Colunas sem nome (célula de cabeçalho vazia) não geram aviso aqui — são descartadas posteriormente.

### `_extrair_tabela(df_bruto, contexto, log_avisos) -> DataFrame`

Extrai a tabela de dados de um DataFrame bruto:
1. Localiza o cabeçalho via `_localizar_linha_cabecalho` (registra `[SKIP]` em `log_avisos` se não encontrar).
2. Normaliza os nomes de coluna e resolve duplicatas via `_consolidar_colunas_duplicadas`.
3. Descarta colunas sem cabeçalho **por posição** (`iloc`), nunca por nome — evita que nomes duplicados residuais causem expansão combinatória de colunas ao selecionar por `df[[nome, nome]]`.
4. Remove colunas e linhas totalmente vazias.
5. Remove a linha de total via `_remover_linha_de_total`.

### `_tratar_tipos(df) -> DataFrame`

Converte as colunas listadas em `COLUNAS_NUMERICAS` para `float` (via `pd.to_numeric(..., errors='coerce')`), arredondando para 2 casas decimais.

### `_ler_abas_validas(caminho_arquivo) -> list[(str, DataFrame)]`

Abre o workbook **uma única vez** e retorna **apenas a primeira aba visível** encontrada (abas ocultas nunca concorrem à posição de "primeira aba"). Isso resolve casos como planilhas que têm uma aba de dados relevante seguida de abas auxiliares/rascunho (`Planilha1`) ou abas antigas ocultas.

### `processar_boletins(pasta_origem_input, pasta_destino_input, pasta_backup_input) -> None`

Função principal, dividida em 4 fases:

1. **Fase 1 — Leitura e extração**: lista `*.xlsx` em `pasta_origem` (ignorando `~$*` e arquivos já consolidados, identificados pelo padrão `medicao_consolidada_` no nome), lê a primeira aba visível de cada um via `_ler_abas_validas`, extrai a tabela via `_extrair_tabela`, adiciona `arquivo_origem`. Erros por aba/arquivo são capturados individualmente e registrados em `log_avisos` sem interromper o restante do lote.
2. **Fase 2 — Consolidação e mapeamento**: aplica `MAPEAMENTO_COLUNAS`, concatena todos os DataFrames (`sort=False`), aplica `_tratar_tipos`, reordena colunas conforme `ORDEM_PADRAO` (extras ordenadas alfabeticamente ao final, `arquivo_origem` sempre primeiro).
3. **Fase 3 — Exportação**: grava `pasta_destino/medicao_consolidada_geral_<timestamp>.csv` (separador `;`, `utf-8-sig`). Remove um arquivo de saída pré-existente de mesmo nome antes de gravar; aborta com mensagem clara se o arquivo estiver aberto em outro programa.
4. **Fase 4 — Backup e limpeza**: compacta todos os arquivos processados com sucesso num único `.zip` (`arquivo-processado_medicao_<timestamp>.zip`) em `pasta_backup`, depois **exclui** os arquivos originais da pasta de origem. Falhas ao excluir um arquivo individual são registradas sem interromper o processo.

Ao final, imprime um resumo (arquivos lidos/compactados, colunas padronizadas/extras, linhas extraídas vs. processadas, com alerta caso não batam).

## Formato de saída

Um único arquivo `.csv` por execução (`medicao_consolidada_geral_YYYYMMDDHHMMSS.csv`, separador `;`, `utf-8-sig`) e um `.zip` de backup dos arquivos de origem processados (`arquivo-processado_medicao_YYYYMMDDHHMMSS.zip`).
