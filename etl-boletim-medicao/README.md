# etl-boletim-medicao

Consolida boletins de medição (planilhas `.xlsx` com múltiplas abas) numa única planilha `.xlsx`, padronizando os nomes de coluna e arquivando os arquivos de origem já processados.

## Estrutura

Um único módulo serve produção e homologação; o ambiente é argumento do
runner. Qualquer melhoria (ex.: nova variação de cabeçalho no `MAPEAMENTO`)
vale para os dois automaticamente.

```
etl-boletim-medicao/
  load_boletim_medicao.py   — extração/consolidação (colunas restritas a ORDEM_PADRAO)
  normalizar_boletim.py     — normalização auditável de valores deslocados
  run_boletim_medicao.py    — ponto de entrada único: --ambiente producao|homologacao
  test_normalizar_boletim.py — testes unitários de normalizar_boletim.py (uv run pytest)
  validar_formato_colunas.py — relatório read-only de conformidade de formato por coluna
```

## Como executar

```bash
uv run etl-boletim-medicao/run_boletim_medicao.py --ambiente homologacao
uv run etl-boletim-medicao/run_boletim_medicao.py --ambiente producao
# raiz de dados: D:/base-geral (ou variável BASE_GERAL); --raiz sobrepõe
```

Ajuste os caminhos `origem`, `destino` e `processados` no início de cada
`run_*.py` antes de rodar.

## Funcionalidades (comuns aos dois módulos)

### `MAPEAMENTO` (dicionário)

Tabela "de-para" que unifica todas as variações de grafia observadas nos boletins (com/sem acento, abreviações, quebras de linha no cabeçalho) para um nome de coluna canônico. Cobre os campos: `boletim`, `nota_fiscal`, `valor`, `data_envio`, `origem_lancamento`, `municipio`, `competencia`, `descricaonotafiscal`, `cp`, `tipomedicao`, `estrutura`, `medicao`, `domiciliofiscal`, `codigotarifafiscal`, além de campos já padronizados (`distribuidora`, `regional`, `parceiro`, `equipe`, `iva`).

### `ORDEM_PADRAO` (lista)

Define a ordem final das colunas no arquivo consolidado.

- O consolidado é **restrito às colunas de `ORDEM_PADRAO`** (+ `arquivo_origem`); coluna extra não mapeada é descartada — o layout precisa ser canônico antes da correção de deslocamentos rodar. (Até 2026-09-02 produção anexava colunas extras ao final; esse comportamento não existe mais.)

### `_localizar_linha_cabecalho(df_bruto) -> int`

Varre as linhas de um DataFrame bruto (sem cabeçalho definido) procurando a linha que contenha simultaneamente os textos `distribuidora` e `regional` — essa é considerada a linha real de cabeçalho da tabela. Retorna `-1` se não encontrar.

### `_processar_aba(sheet, nome_arquivo) -> DataFrame | None`

Processa uma aba (`worksheet`) individual:
1. Converte a aba inteira em DataFrame e localiza a linha de cabeçalho via `_localizar_linha_cabecalho`.
2. Usa essa linha como nomes de coluna (normalizados para minúsculas/sem espaços nas pontas) e descarta tudo antes dela.
3. Remove linhas e colunas totalmente vazias, remove colunas duplicadas.
4. Aplica o `MAPEAMENTO` para renomear colunas.
5. Adiciona a coluna `arquivo_origem` com o nome do arquivo de origem.
6. Descarta colunas sem nome (`unnamed*`, vazias, `nan`/`none`).

Retorna `None` se a aba não tiver cabeçalho reconhecível ou ficar vazia após a limpeza.

### `_processar_arquivo(caminho, nome_saida=None) -> (list[DataFrame], bool)`

Abre um arquivo `.xlsx` em modo somente leitura (`read_only=True`) e processa todas as abas visíveis (`sheet_state == 'visible'`) através de `_processar_aba`. Abas ocultas são ignoradas. Retorna a lista de DataFrames extraídos e um booleano indicando se ao menos um dado foi extraído com sucesso. Erros de abertura de arquivo (`InvalidFileException`, `PermissionError`, `OSError`) são capturados e logados sem interromper o pipeline.

### `processar_boletins(pasta_origem, pasta_destino, pasta_processados=None, ...) -> None`

Função principal, orquestra o pipeline completo:

1. **Descoberta**: lista arquivos `*.xlsx` em `pasta_origem`, ignorando arquivos de trava do Excel (`~$*`) e consolidados já existentes em `pasta_destino` (padrão `boletim-medicao-consolidado_*.xlsx`).
2. **Extração**: para cada arquivo, verifica se já existe um arquivo de mesmo nome em `pasta_processados`; se sim, prefixa o nome de saída com um número aleatório (1000–9999) para evitar sobrescrita. Processa o arquivo via `_processar_arquivo`.
3. **Consolidação**: concatena todos os DataFrames extraídos e reordena colunas conforme `ORDEM_PADRAO`. Esse layout já canônico é o que a etapa de correção de deslocamentos (abaixo) recebe como entrada, não a ordem crua de concatenação.
4. **Exportação**: grava o resultado em `pasta_destino/boletim-medicao-consolidado_<timestamp>.xlsx` (aba `Boletins`). Se o arquivo de destino estiver aberto em outro programa (`PermissionError`), aborta com mensagem de erro.
5. **Arquivamento**: **move** (`shutil.move`) os arquivos de origem processados com sucesso para `pasta_processados`. Arquivos sem dados válidos permanecem na origem para inspeção manual.
6. **Relatório**: registra via `logging` um resumo com total de arquivos, arquivos processados/com erro, linhas extraídas vs. salvas, e alerta se as duas contagens não baterem (possível eliminação de duplicatas/linhas vazias).

Se `pasta_processados` não for informado, usa `pasta_origem/arquivos-processados` como padrão.

## Formato de saída

Um único arquivo `.xlsx` por execução, nomeado `boletim-medicao-consolidado_YYYYMMDDHHMMSS.xlsx`, contendo uma aba `Boletins` com todas as linhas extraídas de todos os arquivos/abas processados.

## Normalização de deslocamentos e filtros de qualidade

O runner ativa, nos dois ambientes, `normalizar_deslocamentos=True`: gera a
planilha normalizada e o relatório CSV de auditoria em `<raiz>/auditoria/`.
A realocação só ocorre quando a assinatura da coluna destino encontra
exatamente um candidato; casos ambíguos são registrados como `AMBIGUA` e
permanecem sem movimentação.

Também ativa `remover_sem_boletim=True` (linha sem número de boletim depois
da normalização é descartada — o grão do domínio é 1 linha = 1 boletim),
`excluir_boletim_nfse_duplicado=True` e `remover_boletim_sem_valor_duplicado=True`.

Histórico: até 2026-09-19 existiam duas cópias do módulo (`producao/` e
`homologacao/`) que diferiam em um único caminho; as regras acima foram
validadas em homologação e promovidas para produção em 2026-09-02, e as
cópias foram unificadas em 2026-09-19 (o ambiente virou argumento do runner).
