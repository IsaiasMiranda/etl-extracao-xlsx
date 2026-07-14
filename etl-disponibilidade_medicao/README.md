# etl-disponibilidade_medicao

Lê planilhas de Medição de Disponibilidade em layout "largo" (uma coluna por dia do mês), transforma esse layout em formato longo (*unpivot*) e consolida tudo num único CSV, copiando os arquivos de origem processados para uma pasta de histórico.

## Arquivos

- [load_disponibilidade_medicao.py](load_disponibilidade_medicao.py) — lógica de extração, unpivot e consolidação.
- [run_disponibilidade_medicao.py](run_disponibilidade_medicao.py) — ponto de entrada; define as pastas de origem/destino/processados e chama `processar_boletins`.

## Como executar

```bash
uv run etl-disponibilidade_medicao/run_disponibilidade_medicao.py
```

Ajuste os caminhos `origem`, `arquivo_consolidado` e `arquivo_processado` no início de [run_disponibilidade_medicao.py](run_disponibilidade_medicao.py) antes de rodar.

## Modelo de dados final

Colunas do CSV consolidado (`ORDEM_PADRAO`): `arquivo_origem`, `tipo_equipe`, `Equipe`, `Faltas`, `Data`, `Qtd_Horas`, `Valor_Hora`, `Total_Geral`.

## Funcionalidades

### `_extrair_tipo_equipe(nome_arquivo) -> str`

Identifica a categoria macro do arquivo a partir do nome: `SEED MONEY`, `PERDAS` ou `SMC` (busca por substring, case-insensitive). Se nenhuma dessas categorias for encontrada, usa como *fallback* o nome do arquivo limpo (removendo prefixo numérico `NN_NN` e o texto `Medição_Disponibilidade_`).

### `_localizar_linha_cabecalho(ws) -> int`

Varre as primeiras 10 linhas da planilha procurando uma célula cujo valor (normalizado) seja exatamente `"equipes"`. Retorna o número da linha (1-based) ou `-1` se não encontrar.

### `_processar_aba(ws, nome_arquivo) -> DataFrame | None`

Processa uma aba:
1. Determina `tipo_equipe` via `_extrair_tipo_equipe`.
2. Localiza a linha de cabeçalho (deve conter `Equipes`) e, dentro dela, as colunas `Equipes` e `Faltas`.
3. Identifica todas as colunas de cabeçalho cujo valor é uma data (`hasattr(h, 'strftime')`) — essas são as colunas de dias do mês.
4. **Correção de deslocamento**: como o template sempre reserva 31 colunas para dias (independente do mês ter 28–31 dias), as colunas `Total Geral` e `VLR_HR` são localizadas por deslocamento fixo (`+31` e `+32`) a partir da primeira coluna de data — não pelo nome do cabeçalho.
5. Para cada linha cuja `Equipe` comece com o prefixo `"PA-"` (demais linhas são ignoradas), gera **um registro por dia** (unpivot), com `Data`, `Qtd_Horas`, `Valor_Hora` e `Total_Geral` repetidos.

Retorna `None` se o cabeçalho, as colunas `Equipes`/`Faltas` ou nenhuma coluna de data forem encontrados, ou se não houver registros válidos.

### `_processar_arquivo(caminho, nome_saida=None) -> (list[DataFrame], bool)`

Abre o workbook completo (não em modo `read_only`, pois precisa ler valores de data/formatação) e processa todas as abas visíveis via `_processar_aba`. Antes de processar, normaliza o nome do arquivo removendo espaços irregulares antes da extensão `.xlsx`.

### `montar_modelo(caminho_arquivo, sheet_name=None) -> DataFrame`

API pública auxiliar (não usada pelo pipeline principal, útil para uso ad-hoc/notebooks): processa um único arquivo — todas as abas visíveis ou uma aba específica — e retorna um único DataFrame já ordenado por `Equipe`/`Data`. Lança `ValueError` se nenhum dado válido for extraído.

### `processar_boletins(pasta_origem, pasta_destino, pasta_processados) -> None`

Função principal, orquestra o pipeline completo:

1. **Descoberta**: busca `*.xlsx` recursivamente (`rglob`) em `pasta_origem`, ignorando arquivos de trava (`~$*`) e consolidados já existentes em `pasta_destino` (`medicao-disponibilidade-consolidada_*.csv`).
2. **Extração**: processa cada arquivo via `_processar_arquivo` (mantendo o nome original de saída, sem deduplicação de nome).
3. **Consolidação**: concatena os DataFrames extraídos e ordena por `Equipe`/`Data`.
4. **Exportação**: grava `pasta_destino/medicao-disponibilidade-consolidada_<timestamp>.csv` (separador `;`, `utf-8-sig`).
5. **Arquivamento**: **copia** (`shutil.copy2`, preservando metadados) os arquivos processados para `pasta_processados`, **sobrescrevendo** (remove antes de copiar) se já existir um arquivo de mesmo nome. Ao contrário de `etl-boletim-medicao`, os arquivos de origem **não são removidos** da pasta de origem.
6. **Relatório**: imprime via `logging` um resumo com total de arquivos, consolidados, ignorados/erros, registros gerados e salvos.

## Formato de saída

Um único arquivo `.csv` por execução (`medicao-disponibilidade-consolidada_YYYYMMDDHHMMSS.csv`), em formato longo (uma linha por Equipe/Data), separador `;`, codificação `utf-8-sig`.
