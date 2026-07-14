# CLAUDE.md

Este arquivo fornece orientações ao Claude Code (claude.ai/code) para trabalhar com o código deste repositório.

## Visão geral do projeto

`etl-extracao-xlsx` é um conjunto de pipelines ETL em Python, independentes entre si, que extraem dados de planilhas Excel produzidas manualmente (boletins de medição, relatórios de disponibilidade, exportações do sistema SIGA), limpam/padronizam esses dados e consolidam tudo em um único arquivo CSV/XLSX pronto para carga em banco de dados ou análise. Veja o [README.md](README.md) para a descrição completa de domínio de cada pipeline.

Não há um ponto de entrada único para a aplicação — o `main.py` na raiz do repositório é um placeholder sem uso real. Cada pipeline vive na sua própria pasta `etl-*/` e é executado de forma independente.

## Comandos

```bash
uv sync                                            # instala/trava as dependências em .venv/
uv run etl-boletim-medicao/run_boletim_medicao.py  # executa um pipeline específico
```

O mesmo padrão (`uv run etl-<nome>/run_<nome>.py`) vale para os cinco pipelines: `etl-boletim-medicao`, `etl-disponibilidade_medicao`, `etl-explode-bm-nfse`, `etl-medicao`, `etl-siga-medicao`.

Não há suíte de testes, linter ou etapa de build configurada neste repositório — não invente comandos para isso.

## Arquitetura

Cada pipeline em `etl-*/` é dividido em exatamente dois arquivos, seguindo a mesma convenção `load_*` / `run_*`:

- **`load_<nome>.py`** — uma biblioteca autocontida com toda a lógica de negócio: detecção de cabeçalho, mapeamento de nomes de coluna, limpeza de linhas, e uma função pública de orquestração `processar_boletins(...)` (ou nome equivalente). É aqui que ficam as regras específicas de cada pipeline e onde mudanças na lógica de extração/limpeza devem ser feitas.
- **`run_<nome>.py`** — um ponto de entrada executável, enxuto. Define os caminhos fixos de origem/destino/backup (`Path(r'D:\...')`, caminhos locais/de rede específicos da máquina) e chama a função de orquestração do `load_*.py` correspondente. Para ajustar onde um pipeline lê/grava, edite aqui, não no `load_*.py`.

Os dois arquivos se importam por nome de módulo simples (`from load_boletim_medicao import processar_boletins`), não por caminho de pacote — cada pasta `etl-*/` é tratada como seu próprio diretório de trabalho, não como um pacote Python. Os scripts precisam ser executados tendo essa pasta como raiz de módulo resolvível (`uv run etl-x/run_x.py` funciona porque o Python adiciona o diretório do script ao `sys.path`).

### Fluxo comum dos pipelines

Apesar de cada pipeline ter suas particularidades de layout, todos seguem o mesmo fluxo de quatro fases dentro do `load_*.py`:

1. **Descoberta** — busca arquivos `*.xlsx` (às vezes `*.xls`) na pasta de origem, ignora arquivos de trava do Excel (`~$*`) e saídas consolidadas já geradas, ignora abas ocultas (`sheet_state != 'visible'`).
2. **Extração** — abre cada workbook com `openpyxl`, localiza dinamicamente a linha real de cabeçalho (varrendo em busca de combinações de palavras-chave conhecidas, já que o cabeçalho raramente está na linha 1 — logos/textos soltos ficam acima dele), e monta um `pandas.DataFrame` a partir desse ponto.
3. **Transformação** — normaliza nomes de coluna (minúsculas, sem acento, pontuação virando `_`) e aplica um dicionário `MAPEAMENTO`/`MAPEAMENTO_COLUNAS` específico do pipeline, que unifica variações conhecidas de grafia/nome do mesmo campo num único nome de coluna canônico, depois reordena as colunas conforme uma lista fixa `ORDEM_PADRAO` (colunas extras não mapeadas são anexadas no final, não descartadas).
4. **Carga** — concatena todos os DataFrames extraídos, grava um único arquivo de saída com timestamp, e arquiva os arquivos de origem processados com sucesso (movidos, copiados ou compactados em zip, dependendo do pipeline — ver tabela abaixo) para que uma nova execução não os reprocesse.

Desvios específicos de cada pipeline em relação a esse fluxo, úteis na hora de depurar um deles:

| Pipeline | Saída | Arquivamento dos arquivos de origem |
|---|---|---|
| `etl-boletim-medicao` | um único `.xlsx` | **movidos**; nome de arquivo duplicado recebe um prefixo numérico aleatório em vez de ser sobrescrito |
| `etl-disponibilidade_medicao` | um único `.csv` | **copiados** (`shutil.copy2`, origem permanece intacta); também faz unpivot de uma coluna por dia em linhas longas |
| `etl-explode-bm-nfse` | um `.csv` por arquivo de entrada (`*_explodido.csv`) | não arquivado; origem e saída ficam na mesma pasta |
| `etl-medicao` | um único `.csv` | arquivos de origem compactados em backup `.zip`, depois **excluídos** da origem |
| `etl-siga-medicao` | um `.csv` **por mês/ano** (segmentado por uma coluna `data` interpretada) | arquivos de origem compactados em backup `.zip`, depois **excluídos** da origem |

Quando um pipeline exclui ou move seus arquivos de origem, executá-lo novamente sobre a mesma pasta não reprocessa o que já foi arquivado — isso é idempotência intencional, não um bug a "corrigir" caso você veja menos arquivos processados numa segunda execução.

## Histórico de mudanças

Resumo de cada commit no histórico do projeto, do mais antigo para o mais recente.

### 2026-06-30 — `6a59d82` / `33626eb` — Commit inicial
Estrutura inicial do repositório: quatro pipelines (`etl-boletim-medicao`, `etl-explode-bm-nfse`, `etl-medicao-disponibilidade`, `etl-siga-medicao`), cada um com seu `load_*.py`/`run_*.py`. Incluía também arquivos de rascunho (`* copy.py`, `* - Copia.py`) e um arquivo de exemplo `modelo-final/tabela_medicao.xlsx`.

### 2026-06-30 — `f0c3f49` — Merge
Merge da branch `main` remota após o commit inicial (sem alterações de conteúdo).

### 2026-06-30 — `a2dfcc1` — Alterações na estrutura de saída
Ajustes no `.gitignore` e no pipeline de medição de disponibilidade (nome do módulo na época: `load_medicao_disponibilidade.py`).

### 2026-07-02 — `9bf10ad` — Copiar em vez de mover
No pipeline de disponibilidade, trocado o comportamento de **mover** os arquivos de origem processados para **copiar** (`shutil.copy2`) para a pasta de destino, sobrescrevendo o arquivo existente em vez de falhar — os arquivos de origem passam a permanecer intactos após o processamento.

### 2026-07-02 — `8e11f24` — Melhorias na extração SIGA/medição
Limpeza de arquivos de rascunho duplicados (`* - Copia.py`, `* copy.py`, `run_* copy.py`) em `etl-boletim-medicao` e `etl-siga-medicao`. Ajustes nos scripts `load_siga_medicao.py` e nos respectivos `run_*.py` (caminhos e pequenas correções de execução).

### 2026-07-10 — `b08dc00` — Melhorias na medição
Adição do novo pipeline `etl-medicao` (`load_medicao.py` com detecção dinâmica de cabeçalho por palavras-chave, mapeamento de colunas, tratamento de tipos numéricos, exportação em CSV e backup em `.zip`; `run_medicao.py` como entrypoint). Removido o arquivo de exemplo `modelo-final/tabela_medicao.xlsx`, que não fazia mais parte do fluxo.

### 2026-07-14 — (não commitado nesta sessão) — Rebranding do projeto
Nome do projeto alterado de `ws-bm` para `etl-extracao-xlsx` em `pyproject.toml` e `main.py`; `README.md` reescrito com a documentação completa da proposta e das cinco pipelines; repositório no GitHub renomeado de `IsaiasMiranda/etl-python-xlsx` para `IsaiasMiranda/etl-extracao-xlsx` e remote `origin` local atualizado.
