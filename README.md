# etl-extracao-xlsx

Conjunto de pipelines ETL em Python para extrair, limpar, padronizar e consolidar planilhas Excel (`.xlsx`) provenientes de diferentes processos operacionais (medição, boletins, disponibilidade, SIGA), gerando arquivos consolidados (CSV/XLSX) prontos para carga em banco de dados ou uso analítico.

## Proposta do projeto

As áreas de operação geram, todos os meses, dezenas (às vezes centenas) de planilhas Excel espalhadas em pastas de rede — cada uma exportada por uma pessoa ou sistema diferente, com pequenas variações de layout: nomes de coluna com e sem acento, cabeçalhos em linhas diferentes, colunas duplicadas, linhas de "total" no rodapé, abas ocultas, etc.

Consolidar isso manualmente é lento e sujeito a erro. O objetivo deste projeto é automatizar essa etapa de **extração e limpeza (ETL)**:

1. **Extrair (Extract)** — varrer uma pasta de origem em busca de planilhas `.xlsx`/`.xls`, abrindo cada arquivo e identificando dinamicamente onde está o cabeçalho real da tabela (ignorando logotipos, textos soltos e abas ocultas).
2. **Transformar (Transform)** — normalizar nomes de colunas (minúsculas, sem acento, sem caracteres especiais), aplicar um mapeamento "de-para" para unificar variações do mesmo campo (ex.: `"Nº Nota Fiscal"`, `"Numero Nota Fiscal"` e `"NF"` viram todos `nota_fiscal`), remover linhas de total/vazias, tratar tipos numéricos e, em alguns casos, "explodir" (unpivot) colunas com múltiplos valores em linhas separadas.
3. **Carregar (Load)** — concatenar todas as planilhas processadas em um único arquivo de saída (CSV ou XLSX) com nome carimbado por data/hora, e mover ou copiar os arquivos de origem já processados para uma pasta de backup/histórico (em alguns casos compactando em `.zip`), evitando reprocessamento.

Cada processo de negócio tem seu próprio layout de planilha e suas próprias regras de limpeza, por isso o projeto é organizado como **vários pipelines independentes**, um por domínio de dados, em vez de um único script genérico.

## Estrutura do repositório

```
etl-extracao-xlsx/
├── etl-boletim-medicao/          # Boletins de medição (múltiplas abas por arquivo)
│   ├── load_boletim_medicao.py   # Regras de extração/limpeza/consolidação
│   └── run_boletim_medicao.py    # Ponto de entrada (define pastas de origem/destino)
├── etl-disponibilidade_medicao/  # Medição de disponibilidade (planilha "larga", 1 coluna por dia)
│   ├── load_disponibilidade_medicao.py
│   └── run_disponibilidade_medicao.py
├── etl-explode-bm-nfse/          # Explode uma coluna com múltiplos valores (ex.: vários boletins numa célula) em várias linhas
│   ├── load_explode_bm_nfse.py
│   └── run_explode_bm_nfse.py
├── etl-medicao/                  # Consolidação geral de medição, com backup em .zip
│   ├── load_medicao.py
│   └── run_medicao.py
├── etl-siga-medicao/             # Dados do sistema SIGA, segmentados por mês/ano na saída
│   ├── load_siga_medicao.py
│   └── run_siga_medicao.py
├── main.py                       # Ponto de entrada mínimo do projeto (placeholder)
├── pyproject.toml
└── uv.lock
```

Cada pasta `etl-*` segue o mesmo padrão:

- **`load_*.py`** — biblioteca com toda a lógica de negócio: localização de cabeçalho, mapeamento de colunas, limpeza e função principal `processar_boletins(...)` (ou equivalente) que orquestra o pipeline completo.
- **`run_*.py`** — script executável que define os caminhos de pasta de origem, destino e backup/processados, e chama a função principal do respectivo `load_*.py`.

## Pipelines disponíveis

| Pipeline | O que faz | Saída |
|---|---|---|
| `etl-boletim-medicao` | Lê boletins de medição (várias abas por arquivo), localiza o cabeçalho pela presença de "distribuidora"/"regional", aplica um dicionário de mapeamento extenso para padronizar nomes de coluna e consolida tudo num único arquivo. Move os arquivos de origem processados para uma pasta de histórico (renomeando com prefixo aleatório em caso de nome duplicado). | `.xlsx` consolidado com timestamp |
| `etl-disponibilidade_medicao` | Lê planilhas de disponibilidade em layout "largo" (uma coluna por dia do mês + totais), identifica a categoria do arquivo pelo nome (SEED MONEY, PERDAS, SMC) e faz *unpivot* das colunas de dia em linhas (`Data`, `Qtd_Horas`). Copia (não move) os arquivos processados. | `.csv` consolidado com timestamp |
| `etl-explode-bm-nfse` | Recebe planilhas com uma coluna (ex.: `boletim`) contendo múltiplos valores separados por vírgula/ponto e vírgula/barra/quebra de linha, e "explode" cada valor em uma linha própria. | `.csv` por arquivo de origem (`*_explodido.csv`) |
| `etl-medicao` | Pipeline mais genérico: localiza o cabeçalho por palavras-chave (nota, ordem, data, equipe, valor, status...), remove linha de total, padroniza colunas numéricas e nomes de coluna via mapeamento, e faz backup dos arquivos de origem compactando-os em `.zip` antes de excluí-los da origem. | `.csv` consolidado com timestamp + `.zip` de backup |
| `etl-siga-medicao` | Lê planilhas exportadas do sistema SIGA, localiza cabeçalho pela presença de "ordem de serviço"/"status da atividade", normaliza colunas e **segmenta a saída por mês/ano** (`Mes_Ano`, extraído da coluna `data`). Também compacta os arquivos de origem em `.zip` antes de excluí-los. | um `.csv` por mês/ano + `.zip` de backup |

Todos os pipelines:

- Ignoram arquivos temporários do Excel (`~$*.xlsx`);
- Ignoram abas ocultas (`sheet_state != 'visible'`);
- Evitam reprocessar arquivos consolidados já existentes na pasta de destino;
- Registram um resumo da execução (arquivos lidos, linhas extraídas x salvas, avisos/erros) via `print`/`logging`.

## Requisitos

- Python 3.13+ (ver [.python-version](.python-version))
- [uv](https://docs.astral.sh/uv/) para gestão de ambiente e dependências
- Dependências principais: `pandas`, `openpyxl` (ver [pyproject.toml](pyproject.toml))

## Instalação

```bash
uv sync
```

Isso cria o ambiente virtual em `.venv/` e instala as dependências travadas em `uv.lock`.

## Uso

Cada pipeline é independente e é executado a partir da sua própria pasta. Antes de rodar, ajuste os caminhos de origem/destino/backup definidos no início do respectivo `run_*.py` (atualmente fixos via `Path(r'...')`, apontando para pastas locais/rede específicas do ambiente onde o projeto roda).

Exemplo — consolidar boletins de medição:

```bash
uv run etl-boletim-medicao/run_boletim_medicao.py
```

Exemplo — consolidar medições do SIGA (gera um CSV por mês/ano):

```bash
uv run etl-siga-medicao/run_siga_medicao.py
```

O padrão se repete para os demais pipelines (`etl-disponibilidade_medicao`, `etl-explode-bm-nfse`, `etl-medicao`), sempre executando o `run_*.py` correspondente.

## Licença

Distribuído sob a licença MIT — veja [LICENSE](LICENSE).
