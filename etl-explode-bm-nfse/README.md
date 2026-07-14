# etl-explode-bm-nfse

Lê planilhas `.xlsx`/`.xls` que possuem uma coluna com múltiplos valores concatenados numa única célula (ex.: vários números de boletim separados por vírgula) e "explode" cada valor em uma linha própria, exportando o resultado em CSV.

## Arquivos

- [load_explode_bm_nfse.py](load_explode_bm_nfse.py) — lógica de explode e exportação.
- [run_explode_bm_nfse.py](run_explode_bm_nfse.py) — ponto de entrada; usa a própria pasta do script como origem e destino.

## Como executar

```bash
uv run etl-explode-bm-nfse/run_explode_bm_nfse.py
```

Diferente dos demais pipelines, este **não usa caminhos fixos de rede** — [run_explode_bm_nfse.py](run_explode_bm_nfse.py) resolve `Path(__file__).parent` em tempo de execução e usa essa mesma pasta como origem e destino, ou seja, basta colocar as planilhas `.xlsx`/`.xls` dentro de `etl-explode-bm-nfse/` antes de rodar.

## Funcionalidades

### `quebrar_valores_em_linhas(df, coluna) -> DataFrame`

Recebe um DataFrame e uma coluna-alvo, e expande essa coluna quebrando por delimitadores em linhas separadas:

1. Converte a coluna para string.
2. Divide cada valor por vírgula, ponto e vírgula, barra (`/` ou `|`) ou quebra de linha (regex `[;,/|\n]+`).
3. Usa `DataFrame.explode` para transformar cada item da lista resultante em uma linha própria (repetindo as demais colunas).
4. Remove espaços nas pontas de cada valor e descarta linhas cujo valor resultante seja vazio, `"nan"` ou `"none"`.

Se a coluna informada não existir no DataFrame, registra um aviso (`logger.warning`) e devolve o DataFrame original sem alterações.

### `processar_arquivos_explode(pasta_origem, pasta_destino, coluna_alvo='boletim') -> None`

Função principal:

1. **Descoberta**: busca arquivos `*.xlsx` e `*.xls` em `pasta_origem`, ignorando arquivos de trava do Excel (`~$*`).
2. **Leitura**: para cada arquivo, lê a primeira aba com `pd.read_excel` e normaliza os nomes de coluna (strip + minúsculas).
3. **Transformação**: aplica `quebrar_valores_em_linhas` na `coluna_alvo` (por padrão `boletim`).
4. **Exportação**: grava `pasta_destino/<nome_original>_explodido.csv` (codificação `utf-8`).
5. Registra, por arquivo, quantas linhas existiam antes e depois do explode. Erros por arquivo são capturados individualmente (`try/except` por iteração) e não interrompem o processamento dos demais arquivos.

## Formato de saída

Um arquivo `.csv` por planilha de origem, nomeado `<nome_original>_explodido.csv`, gravado na mesma pasta de origem (por padrão). Não há arquivamento/remoção dos arquivos de origem — origem e saída convivem na mesma pasta.
