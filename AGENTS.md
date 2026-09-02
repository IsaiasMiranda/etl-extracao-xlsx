# AGENTS.md

Este arquivo fornece orientações ao Codex (Codex.ai/code) para trabalhar com o código deste repositório.

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

### 2026-08-14 — `0103a40` — Mapeamento heurístico por palavras-chave
Fallback pra cabeçalho fora do `MAPEAMENTO` exato: qualquer combinação de
"folha"+"registro" mapeia pra `boletim`, "pep"+"ordem" mapeia pra
`origem_lancamento`. Só em `etl-boletim-medicao` (na época ainda um único
`load_boletim_medicao.py`, antes da separação produção/homologação de
2026-09-01 abaixo).

### 2026-08-31 — piloto isolado de normalização de boletim
Criados `normalizar_boletim.py`, `run_boletim_medicao_homologacao.py` e testes
unitários para realocação auditável por assinatura. O modo padrão de produção
permanece desativado. A execução dbt/Hop está condicionada à criação do banco
separado `elinsa_homologacao`; o usuário atual não possui permissão `CREATEDB`.

### 2026-09-01 — `16954aa` — Separação de módulos produção/homologação
`load_boletim_medicao.py` duplicado em `producao/` e `homologacao/` (antes
compartilhado — causou um conflito de merge real entre uma sessão mexendo em
melhorias de mapeamento de cabeçalho e outra integrando a normalização
auditável). Cada módulo evolui independente daqui pra frente:
- `producao/`: `load_boletim_medicao.py` (extração/consolidação simples,
  preserva colunas extras) + `run_boletim_medicao.py`.
- `homologacao/`: `load_boletim_medicao.py` (com os parâmetros de homologação
  — `normalizar_deslocamentos`, `remover_sem_boletim`,
  `excluir_boletim_nfse_duplicado`, `remover_boletim_sem_valor_duplicado`) +
  `normalizar_boletim.py` + `run_boletim_medicao_homologacao.py` +
  `test_normalizar_boletim.py`.
Uma melhoria feita num lado **não propaga automaticamente** para o outro —
replicar manualmente se for o caso (documentado no README do pipeline).
Pipeline de homologação revalidado ponta a ponta após a separação: 348.335
linhas, teste `assert_boletim_medicao_homologacao_padroes` PASS.

### 2026-09-01 — `12db8da` — Normalização de cabeçalho de produção trazida para homologação
`homologacao/load_boletim_medicao.py` adotou de produção (sem tocar
`producao/load_boletim_medicao.py`): `_normalizar_cabecalho()` (remove
acento/pontuação), `MAPEAMENTO` reorganizado + `MAPEAMENTO_HEURISTICO`, e
preservação de colunas extras não mapeadas (antes descartadas, achado real:
32 colunas eram perdidas). **Bug real encontrado e corrigido só na cópia de
homologação**: a chave `'folhregsrv': 'boletim'` existia no `MAPEAMENTO`
antigo e foi perdida na reorganização de produção — sem o fix, 6 arquivos
reais (2 deles 100% do arquivo) perdiam ~23 mil boletins. **Mesmo gap
continua em `producao/load_boletim_medicao.py`, não corrigido lá por pedido
explícito do usuário.** Mantido intacto o `_normalizar_periodo_medicao` rico
da homologação (reconhece ponto/barra/`mm/yyyy`/mês por extenso) — versão
simples de produção erraria/anularia 166.003 linhas (45% do corpus) por não
reconhecer separador de barra. Validado ponta a ponta: 348.335 linhas, 0
regressão, teste PASS.

### 2026-09-01 — `e7cdb73` — Recuperação de cp/iva entre arquivos + restrição de origem_lancamento ao formato PEP
`_recuperar_cp_iva_entre_arquivos`: quando o mesmo boletim aparece em várias
cópias/reenvios e ao menos uma traz `cp`/`iva` preenchido sem conflito com as
demais, propaga pras cópias nulas (recuperou 117.939 valores em escala, 0
falso-positivo medido). `origem_lancamento` restrito ao formato PEP/coletor
de custo único (`xx-xxxxxxxxxxx.x.xxxx.x`) — **removido deliberadamente** o
2º formato antes aceito (número SAP puro de 7-10 dígitos, 46% dos 286.657
valores preenchidos no corpus histórico), a pedido do usuário. Esses valores
passam a virar `NULL` em vez de serem aceitos.

### 2026-09-02 — Correção dos 2 testes desatualizados de `origem_lancamento`
Sessão de auditoria produção×homologação no repo `elinsa` (dbt) encontrou a
suíte pytest de `homologacao/test_normalizar_boletim.py` vermelha (2 de 24
testes falhando) — `test_origem_lancamento_aceita_pep_e_numero_sap_remove_
sufixo_excel` e `test_origem_lancamento_nao_e_roubado_para_boletim` ainda
esperavam que um número SAP puro sobrevivesse em `origem_lancamento`,
comportamento anterior à restrição de `e7cdb73` (acima) — os testes nunca
tinham sido atualizados junto com a mudança de regra de negócio. Corrigidos
(1º teste renomeado pra `test_origem_lancamento_aceita_so_pep_numero_sap_
puro_vira_nulo`, ambos agora esperam `pd.isna(...)` pro caso SAP puro).
`uv run --with pytest pytest etl-boletim-medicao/homologacao -q` → 24/24
PASS. Sem mudança de regra de negócio, só de expectativa de teste.

### 2026-09-02 (continuação) — Promoção de `homologacao/` para `producao/`
Usuário confirmou "homologação testado e dentro das normas estabelecidas" e
pediu o caminho para produção — sessão inteira de auditoria (repo `elinsa`)
comparou `elinsa_produ` × `elinsa_homologacao` boletim a boletim antes de
decidir promover (ver CLAUDE.md do repo `elinsa`, seção 8, para os números).
Plano mostrado e aprovado via plan mode antes de executar.

- **`producao/load_boletim_medicao.py`**: substituído por completo pela
  versão de `homologacao/` (schema de saída já é idêntico desde `12db8da`,
  2026-09-01 — só a lógica de limpeza mudou). Corrige `folhregsrv` ausente
  do `MAPEAMENTO` (gap que existia em `producao/` desde a reorganização de
  2026-09-01, documentado como "não corrigido lá por pedido do usuário" —
  agora corrigido, deixa de ser um gap exclusivo de homologação). 2
  comentários de cabeçalho que comparavam esta versão contra uma "produção"
  mais simples (agora inexistente, já que este arquivo passou a SER
  produção) foram reescritos para não confundir leitura futura — nenhuma
  mudança de lógica.
- **`producao/normalizar_boletim.py`** (novo): cópia de
  `homologacao/normalizar_boletim.py`, sem alteração — mesma decisão de
  duplicação deliberada já tomada pra `load_boletim_medicao.py` em
  `16954aa` (cada módulo evolui independente, sem risco de conflito de
  merge entre uma mudança futura em homologação e o código já promovido).
- **`producao/run_boletim_medicao.py`**: `processar_boletins(...)` ganhou
  as 5 flags de limpeza que `homologacao/run_boletim_medicao_homologacao.py`
  já usava (`normalizar_deslocamentos`, `caminho_auditoria`,
  `remover_sem_boletim`, `excluir_boletim_nfse_duplicado`,
  `remover_boletim_sem_valor_duplicado`), apontando pras pastas REAIS de
  produção (`D:\base-geral\base-boletim-medicao\...`, sem renomear nada) —
  `homologacao/` continua intocado e continua existindo como ambiente
  separado para validar mudanças futuras antes de promover de novo.
- Validado: `uv run --with pytest pytest etl-boletim-medicao -q` → 24/24
  PASS (lógica de `normalizar_boletim.py` não mudou, só a integração em
  `producao/`); smoke test confirmou que `producao/load_boletim_medicao.py`
  importa `normalizar_boletim` corretamente da própria pasta (mesmo
  mecanismo de resolução de import já usado em `homologacao/`).
- **Pendência explícita**: este commit só troca o CÓDIGO — não roda o
  pipeline contra os arquivos reais de produção nem toca em nenhum banco.
  O corte real de dados (rodar `producao/run_boletim_medicao.py` de
  verdade + Hop + `dbt --full-refresh` em `elinsa_produ`) é uma decisão
  separada, documentada como Fase 2 no plano da sessão, condicionada a um
  relatório de impacto (Fase 1, feito no repo `elinsa`) — não faz parte
  deste commit.
