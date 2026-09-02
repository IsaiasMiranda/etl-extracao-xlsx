"""Normalização auditável de valores deslocados no consolidado de boletins."""

from __future__ import annotations

import re
import unicodedata
from typing import Callable

import pandas as pd


COLUNAS_NORMALIZAVEIS = (
    "distribuidora", "regional", "tipo_medicao", "estrutura", "medicao",
    "periodo_medicao", "parceiro", "municipio", "equipe", "desc_nota_fiscal",
    "boletim", "valor_bm", "data_envio_bm", "origem_lancamento", "cp", "iva",
    "domicilio_fiscal", "codigo_tarifa_fiscal", "identificador_medicao",
    "identificador_agrupamento", "contrato", "processo", "texto_boletim",
)


def _texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return re.sub(r"\s+", " ", str(valor).strip())


def _remover_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


# Sentinelas fabricados na própria planilha de origem (não pelo pipeline):
# "NÃO INFORMADO" (com/sem parênteses) e "(EM BRANCO)" em variações de
# maiúsc./minúsc./espaçamento. Achado real na homologação em escala de
# 2026-08-31: 14.297 linhas em "estrutura", 984+728+11 em "regional", 11 em
# "tipo_medicao"/"municipio"/"desc_nota_fiscal". Doutrina do projeto: Silver
# não fabrica valor de negócio, nulo real permanece nulo -- mesma regra
# aplicada aqui a um placeholder que já chega pronto da origem, não só a
# COALESCE que o próprio pipeline poderia introduzir.
def _e_sentinela(texto: str) -> bool:
    compacto = re.sub(r"[\s()]+", "", _remover_acentos(texto).upper())
    return compacto in {"NAOINFORMADO", "EMBRANCO"}


def _normalizar_texto(valor, remover_acentos: bool = False):
    texto = _texto(valor)
    if not texto or _e_sentinela(texto):
        return pd.NA
    if remover_acentos:
        texto = _remover_acentos(texto)
    return texto


def _normalizar_boletim(valor):
    texto = _texto(valor)
    return pd.NA if not texto else re.sub(r"[.,]0$", "", texto.upper())


def _normalizar_origem_lancamento(valor):
    # Mesma limpeza estrutural de sufixo ".0"/",0" já aplicada a boletim/
    # contrato (artefato de número inteiro lido como float pelo Excel) --
    # achado real na homologação de 2026-09-01: 5.881 linhas com número
    # SAP puro tipo "8000012470.0" em vez de "8000012470", nunca corrigidas
    # porque origem_lancamento caía no `_normalizar_texto` genérico. Não
    # força maiúsculas (diferente de `_normalizar_boletim`) -- os códigos
    # PEP já chegam consistentemente em maiúsculas na origem, sem achado
    # real de minúscula pra justificar a mudança de comportamento.
    texto = _texto(valor)
    if not texto or _e_sentinela(texto):
        return pd.NA
    return re.sub(r"[.,]0$", "", texto)


# Origem de lançamento RESTRITA a um único formato (2026-09-01, a pedido do
# usuário, com impacto medido e aprovado antes de aplicar): só o código
# PEP/coletor de custo, no formato EXATO "xx-xxxxxxxxxxx.x.xxxx.x"
# (2-11.1.4.1 caracteres, ex.: "PA-2652502SMC1.1.0125.D",
# "AP-2402304EME1.F.0056.D" -- letras/dígitos genéricos em cada segmento,
# não só PA/AP, pra não travar se aparecer outra sigla de estado). Tamanho
# de segmento fixo confirmado contra os 154.959 valores PEP reais (0
# exceção) -- não usar `+`/tamanho livre aqui, ficaria mais permissivo do
# que o formato de negócio de verdade.
#
# ANTES (até 2026-09-01) também aceitava número SAP puro de 7 a 10 dígitos
# ("6258513"/"6001429759") como 2º formato válido -- 131.698 dos 286.657
# valores preenchidos (46%) eram esse formato. Removido deliberadamente:
# o negócio só reconhece o código PEP como "origem de lançamento" válida;
# os valores SAP puro passam a virar NULL (doutrina do projeto: não
# fabricar/manter valor fora do contrato aprovado da coluna). Reflexo
# obrigatório em `tests/assert_boletim_medicao_homologacao_padroes.sql`
# (repo elinsa) -- mesmo regex, mantido em sincronia manual.
_REGEX_ORIGEM_LANCAMENTO_VALIDO = re.compile(
    r"[A-Z0-9]{2}-[A-Z0-9]{11}\.[A-Z0-9]\.[0-9]{4}\.[A-Z0-9]"
)


def _validar_origem_lancamento_final(valor):
    # Validação PRÓPRIA de "origem_lancamento" (não via ASSINATURAS --
    # mesmo motivo de "contrato"/"periodo_medicao": o formato "7 a 10
    # dígitos puros" é genérico demais e colidiria com boletim/valor_bm
    # se entrasse no motor de busca-e-realocação). Chamada só no FINAL de
    # `normalizar_dataframe`, depois que origem_lancamento já teve chance
    # de ceder valor deslocado pra outra coluna (segue elegível como
    # ORIGEM de candidatos, só não é destino de busca).
    if pd.isna(valor) or re.fullmatch(_REGEX_ORIGEM_LANCAMENTO_VALIDO, valor):
        return valor
    return pd.NA


def _validar_contrato_final(valor):
    # Validação PRÓPRIA de "contrato" (não via ASSINATURAS -- ver comentário
    # do dict mais abaixo): o que sobrar tem que ser um número SAP puro,
    # senão vira nulo. Chamada só no FINAL de `normalizar_dataframe`, depois
    # que "contrato" já teve chance de ceder um valor deslocado (ex.: um
    # iva/cp que acabou lá por engano) pra outra coluna -- validar antes
    # disso apagaria o valor antes da busca por candidatos rodar. Achado
    # real na homologação em escala de 2026-08-31: 11 linhas com o
    # placeholder "( EM BRANCO )" da própria planilha de origem.
    if pd.isna(valor) or re.fullmatch(r"[0-9]+", valor):
        return valor
    return pd.NA


def _normalizar_data(valor):
    texto = _texto(valor)
    if not texto:
        return pd.NA
    texto = re.sub(r"\s+.*$", "", texto)
    match = re.fullmatch(r"([0-9]{4})/([0-9]{2})/([0-9]{2})", texto)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    match = re.fullmatch(r"([0-9]{2})/([0-9]{2})/([0-9]{4})", texto)
    if match:
        return f"{match.group(1)}.{match.group(2)}.{match.group(3)}"
    return texto


def _finalizar_data_envio_bm(valor):
    # Conversão final de data_envio_bm pro padrão dd/mm/yyyy (2026-09-01:
    # pedido inicial era formato americano MM/DD/YYYY, revertido no mesmo
    # dia pra manter dd/mm/yyyy). Roda UMA ÚNICA VEZ aqui, fora de
    # `_normalizar_data` (chamada 2x internamente, via
    # `_normalizar_dataframe_base`, linhas 403/410) -- qualquer formato
    # "2-2-4 com separador" tem a mesma forma de dd/mm/yyyy; se a
    # reordenação fosse feita direto ali, a 2ª chamada reinterpretaria a
    # saída da 1ª como entrada de novo (bug de idempotência). Os 2
    # formatos intermediários que `_normalizar_data` produz hoje
    # (`YYYY-MM-DD` traço, `DD.MM.YYYY` ponto) foram escolhidos de
    # propósito por não colidirem em forma com nenhuma entrada aceita --
    # aqui, no passo final isolado (roda 1x só), os dois convertem pra
    # dd/mm/yyyy sem ambiguidade (o formato ponto já está na ordem
    # dia-mês-ano, só troca o separador; o formato traço, ano-mês-dia,
    # precisa reordenar).
    if pd.isna(valor):
        return valor
    texto = str(valor)
    match = re.fullmatch(r"([0-9]{4})-([0-9]{2})-([0-9]{2})", texto)
    if match:
        ano, mes, dia = match.groups()
    else:
        match = re.fullmatch(r"([0-9]{2})\.([0-9]{2})\.([0-9]{4})", texto)
        if match:
            dia, mes, ano = match.groups()
        else:
            return valor
    if (ano, mes, dia) == ("1900", "01", "01"):
        # Sentinela de data vazia do Excel (dia zero da serial numérica de
        # data) -- não é uma data de negócio real. Achado real na
        # homologação de 2026-09-01: 42.110 linhas (23% do preenchimento
        # de data_envio_bm), concentradas em 4 arquivos da mesma família
        # ("boletim_medicao.xlsx"/"Boletim Medição 8/9/10.xlsx"). Mesma
        # doutrina já aplicada aos sentinelas de texto "NÃO INFORMADO"/
        # "(EM BRANCO)" -- vira nulo real, não data fabricada.
        return pd.NA
    return f"{dia}/{mes}/{ano}"


def _normalizar_dataframe_base(df: pd.DataFrame) -> pd.DataFrame:
    saida = df.copy()
    for coluna in COLUNAS_NORMALIZAVEIS:
        if coluna not in saida.columns:
            saida[coluna] = pd.NA
        if coluna in ("boletim", "contrato"):
            # A validação de que só sobrou número SAP puro em "contrato"
            # acontece só no FINAL de `normalizar_dataframe` (função
            # `_validar_contrato_final`), depois que ele já teve chance de
            # ceder um valor deslocado (ex.: iva/cp) pra outra coluna --
            # validar aqui, cedo demais, apagaria esse valor antes da busca
            # por candidatos rodar (quebrava recuperação legítima).
            saida[coluna] = saida[coluna].map(_normalizar_boletim)
        elif coluna == "data_envio_bm":
            saida[coluna] = saida[coluna].map(_normalizar_data)
        elif coluna == "origem_lancamento":
            saida[coluna] = saida[coluna].map(_normalizar_origem_lancamento)
        elif coluna == "municipio":
            saida[coluna] = saida[coluna].map(
                lambda valor: _normalizar_texto(valor, remover_acentos=True)
            )
        else:
            saida[coluna] = saida[coluna].map(_normalizar_texto)
    return saida


def _regex(pattern: str, valor: str) -> bool:
    return bool(re.fullmatch(pattern, valor))


def _match_boletim(valor: str) -> bool:
    return _regex(r"[0-9]+(?:[.,]0)?", valor)


# Achado real na homologação em escala de 2026-08-31: a assinatura antiga
# aceitava QUALQUER dígito solto sem decimal ([0-9]+), então um "boletim"
# de 10 dígitos (ex. "1011587571") passava por "valor_bm" válido -- em
# 2.086 linhas onde valor_bm precisava de correção, o motor "recuperava"
# (errado) o número do boletim pra lá, apagando-o (e derrubando a linha
# depois, via `remover_sem_boletim`). Fix: dígito solto sem decimal só é
# aceito até 6 dígitos (maior valor real observado no corpus histórico
# sem separador decimal); COM decimal (vírgula ou ponto), qualquer
# tamanho de parte inteira continua aceito -- boletim nunca tem decimal,
# então essa forma nunca colide com ele (ex.: "1007155,63" continua ok).
_REGEX_VALOR_BM = (
    r"[+-]?(?:[0-9]{1,3}(?:\.[0-9]{3})+(?:,[0-9]+|\.[0-9]+)?"
    r"|[0-9]+(?:,[0-9]+|\.[0-9]+)"
    r"|[0-9]{1,6})"
)


def _match_valor_bm(valor: str) -> bool:
    return _regex(_REGEX_VALOR_BM, valor)


def _match_data(valor: str) -> bool:
    return _regex(
        r"(?:[0-9]{2}\.[0-9]{2}\.[0-9]{4}|[0-9]{2}/[0-9]{2}/[0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
        valor,
    )


def _match_cp(valor: str) -> bool:
    return _regex(r"[A-Z]{2}[0-9]{2}", valor)


def _match_iva(valor: str) -> bool:
    return _regex(r"[A-Z][0-9]", valor)


_REGEX_PERIODO_MEDICAO_VALIDO = re.compile(
    r"[0-9]{2}\.[0-9]{2}\.[0-9]{4} a [0-9]{2}\.[0-9]{2}\.[0-9]{4}"
)


def _validar_periodo_medicao_final(valor):
    # Validação PRÓPRIA de "periodo_medicao" (não via ASSINATURAS -- ver
    # comentário do dict mais abaixo, mesmo padrão de `_validar_contrato_
    # final`). Formato exigido: "dd.mm.yyyy a dd.mm.yyyy", o mesmo que
    # `load_boletim_medicao._normalizar_periodo_medicao` já garante quando
    # reconhece o valor. Achado real na homologação em escala de 2026-08-31:
    # sem essa validação, um "1900-01-01" (sentinela de data vazia do
    # Excel) ou uma data avulsa vindo de `data_envio_bm` sobrevivia intacto.
    if pd.isna(valor) or re.fullmatch(_REGEX_PERIODO_MEDICAO_VALIDO, valor):
        return valor
    return pd.NA


def _match_processo_origem(valor: str) -> bool:
    return bool(re.fullmatch(r"P[AB]-[0-9]{7}[A-Z0-9]+(?:\.[A-Z0-9]+)+", valor))


def _match_domicilio(valor: str) -> bool:
    return bool(re.fullmatch(r"P[AB] [0-9]{7}", valor))


def _match_tarifa(valor: str) -> bool:
    return bool(re.fullmatch(r"LC[0-9]{3}_[0-9]{2}\.[0-9]{2}", valor))


def _match_identificador_medicao(valor: str) -> bool:
    return _regex(r"[0-9]+(?:\.[0-9]+){4,}", valor)


def _match_identificador_agrupamento(valor: str) -> bool:
    return _regex(r"[0-9]+-[0-9]+-[0-9]{8}-[0-9]+", valor)


# Além da forma composta "texto - texto", "processo" também aparece como um
# código curto isolado (sem hífen separador) em boletins de determinados
# leiautes. Lista fechada observada na homologação em escala de 2026-08-31
# (53.728 linhas rejeitadas indevidamente, só esses 8 valores distintos —
# sem cauda longa). Revisar esta lista se um novo código aparecer rejeitado
# no CSV de auditoria (tipo=INVALIDO_NAO_NORMALIZADO, coluna_destino=processo).
_PROCESSO_CODIGOS_CONHECIDOS = frozenset({
    "CUSTEIO", "GEOM", "SMC", "SEED MONEY", "PLPT", "GSTC",
    "NORMALIZAÇÃO BT", "SUBSTITUIÇÃO DE MD",
})


def _match_processo(valor: str) -> bool:
    # Processo é uma descrição textual composta por duas partes separadas
    # por hífen (estrutura livre, só a forma importa) OU um dos códigos
    # curtos conhecidos acima (comparação sem distinguir maiúsc./minúsc.).
    if re.fullmatch(r"[^-]+\s+-\s+.+", valor):
        return True
    return valor.upper() in _PROCESSO_CODIGOS_CONHECIDOS


def _match_texto_boletim(valor: str) -> bool:
    return bool(re.fullmatch(r"[^>]+(?:>[^>]*){2,}", valor))


# Regras estruturadas do piloto. Texto livre não é usado como assinatura.
#
# "contrato" foi testado como destino (achado real: sua assinatura antiga
# nunca batia e anulava 100% do campo) e removido de novo na mesma
# homologação de 2026-08-31: número SAP puro ([0-9]+) é indistinguível, por
# formato, de "boletim"/"origem_lancamento" -- em qualquer arquivo cujo
# leiaute simplesmente não tinha coluna Contrato, o motor "recuperava"
# (errado) o número do boletim pra lá, 117 mil vezes. "contrato" recebe só
# limpeza estrutural (".0"/",0", em `_normalizar_dataframe_base`) e uma
# validação PRÓPRIA isolada no final de `normalizar_dataframe`
# (`_validar_contrato_final`) -- nunca busca nem é candidato de
# realocação vindo/indo de `boletim`/`valor_bm` (ver `candidatos` abaixo).
#
# "periodo_medicao" também foi removido daqui na mesma homologação: sua
# assinatura antiga usava `str.contains` frouxo demais (qualquer trecho
# "dígito-separador-dígito-separador-dígito"), então `data_envio_bm`/
# `origem_lancamento`/`identificador_agrupamento` sempre pareciam
# candidatos válidos -- 100% das 26.667 "correções" produziam um valor
# fora do padrão "dd.mm.yyyy a dd.mm.yyyy" (nenhuma dessas 3 colunas
# contém genuinamente um período de medição; um "1900-01-01" -- sentinela
# de data vazia do Excel -- chegou a vazar assim pra "periodo_medicao").
# Mesmo tratamento: só validação PRÓPRIA isolada no final
# (`_validar_periodo_medicao_final`), nunca busca-e-realocação.
ASSINATURAS: dict[str, Callable[[str], bool]] = {
    "boletim": _match_boletim,
    "valor_bm": _match_valor_bm,
    "data_envio_bm": _match_data,
    "cp": _match_cp,
    "iva": _match_iva,
    "domicilio_fiscal": _match_domicilio,
    "codigo_tarifa_fiscal": _match_tarifa,
    "identificador_medicao": _match_identificador_medicao,
    "identificador_agrupamento": _match_identificador_agrupamento,
    "processo": _match_processo,
    "texto_boletim": _match_texto_boletim,
}


ASSINATURAS_REGEX = {
    # Deliberadamente SÓ os 2 tamanhos reais observados no corpus histórico
    # (7 e 10 dígitos) -- não o `[0-9]+` irrestrito de `_match_boletim`.
    # Este regex também decide quem é candidato a relocação PARA "boletim"
    # (via `_serie_match`/`candidatos`), e dígitos soltos de outras colunas
    # (valor_bm sem decimal, "20"/"40" de cp malformado) têm tamanho
    # pequeno/genérico que nunca é um boletim de verdade. Achado real na
    # homologação em escala de 2026-08-31: sem essa restrição, 312 valores
    # de valor_bm e 219 de cp (100% dos "boletins de 2 dígitos" -- só
    # "20"/"40") eram roubados pra boletim. `_match_boletim` (usado só na
    # validação final do PRÓPRIO valor, `_is_invalid`) continua frouxo de
    # propósito -- não queremos anular um boletim genuíno só por ter um
    # tamanho fora do esperado, só impedir que OUTRA coluna seja
    # erroneamente tratada como candidata a preenchê-lo.
    "boletim": r"(?:[0-9]{7}|[0-9]{10})(?:[.,]0)?",
    "valor_bm": _REGEX_VALOR_BM,
    "data_envio_bm": r"(?:[0-9]{2}\.[0-9]{2}\.[0-9]{4}|[0-9]{2}/[0-9]{2}/[0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
    "cp": r"[A-Z]{2}[0-9]{2}",
    "iva": r"[A-Z][0-9]",
    "domicilio_fiscal": r"[A-Z]{2} [0-9]{7}",
    "codigo_tarifa_fiscal": r"LC[0-9]{3}_[0-9]{2}\.[0-9]{2}",
    "identificador_medicao": r"[0-9]+(?:\.[0-9]+){4,}",
    "identificador_agrupamento": r"[0-9]+-[0-9]+-[0-9]{8}-[0-9]+",
    "contrato": r"[0-9]+",
    # Deliberadamente SÓ o formato "texto - texto" aqui (não o whitelist de
    # `_match_processo`) -- este regex também decide quem é candidato a
    # relocação para "processo" (via `_serie_match`/`candidatos`), e
    # "CUSTEIO" já é o valor legítimo de `tipo_medicao` em ~99% das linhas.
    # Incluir o whitelist aqui roubava tipo_medicao/parceiro pra dentro de
    # processo (achado real, 2026-08-31: 12.838 linhas de tipo_medicao
    # indevidamente movidas). O whitelist só vale pra validar o valor que
    # já está no PRÓPRIO processo (`_match_processo`, usado só no check
    # final `_is_invalid`), nunca pra puxar de outra coluna.
    "processo": r"[^-]+\s+-\s+.+",
    "texto_boletim": r"[^>]+(?:>[^>]*){2,}",
}


def _serie_match(serie: pd.Series, destino: str) -> pd.Series:
    # "periodo_medicao" não tem mais caso especial aqui: saiu de ASSINATURAS
    # (ver comentário do dict acima), então este destino nunca é chamado com
    # ele -- e como candidato ORIGEM pra outro destino, seu formato "dd.mm.
    # yyyy a dd.mm.yyyy" já não bate por acaso com nenhuma assinatura real.
    valores = serie.fillna('').astype('string').str.strip()
    return valores.str.fullmatch(ASSINATURAS_REGEX[destino], na=False)


def _corrigir_layout_rotacionado(saida: pd.DataFrame, auditoria: list[dict[str, object]]) -> None:
    """Corrige o layout que desloca/rotaciona as 24 colunas de negócio.

    A assinatura usa o arquivo no primeiro campo e a cadeia inequívoca
    processo-origem -> CP -> IVA -> domicílio -> tarifa. Sem essa cadeia, a
    linha não é alterada automaticamente.
    """
    arquivo = saida['distribuidora'].fillna('').astype('string').str.strip()
    cp_atual = saida['cp'].fillna('').astype('string').str.strip()
    iva_atual = saida['iva'].fillna('').astype('string').str.strip()
    domicilio_atual = saida['domicilio_fiscal'].fillna('').astype('string').str.strip()
    tarifa_atual = saida['codigo_tarifa_fiscal'].fillna('').astype('string').str.strip()
    mascara = (
        arquivo.str.lower().str.endswith('.xlsx')
        & cp_atual.map(_match_processo_origem)
        & iva_atual.map(_match_cp)
        & domicilio_atual.map(_match_iva)
        & tarifa_atual.map(_match_domicilio)
        & saida['identificador_medicao'].fillna('').astype('string').str.strip().map(_match_tarifa)
    )
    if not mascara.any():
        return

    colunas = list(COLUNAS_NORMALIZAVEIS) + ['arquivo_origem']
    if 'arquivo_origem' not in saida.columns:
        saida['arquivo_origem'] = pd.NA
    antigos = saida.loc[mascara, colunas].copy()
    rotacionados = antigos.iloc[:, 1:].copy()
    rotacionados[antigos.columns[0]] = antigos.iloc[:, 0]
    rotacionados.columns = colunas
    saida.loc[mascara, colunas] = rotacionados.to_numpy()
    for indice in antigos.index:
        auditoria.append({
            'linha': indice,
            'tipo': 'LAYOUT_ROTACIONADO_CORRIGIDO',
            'coluna_origem': 'distribuidora',
            'coluna_destino': 'arquivo_origem',
            'valor_anterior_destino': antigos.at[indice, 'arquivo_origem'],
            'valor_normalizado': antigos.at[indice, 'distribuidora'],
        })


def _corrigir_nota_fiscal_deslocada(saida: pd.DataFrame, auditoria: list[dict[str, object]]) -> None:
    """Remove a nota fiscal que entrou indevidamente antes de ``cp``.

    Alguns leiautes antigos trazem a NF, que não pertence ao contrato da
    Bronze, apesar de o cabeçalho não a declarar. Neles, o valor de ``cp`` é
    uma NF de dez dígitos e a cadeia seguinte é inequívoca: CP -> IVA ->
    domicílio -> tarifa -> identificador. Só essa cadeia permite mover os
    campos; linhas sem confirmação ficam intactas para análise manual.
    """
    cp = saida['cp'].fillna('').astype('string').str.strip()
    iva = saida['iva'].fillna('').astype('string').str.strip()
    domicilio = saida['domicilio_fiscal'].fillna('').astype('string').str.strip()
    codigo = saida['codigo_tarifa_fiscal'].fillna('').astype('string').str.strip()
    identificador = saida['identificador_medicao'].fillna('').astype('string').str.strip()
    mascara = (
        cp.str.fullmatch(r'[0-9]{10}', na=False)
        & iva.map(_match_cp)
        & domicilio.map(_match_iva)
        & codigo.map(_match_domicilio)
        & identificador.map(_match_tarifa)
    )
    if not mascara.any():
        return

    colunas = list(COLUNAS_NORMALIZAVEIS[14:]) + ['arquivo_origem']
    antigos = saida.loc[mascara, colunas].copy()
    deslocados = antigos.copy()
    for atual, proxima in zip(colunas, colunas[1:]):
        deslocados[atual] = antigos[proxima]
    deslocados[colunas[-1]] = pd.NA
    saida.loc[mascara, colunas] = deslocados.to_numpy()
    for indice in antigos.index:
        auditoria.append({
            'linha': indice,
            'tipo': 'NOTA_FISCAL_DESLOCADA_REMOVIDA',
            'coluna_origem': 'cp',
            'coluna_destino': 'cp',
            'valor_anterior_destino': antigos.at[indice, 'cp'],
            'valor_normalizado': saida.at[indice, 'cp'],
        })


def _is_invalid(valor, coluna: str) -> bool:
    if pd.isna(valor) or not _texto(valor):
        return False
    if coluna not in ASSINATURAS:
        return False
    return not ASSINATURAS[coluna](_texto(valor))


# "cp"/"iva" recuperáveis ENTRE ARQUIVOS -- achado real na homologação em
# escala de 2026-09-01: o mesmo boletim chega em várias cópias/reenvios do
# consolidado (mesmo lote reenviado por e-mail, planilha-resumo sem a
# coluna, etc.); em 6.339 boletins (cp) / 6.259 (iva) pelo menos uma cópia
# traz o valor preenchido enquanto outra cópia do MESMO boletim veio nula
# -- 56.182/71.646 (cp) e 58.309/76.552 (iva) das linhas nulas medidas têm
# essa saída. Medido também: quando um boletim tem >1 cópia preenchida,
# elas NUNCA divergem entre si (0 conflitos em 6.339+6.259 boletins) --
# mesmo assim a função trata conflito como caso legítimo (não é garantia
# estrutural, só o que o corpus mostrou até agora) e não escolhe valor "no
# chute": vira AMBIGUA_ENTRE_ARQUIVOS na auditoria, linha intocada.
#
# Diferente de `_corrigir_layout_rotacionado`/`_corrigir_nota_fiscal_
# deslocada` (deslocamento DENTRO da mesma linha) e do loop de ASSINATURAS
# (candidato de OUTRA coluna da mesma linha) -- esta é recuperação ENTRE
# LINHAS diferentes que compartilham o mesmo `boletim` já normalizado.
# Roda depois do loop de ASSINATURAS de propósito: só tenta a busca entre
# arquivos no que sobrar nulo depois de toda chance de correção dentro da
# própria linha já ter sido esgotada.
def _recuperar_cp_iva_entre_arquivos(
    saida: pd.DataFrame, auditoria: list[dict[str, object]]
) -> None:
    boletins = saida['boletim'].fillna('').astype('string').str.strip()
    tem_boletim = boletins != ''

    for coluna in ('cp', 'iva'):
        preenchido = tem_boletim & saida[coluna].notna()
        nulo = tem_boletim & saida[coluna].isna()
        if not preenchido.any() or not nulo.any():
            continue

        distintos_por_boletim = (
            saida.loc[preenchido, [coluna]]
            .assign(boletim=boletins[preenchido])
            .groupby('boletim')[coluna]
            .agg(lambda serie: tuple(sorted(serie.unique())))
        )
        valor_unico = distintos_por_boletim[
            distintos_por_boletim.map(len) == 1
        ].map(lambda valores: valores[0])
        boletins_ambiguos = set(
            distintos_por_boletim[distintos_por_boletim.map(len) > 1].index
        )

        valor_recuperado = boletins.map(valor_unico)
        recuperavel = nulo & valor_recuperado.notna()
        for indice in saida.index[recuperavel]:
            novo_valor = valor_recuperado.at[indice]
            saida.at[indice, coluna] = novo_valor
            auditoria.append({
                'linha': indice,
                'tipo': 'RECUPERADA_ENTRE_ARQUIVOS',
                'coluna_origem': coluna,
                'coluna_destino': coluna,
                'valor_anterior_destino': pd.NA,
                'valor_normalizado': novo_valor,
            })

        ambiguo = nulo & boletins.isin(boletins_ambiguos)
        for indice in saida.index[ambiguo]:
            candidatos = distintos_por_boletim.at[boletins.at[indice]]
            auditoria.append({
                'linha': indice,
                'tipo': 'AMBIGUA_ENTRE_ARQUIVOS',
                'coluna_destino': coluna,
                'candidatos': '; '.join(str(v) for v in candidatos),
            })


def normalizar_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Normaliza texto e realoca apenas candidatos unívocos por assinatura.

    A operação usa máscaras vetorizadas para permanecer viável no volume
    histórico da Bronze; não percorre célula a célula em Python.
    """
    saida = _normalizar_dataframe_base(df)
    auditoria: list[dict[str, object]] = []
    _corrigir_layout_rotacionado(saida, auditoria)
    _corrigir_nota_fiscal_deslocada(saida, auditoria)
    # As correções estruturais também movem valores entre colunas; reaplica
    # a normalização básica para que, por exemplo, ``101...,0`` não volte a
    # aparecer em ``boletim`` após uma rotação.
    saida = _normalizar_dataframe_base(saida)

    for destino in ASSINATURAS:
        destino_valores = saida[destino].fillna('').astype('string').str.strip()
        precisa_corrigir = ~_serie_match(saida[destino], destino)
        if not precisa_corrigir.any():
            continue

        candidatos = {
            origem: _serie_match(saida[origem], destino) & ~saida[origem].fillna('').astype('string').str.strip().eq('')
            for origem in COLUNAS_NORMALIZAVEIS
            # "contrato" só fica de fora do pool de candidatos quando o
            # destino também aceita dígitos soltos (boletim/valor_bm) --
            # aí sim ele é indistinguível por formato de um número SAP
            # puro. Achado real na homologação em escala de 2026-08-31:
            # nas duas direções (contrato->boletim e boletim->contrato) o
            # motor "recuperava" errado um dígito que já estava no lugar
            # certo. Continua elegível como origem pra destinos de formato
            # distinto (ex.: iva/cp), onde não há essa ambiguidade real.
            #
            # "origem_lancamento" fica de fora do pool só quando o destino
            # é "boletim": o código de coletor de custo/PEP que mora lá
            # também é um número SAP puro de 10 dígitos, mesmo tamanho do
            # boletim real (achado real, mesma homologação: 9.791 linhas
            # com um código tipo "6020018020"/"8000012474" -- prefixo bem
            # diferente do "101..." de boletim real -- indevidamente
            # copiado pra boletim, apagando o código original de lá).
            #
            # "parceiro" fica de fora do pool só quando o destino é
            # "processo": o nome da parceira ("ELINSA - ELETROTÉCNICA...")
            # é formatado como "texto - texto", mesma assinatura de
            # `processo` -- indistinguível por formato. Achado real na
            # homologação de 2026-09-01: em leiautes sem coluna Processo
            # (ex. "BOLETIM DE MEDIÇÃO"), o motor "recuperava" (errado) o
            # nome da parceira pra dentro de `processo`, afetando 86.054
            # linhas (22% do dataset) com um valor que nunca existiu na
            # origem pra esse campo.
            if origem != destino
            and not (origem == 'contrato' and destino in ('boletim', 'valor_bm'))
            and not (origem == 'origem_lancamento' and destino == 'boletim')
            and not (origem == 'parceiro' and destino == 'processo')
        }
        quantidade = sum(m.astype('int8') for m in candidatos.values())
        univoco = precisa_corrigir & quantidade.eq(1)
        ambiguo = precisa_corrigir & quantidade.gt(1)

        for origem, mascara in candidatos.items():
            mover = univoco & mascara
            if not mover.any():
                continue
            valores = saida.loc[mover, origem]
            if destino == 'boletim':
                valores = valores.map(_normalizar_boletim)
            elif destino == 'data_envio_bm':
                valores = valores.map(_normalizar_data)
            anteriores = saida.loc[mover, destino].copy()
            saida.loc[mover, destino] = valores
            saida.loc[mover, origem] = pd.NA
            for indice, valor in valores.items():
                auditoria.append({
                    'linha': indice,
                    'tipo': 'CORRIGIDA',
                    'coluna_origem': origem,
                    'coluna_destino': destino,
                    'valor_anterior_destino': anteriores.at[indice],
                    'valor_normalizado': valor,
                })

        for indice in saida.index[ambiguo]:
            encontrados = [
                f'{origem}={saida.at[indice, origem]}'
                for origem, mascara in candidatos.items()
                if mascara.loc[indice]
            ]
            auditoria.append({
                'linha': indice,
                'tipo': 'AMBIGUA',
                'coluna_destino': destino,
                'candidatos': '; '.join(encontrados),
            })

        # A camada normalizada não mantém um valor incompatível com o
        # contrato da coluna. O original continua preservado na tabela
        # *_original e o valor rejeitado fica na auditoria para tratamento.
        for indice in saida.index[precisa_corrigir]:
            valor = saida.at[indice, destino]
            if _is_invalid(valor, destino):
                saida.at[indice, destino] = pd.NA
                auditoria.append({
                    'linha': indice,
                    'tipo': (
                        'INVALIDO_AMBIGUO_NAO_NORMALIZADO'
                        if ambiguo.loc[indice]
                        else 'INVALIDO_NAO_NORMALIZADO'
                    ),
                    'coluna_destino': destino,
                    'valor_anterior_destino': valor,
                    'valor_normalizado': pd.NA,
                })

    # Recuperação ENTRE ARQUIVOS de cp/iva -- roda só depois que o loop de
    # ASSINATURAS acima já esgotou toda chance de correção DENTRO da mesma
    # linha (ver docstring de `_recuperar_cp_iva_entre_arquivos`).
    _recuperar_cp_iva_entre_arquivos(saida, auditoria)

    # "contrato", "periodo_medicao" e "origem_lancamento" não passam pelo
    # loop de ASSINATURAS como DESTINO (não buscam candidato de outra
    # coluna -- ver comentário do dict acima), então cada um tem sua
    # validação própria isolada aqui, só no final, depois que já tiveram
    # chance de ceder um valor deslocado pra outra coluna. Continuam
    # elegíveis como ORIGEM de candidatos pra outros destinos normalmente.
    for coluna, validador in (
        ('contrato', _validar_contrato_final),
        ('periodo_medicao', _validar_periodo_medicao_final),
        ('origem_lancamento', _validar_origem_lancamento_final),
    ):
        invalido = saida[coluna].map(
            lambda v, validador=validador: not pd.isna(v) and pd.isna(validador(v))
        )
        for indice in saida.index[invalido]:
            valor = saida.at[indice, coluna]
            saida.at[indice, coluna] = pd.NA
            auditoria.append({
                'linha': indice,
                'tipo': 'INVALIDO_NAO_NORMALIZADO',
                'coluna_destino': coluna,
                'valor_anterior_destino': valor,
                'valor_normalizado': pd.NA,
            })

    # "data_envio_bm": conversão final pro padrão dd/mm/yyyy -- pedido
    # inicial (2026-09-01) era formato americano, revertido no mesmo dia.
    # Ver docstring de `_finalizar_data_envio_bm` pro detalhe de por que
    # essa conversão roda isolada, 1x só, no fim do pipeline.
    saida['data_envio_bm'] = saida['data_envio_bm'].map(_finalizar_data_envio_bm)

    return saida, pd.DataFrame(auditoria)
