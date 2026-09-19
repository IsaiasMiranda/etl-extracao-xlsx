"""Relatório de conformidade de formato por coluna do consolidado de homologação.

Aplica as mesmas assinaturas (regex) de ``normalizar_boletim.ASSINATURAS_REGEX``
sobre o arquivo consolidado normalizado, sem alterar nada — é só um relatório
de leitura, usado para decidir se o piloto está pronto para virar produção.

Uso:
    uv run etl-boletim-medicao/validar_formato_colunas.py [caminho_consolidado.xlsx]

Se nenhum caminho for informado, usa o consolidado mais recente encontrado em
``D:\\base-geral\\homologacao\\boletim-medicao\\base-consolidado``.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from normalizar_boletim import (
    ASSINATURAS_REGEX,
    _match_processo,
    _validar_origem_lancamento_final,
    _validar_periodo_medicao_final,
)


RAIZ_HOMOLOGACAO = Path(r'D:\base-geral\homologacao\boletim-medicao')
AMOSTRA_MAX = 8


def _localizar_consolidado_mais_recente(pasta: Path) -> Path:
    candidatos = sorted(pasta.glob('boletim-medicao-consolidado_*.xlsx'))
    if not candidatos:
        raise FileNotFoundError(f"Nenhum consolidado encontrado em {pasta}")
    return candidatos[-1]


def _conformidade_coluna(serie: pd.Series, coluna: str) -> dict:
    valores = serie.fillna('').astype('string').str.strip()
    preenchidos = valores.ne('')
    total_preenchido = int(preenchidos.sum())

    if coluna == 'periodo_medicao':
        conforme = valores.map(
            lambda v: (not pd.isna(_validar_periodo_medicao_final(v))) if v else False
        )
    elif coluna == 'origem_lancamento':
        conforme = valores.map(
            lambda v: (not pd.isna(_validar_origem_lancamento_final(v))) if v else False
        )
    elif coluna == 'processo':
        # ASSINATURAS_REGEX['processo'] é deliberadamente mais estreito
        # (só "texto - texto") -- controla quem pode ser CANDIDATO de
        # realocação pra dentro de processo, não o formato final aceito
        # na própria coluna. O formato real aceito (o que o pipeline de
        # fato preserva) é _match_processo, que também inclui a whitelist
        # de códigos curtos conhecidos (CUSTEIO, GEOM, SMC, etc.).
        conforme = valores.map(lambda v: _match_processo(v) if v else False)
    else:
        pattern = ASSINATURAS_REGEX[coluna]
        conforme = valores.str.fullmatch(pattern, na=False)

    conforme_preenchido = conforme & preenchidos
    nao_conforme = preenchidos & ~conforme

    amostra = list(valores[nao_conforme].unique()[:AMOSTRA_MAX])

    return {
        'coluna': coluna,
        'total_linhas': len(serie),
        'preenchidos': total_preenchido,
        'conformes': int(conforme_preenchido.sum()),
        'nao_conformes': int(nao_conforme.sum()),
        'amostra_nao_conforme': amostra,
    }


def gerar_relatorio(caminho_consolidado: Path) -> pd.DataFrame:
    df = pd.read_excel(caminho_consolidado, sheet_name='Boletins', dtype=str)

    colunas_para_checar = (
        list(ASSINATURAS_REGEX.keys()) + ['periodo_medicao', 'origem_lancamento']
    )
    linhas_relatorio = []
    for coluna in colunas_para_checar:
        if coluna not in df.columns:
            linhas_relatorio.append({
                'coluna': coluna,
                'total_linhas': len(df),
                'preenchidos': 0,
                'conformes': 0,
                'nao_conformes': 0,
                'amostra_nao_conforme': ['<coluna ausente no consolidado>'],
            })
            continue
        linhas_relatorio.append(_conformidade_coluna(df[coluna], coluna))

    return pd.DataFrame(linhas_relatorio)


def main() -> None:
    if len(sys.argv) > 1:
        caminho = Path(sys.argv[1])
    else:
        caminho = _localizar_consolidado_mais_recente(RAIZ_HOMOLOGACAO / 'base-consolidado')

    print(f"Lendo consolidado: {caminho}")
    relatorio = gerar_relatorio(caminho)

    pd.set_option('display.max_colwidth', 60)
    pd.set_option('display.width', 160)

    print("\n=== CONFORMIDADE DE FORMATO POR COLUNA ===")
    for _, linha in relatorio.iterrows():
        pct = (
            100 * linha['conformes'] / linha['preenchidos']
            if linha['preenchidos'] else 100.0
        )
        print(
            f"{linha['coluna']:<28} preenchidos={linha['preenchidos']:>6} "
            f"conformes={linha['conformes']:>6} ({pct:5.1f}%) "
            f"nao_conformes={linha['nao_conformes']:>5}"
        )
        if linha['nao_conformes']:
            print(f"    amostra: {linha['amostra_nao_conforme']}")

    caminho_saida = caminho.parent / (caminho.stem + '_relatorio_formato.csv')
    relatorio.drop(columns=['amostra_nao_conforme']).to_csv(
        caminho_saida, index=False, encoding='utf-8-sig'
    )
    print(f"\nResumo salvo em: {caminho_saida}")


if __name__ == '__main__':
    main()
