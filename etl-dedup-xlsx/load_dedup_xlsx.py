"""
load_dedup_xlsx.py
-------------------
Utilitário para localizar e remover, de uma pasta de origem, arquivos .xlsx
duplicados ou com nome quase idêntico — mas SEMPRE validando a duplicidade
pelo CONTEÚDO real das planilhas, nunca apenas pelo nome do arquivo.

Regras:
- Dois arquivos só são considerados duplicados de fato quando o conteúdo
  (valores de todas as células, em todas as abas visíveis) é idêntico.
- Pares com nome muito parecido mas conteúdo diferente NÃO são removidos
  automaticamente — apenas relatados, para revisão manual.
- Arquivos removidos são MOVIDOS para uma pasta de duplicados (nunca
  excluídos direto), seguindo o mesmo princípio de arquivamento antes de
  remover usado nos demais pipelines deste projeto.
"""

import hashlib
import re
import shutil
import unicodedata
import warnings
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore", module="openpyxl")


# =====================================================================
# NORMALIZAÇÃO E SIMILARIDADE DE NOME
# =====================================================================

def _normalizar_nome(nome_arquivo: str) -> str:
    """Remove extensão, acentos, pontuação e marcadores comuns de cópia
    (' (1)', '- Copia', '_copy 2', etc.) para permitir comparar nomes de
    arquivo que representam o "mesmo documento"."""
    nome = Path(nome_arquivo).stem.strip().lower()
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    nome = re.sub(r'\(\d+\)', ' ', nome)
    nome = re.sub(r'[-_\s]*c[o0]pia\s*\d*', ' ', nome)
    nome = re.sub(r'[-_\s]*copy\s*\d*', ' ', nome)
    nome = re.sub(r'[^a-z0-9]+', ' ', nome)
    return nome.strip()


def _nomes_parecidos(nome_a: str, nome_b: str, limiar: float) -> bool:
    if not nome_a or not nome_b:
        return False
    if nome_a == nome_b:
        return True
    return SequenceMatcher(None, nome_a, nome_b).ratio() >= limiar


# =====================================================================
# HASH DE CONTEÚDO (IGNORA METADADOS/BYTES DO .XLSX)
# =====================================================================

def _hash_conteudo(caminho_arquivo: Path) -> str | None:
    """Gera um hash SHA-256 a partir dos VALORES de todas as células de
    todas as abas visíveis do arquivo (ignora abas ocultas, estilos e
    metadados internos do .xlsx, como datas de criação/modificação do
    documento). Dois arquivos com o mesmo conteúdo visível sempre geram o
    mesmo hash, mesmo que os bytes do .xlsx sejam diferentes.
    Retorna None se o arquivo não puder ser lido (corrompido/em uso)."""
    try:
        wb = openpyxl.load_workbook(caminho_arquivo, data_only=True, read_only=True)
    except Exception:
        return None

    hasher = hashlib.sha256()
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible':
                continue
            hasher.update(ws.title.encode('utf-8', errors='ignore'))
            for row in ws.iter_rows(values_only=True):
                for valor in row:
                    hasher.update(repr(valor).encode('utf-8', errors='ignore'))
                hasher.update(b'\x1e')  # separador de linha
            hasher.update(b'\x1d')  # separador de aba
    except Exception:
        return None
    finally:
        wb.close()

    return hasher.hexdigest()


def _parece_copia(caminho: Path) -> bool:
    """Heurística para reconhecer arquivos claramente marcados como cópia
    (ex.: 'Boletim (1).xlsx', 'Boletim - Copia.xlsx')."""
    nome = caminho.stem.lower()
    return bool(re.search(r'\(\d+\)|c[o0]pia|copy', nome))


def _escolher_arquivo_a_manter(arquivos: list) -> Path:
    """Dentre um grupo de arquivos com conteúdo idêntico, escolhe qual
    manter na origem: prioriza o nome que NÃO parece cópia, depois o mais
    antigo (provável original), depois o nome em ordem alfabética."""
    candidatos = sorted(arquivos, key=lambda p: (_parece_copia(p), p.stat().st_mtime, p.name))
    return candidatos[0]


# =====================================================================
# PIPELINE PRINCIPAL
# =====================================================================

def remover_duplicados_xlsx(
    pasta_origem_input: str,
    pasta_duplicados_input: str,
    pasta_unicos_input: str | None = None,
    limiar_similaridade_nome: float = 0.85,
    dry_run: bool = False,
) -> None:
    print("\n" + "=" * 60)
    print("REMOÇÃO DE ARQUIVOS XLSX DUPLICADOS".center(60))
    print("=" * 60)

    pasta_origem = Path(pasta_origem_input)
    pasta_duplicados = Path(pasta_duplicados_input)
    pasta_unicos = Path(pasta_unicos_input) if pasta_unicos_input else None

    if not pasta_origem.exists():
        print(f"\n[❌] Erro Crítico: pasta de origem não encontrada:\n{pasta_origem}")
        return

    arquivos = [f for f in pasta_origem.glob('*.xlsx') if not f.name.startswith('~$')]

    if len(arquivos) < 2:
        print("Menos de 2 arquivos .xlsx encontrados na origem — nada a comparar.")
        return

    print(f"Encontrados {len(arquivos)} arquivo(s) .xlsx. Calculando hash de conteúdo...\n")

    # ---------------------------------------------------------
    # FASE 1: HASH DE CONTEÚDO
    # ---------------------------------------------------------
    hashes = {}
    for caminho in arquivos:
        h = _hash_conteudo(caminho)
        if h is None:
            print(f" -> [!] Não foi possível ler '{caminho.name}' (corrompido ou em uso). Ignorado.")
            continue
        hashes[caminho] = h

    # ---------------------------------------------------------
    # FASE 2: AGRUPAMENTO POR CONTEÚDO IDÊNTICO
    # ---------------------------------------------------------
    grupos_por_hash = {}
    for caminho, h in hashes.items():
        grupos_por_hash.setdefault(h, []).append(caminho)

    grupos_duplicados = [g for g in grupos_por_hash.values() if len(g) > 1]

    # ---------------------------------------------------------
    # FASE 3: NOMES PARECIDOS COM CONTEÚDO DIFERENTE (SÓ RELATÓRIO)
    # ---------------------------------------------------------
    nomes_normalizados = {c: _normalizar_nome(c.name) for c in hashes}
    avisos_nome_parecido = []

    # Comparar nome a nome seria O(n²) e inviável em pastas com milhares de
    # arquivos. Como nomes "quase idênticos" preservam o início do nome
    # (o que muda são sufixos como "(1)", "_v9", "_final"), agrupamos por
    # prefixo do nome normalizado e só comparamos dentro do mesmo grupo.
    baldes_por_prefixo = {}
    for caminho, nome in nomes_normalizados.items():
        chave = nome[:8]
        baldes_por_prefixo.setdefault(chave, []).append(caminho)

    for candidatos in baldes_por_prefixo.values():
        if len(candidatos) < 2:
            continue
        for i in range(len(candidatos)):
            for j in range(i + 1, len(candidatos)):
                a, b = candidatos[i], candidatos[j]
                if hashes[a] == hashes[b]:
                    continue  # já tratado como duplicado de conteúdo na Fase 2
                if _nomes_parecidos(nomes_normalizados[a], nomes_normalizados[b], limiar_similaridade_nome):
                    avisos_nome_parecido.append((a.name, b.name))

    LIMITE_AVISOS_IMPRESSOS = 50
    if avisos_nome_parecido:
        print("Nomes parecidos, porém com CONTEÚDO DIFERENTE (não removidos automaticamente):")
        for nome_a, nome_b in avisos_nome_parecido[:LIMITE_AVISOS_IMPRESSOS]:
            print(f"  [AVISO] '{nome_a}'  ~  '{nome_b}'")
        if len(avisos_nome_parecido) > LIMITE_AVISOS_IMPRESSOS:
            print(f"  ... e mais {len(avisos_nome_parecido) - LIMITE_AVISOS_IMPRESSOS} par(es) parecido(s) (total: {len(avisos_nome_parecido)}).")
        print()

    if not grupos_duplicados and not pasta_unicos:
        print("Nenhum arquivo com conteúdo duplicado encontrado.")
        return

    # ---------------------------------------------------------
    # FASE 4: CÓPIA DOS ÚNICOS (1 ARQUIVO POR CONTEÚDO DISTINTO)
    # ---------------------------------------------------------
    total_unicos = 0
    if pasta_unicos:
        if not dry_run:
            pasta_unicos.mkdir(parents=True, exist_ok=True)

        print(f"Copiando {len(grupos_por_hash)} arquivo(s) único(s) (por conteúdo) para '{pasta_unicos.name}'...\n")
        for grupo in grupos_por_hash.values():
            mantido = _escolher_arquivo_a_manter(grupo) if len(grupo) > 1 else grupo[0]
            if dry_run:
                print(f" -> [DRY-RUN] copiaria: {mantido.name}")
                continue
            destino = pasta_unicos / mantido.name
            try:
                shutil.copy2(str(mantido), str(destino))
                total_unicos += 1
            except Exception as e:
                print(f" -> [!] Falha ao copiar '{mantido.name}' para únicos: {e}")
        print()

    # ---------------------------------------------------------
    # FASE 5: REMOÇÃO DOS DUPLICADOS (MOVE, NUNCA EXCLUI DIRETO)
    # ---------------------------------------------------------
    total_removidos = 0
    if grupos_duplicados:
        if not dry_run:
            pasta_duplicados.mkdir(parents=True, exist_ok=True)

        print(f"Encontrado(s) {len(grupos_duplicados)} grupo(s) de arquivo(s) com conteúdo idêntico:\n")

        for grupo in grupos_duplicados:
            mantido = _escolher_arquivo_a_manter(grupo)
            duplicados = [f for f in grupo if f != mantido]

            print(f" -> MANTIDO: {mantido.name}")
            for duplicado in duplicados:
                if dry_run:
                    print(f"    [DRY-RUN] removeria: {duplicado.name}")
                    continue

                destino = pasta_duplicados / duplicado.name
                if destino.exists():
                    timestamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
                    destino = pasta_duplicados / f"{destino.stem}_{timestamp}{destino.suffix}"

                try:
                    shutil.move(str(duplicado), str(destino))
                    print(f"    REMOVIDO (movido p/ '{pasta_duplicados.name}'): {duplicado.name}")
                    total_removidos += 1
                except Exception as e:
                    print(f"    [!] Falha ao mover '{duplicado.name}': {e}")

    print("\n" + "=" * 60)
    print("RESUMO".center(60))
    print("=" * 60)
    print(f" -> Arquivos analisados:        {len(arquivos)}")
    print(f" -> Grupos de duplicados:       {len(grupos_duplicados)}")
    if dry_run:
        print(" -> Arquivos removidos:         0 (dry-run, nenhum arquivo foi movido)")
        if pasta_unicos:
            print(" -> Arquivos únicos:            0 (dry-run, nenhum arquivo foi copiado)")
    else:
        print(f" -> Arquivos removidos:         {total_removidos}")
        if pasta_unicos:
            print(f" -> Arquivos únicos copiados:   {total_unicos}")
    print(f" -> Nomes parecidos (revisão):  {len(avisos_nome_parecido)}")
    print("=" * 60)
