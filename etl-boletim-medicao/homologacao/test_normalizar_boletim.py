import pandas as pd

from normalizar_boletim import normalizar_dataframe


def test_data_envio_bm_converte_para_formato_brasileiro():
    # Pedido do usuário em 2026-09-01: data_envio_bm sempre no formato
    # dd/mm/yyyy (pedido inicial era formato americano MM/DD/YYYY,
    # revertido no mesmo dia), seja a origem um datetime nativo do Excel
    # (str -> 'YYYY-MM-DD ...', formato intermediário com traço) ou um
    # texto 'DD/MM/YYYY' (formato intermediário com ponto).
    original = pd.DataFrame([
        {'data_envio_bm': '2026-06-18', 'boletim': '1011865031'},
        {'data_envio_bm': '18/06/2026', 'boletim': '1011865032'},
    ])

    normalizado, _ = normalizar_dataframe(original)

    assert normalizado.loc[0, 'data_envio_bm'] == '18/06/2026'
    assert normalizado.loc[1, 'data_envio_bm'] == '18/06/2026'


def test_data_envio_bm_nao_inverte_dia_mes_duas_vezes():
    # A conversão final pro padrão dd/mm/yyyy roda só 1x (fora de
    # `_normalizar_data`, que é chamada 2x internamente) -- garante que
    # uma data já convertida (ano-mês-dia -> dia/mês/ano) não seja
    # reinterpretada como se já estivesse em outra ordem e tenha dia/mês
    # trocados de novo. Dia=25 só é válido como dia (não existe mês 25)
    # -- se invertesse 2x, o resultado ficaria óbvio/errado.
    original = pd.DataFrame([{'data_envio_bm': '2026-03-25'}])

    normalizado, _ = normalizar_dataframe(original)

    assert normalizado.loc[0, 'data_envio_bm'] == '25/03/2026'


def test_data_envio_bm_sentinela_1900_vira_nulo():
    # Achado real na homologação de 2026-09-01: 42.110 linhas (23% do
    # preenchimento de data_envio_bm) tinham a sentinela de data vazia
    # do Excel (dia zero da serial numérica, "01/01/1900") -- não é data
    # de negócio real, vira nulo, mesma doutrina já aplicada aos
    # sentinelas de texto "NÃO INFORMADO"/"(EM BRANCO)". Cobre as 2
    # formas intermediárias possíveis (traço de datetime nativo, ponto
    # de texto DD/MM/YYYY).
    original = pd.DataFrame([
        {'data_envio_bm': '1900-01-01'},
        {'data_envio_bm': '01.01.1900'},
    ])

    normalizado, _ = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'data_envio_bm'])
    assert pd.isna(normalizado.loc[1, 'data_envio_bm'])


def test_origem_lancamento_aceita_pep_e_numero_sap_remove_sufixo_excel():
    # Achado real na homologação de 2026-09-01: 2 formatos válidos
    # (código PEP "PA-.../AP-..." e número SAP puro de 7-10 dígitos) e
    # 5.881 linhas com sufixo ".0" de float do Excel nunca removido
    # (mesma limpeza já aplicada a boletim/contrato).
    original = pd.DataFrame([
        {'origem_lancamento': 'PA-2652502SMC1.1.0125.D'},
        {'origem_lancamento': 'AP-2402304EME1.F.0056.D'},
        {'origem_lancamento': '6001429759.0'},
    ])

    normalizado, _ = normalizar_dataframe(original)

    assert normalizado.loc[0, 'origem_lancamento'] == 'PA-2652502SMC1.1.0125.D'
    assert normalizado.loc[1, 'origem_lancamento'] == 'AP-2402304EME1.F.0056.D'
    assert normalizado.loc[2, 'origem_lancamento'] == '6001429759'


def test_origem_lancamento_fora_do_padrao_vira_nulo():
    original = pd.DataFrame([{'origem_lancamento': 'texto qualquer sem formato'}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'origem_lancamento'])
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'origem_lancamento')
    ).any()


def test_origem_lancamento_pep_exige_tamanho_exato_de_segmento():
    # Achado real do usuário (2026-09-01): o formato PEP tem tamanho FIXO
    # de segmento -- "xx-xxxxxxxxxxx.x.xxxx.x" (2-11.1.4.1 caracteres),
    # confirmado contra 154.959 valores reais sem exceção. Um segmento de
    # tamanho diferente (aqui, 5 dígitos em vez de 4 no 3º grupo) não é
    # um formato de negócio válido e vira nulo -- o regex anterior
    # (`[A-Z0-9]+` de tamanho livre) aceitava isso indevidamente.
    original = pd.DataFrame([{'origem_lancamento': 'PA-2652502SMC1.1.00125.D'}])

    normalizado, _ = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'origem_lancamento'])


def test_normaliza_cp_e_iva_deslocados_de_forma_univoca():
    original = pd.DataFrame([{
        'cp': 'CP-INVALIDO',
        'iva': None,
        'processo': 'AB12',
        'contrato': 'A1',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'cp'] == 'AB12'
    assert normalizado.loc[0, 'iva'] == 'A1'
    assert set(auditoria['tipo']) == {'CORRIGIDA'}
    assert auditoria.loc[auditoria['coluna_destino'] == 'cp', 'valor_anterior_destino'].iloc[0] == 'CP-INVALIDO'


def test_nao_corrige_quando_ha_ambiguidade():
    original = pd.DataFrame([{
        'cp': None,
        'iva': None,
        'processo': 'AB12',
        'contrato': 'CD34',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'cp'])
    assert 'AMBIGUA' in set(auditoria['tipo'])


def test_corrige_layout_rotacionado_com_assinatura_completa():
    original = pd.DataFrame([{
        'distribuidora': 'arquivo.xlsx',
        'regional': 'Eqtl PA',
        'tipo_medicao': 'Nordeste',
        'estrutura': 'Investimento',
        'medicao': 'Produtividade',
        'periodo_medicao': 'Ciclo',
        'parceiro': '01/05/2026 A 31/05/2026',
        'municipio': 'ELINSA',
        'equipe': 'CAMETA',
        'desc_nota_fiscal': 'Construcao',
        'boletim': 'Servico',
        'valor_bm': '1011504739,0',
        'data_envio_bm': '123,46',
        'origem_lancamento': None,
        'cp': 'PA-2608302GED1.2.0004.D',
        'iva': 'CP02',
        'domicilio_fiscal': 'I9',
        'codigo_tarifa_fiscal': 'PA 1502756',
        'identificador_medicao': 'LC116_07.02',
        'identificador_agrupamento': '602.2.2.2026.04.000762',
        'contrato': '105526-1-10042026-49728',
        'processo': '4600026186',
        'texto_boletim': 'Ancora Tecnica',
        'arquivo_origem': None,
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'distribuidora'] == 'Eqtl PA'
    assert normalizado.loc[0, 'cp'] == 'CP02'
    assert normalizado.loc[0, 'iva'] == 'I9'
    assert normalizado.loc[0, 'domicilio_fiscal'] == 'PA 1502756'
    assert normalizado.loc[0, 'codigo_tarifa_fiscal'] == 'LC116_07.02'
    assert normalizado.loc[0, 'arquivo_origem'] == 'arquivo.xlsx'
    assert 'LAYOUT_ROTACIONADO_CORRIGIDO' in set(auditoria['tipo'])


def test_remove_nota_fiscal_deslocada_antes_de_cp():
    original = pd.DataFrame([{
        'cp': '6020050053', 'iva': 'CP04', 'domicilio_fiscal': 'N9',
        'codigo_tarifa_fiscal': 'PA 1506807',
        'identificador_medicao': 'LC116_07.02',
        'identificador_agrupamento': '602.1.2.2026.07.001769',
        'contrato': '110256-1-16072026-57974',
        'processo': '4600029715', 'texto_boletim': 'Âncora Técnica',
        'arquivo_origem': 'ELI>OEST>SNT>NÃO CADASTRADO>JUL26',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'cp'] == 'CP04'
    assert normalizado.loc[0, 'iva'] == 'N9'
    assert normalizado.loc[0, 'domicilio_fiscal'] == 'PA 1506807'
    assert normalizado.loc[0, 'codigo_tarifa_fiscal'] == 'LC116_07.02'
    assert 'NOTA_FISCAL_DESLOCADA_REMOVIDA' in set(auditoria['tipo'])


def test_processo_usa_formato_textual_e_nao_valor_fechado():
    original = pd.DataFrame([{
        'processo': 'Âncora Técnica - Recuperação de Energia',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'processo'] == 'Âncora Técnica - Recuperação de Energia'
    assert auditoria.empty


def test_valor_sem_coluna_compativel_vira_nulo():
    original = pd.DataFrame([{
        'cp': 'FORMATO-SEM-CORRESPONDENCIA',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'cp'])
    assert (auditoria['tipo'] == 'INVALIDO_NAO_NORMALIZADO').any()


def test_contrato_aceita_numero_sap_puro_e_remove_sufixo_excel():
    # Achado real na homologação em escala de 2026-08-31: a assinatura
    # antiga de contrato exigia formato decimal com vírgula, que nunca
    # ocorre de verdade (o campo é um número SAP puro) -- anulava 100%
    # do campo (~268 mil linhas no corpus histórico completo). "valor_bm"
    # aqui é só para a linha não ficar com uma única coluna preenchida --
    # nesse caso degenerado, "contrato" (só dígitos) vira o único candidato
    # a preencher "valor_bm" (cuja assinatura também aceita dígitos soltos),
    # sendo roubado por um mecanismo não relacionado a este teste.
    original = pd.DataFrame([{
        'contrato': '4600026186,0',
        'valor_bm': '150,00',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'contrato'] == '4600026186'
    assert 'tipo' not in auditoria.columns or not (
        auditoria['tipo'] == 'INVALIDO_NAO_NORMALIZADO'
    ).any()


def test_boletim_nao_e_roubado_para_valor_bm():
    # Achado real na homologação em escala de 2026-08-31: a assinatura
    # antiga de valor_bm aceitava dígito solto sem decimal de qualquer
    # tamanho -- um boletim de 10 dígitos passava por "valor_bm" válido e
    # era roubado pra lá (2.086 linhas), apagando o boletim e derrubando a
    # linha inteira depois via remover_sem_boletim.
    original = pd.DataFrame([{'boletim': '1011587571', 'valor_bm': None}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'boletim'] == '1011587571'
    assert pd.isna(normalizado.loc[0, 'valor_bm'])
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'valor_bm')
    ).any()


def test_valor_bm_com_decimal_aceita_parte_inteira_grande():
    # "1007155,63" é um valor real do corpus histórico -- não pode ser
    # rejeitado só porque a parte inteira tem mais de 6 dígitos; o cap de
    # 6 dígitos vale só pra dígito solto SEM decimal (risco de colisão
    # com boletim), nunca pra valor com decimal (boletim nunca tem vírgula).
    original = pd.DataFrame([{'valor_bm': '1007155,63'}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'valor_bm'] == '1007155,63'
    assert auditoria.empty


def test_valor_pequeno_nao_e_roubado_para_boletim():
    # Achado real na homologação em escala de 2026-08-31: a assinatura de
    # candidato pra "boletim" aceitava dígito solto de qualquer tamanho --
    # um valor de "cp" malformado ("20"/"40", 219 linhas) ou um valor
    # pequeno de "valor_bm" (3-6 dígitos, 312 linhas) eram roubados pra
    # boletim, que só tem 2 tamanhos reais no corpus (7 ou 10 dígitos).
    # "cp" acaba virando nulo por conta própria ("20" não é um cp válido,
    # [A-Z]{2}[0-9]{2} exige letras) -- comportamento correto e separado
    # do que este teste cobre; só "valor_bm" tem forma válida por si só.
    original = pd.DataFrame([{'boletim': None, 'cp': '20', 'valor_bm': '953'}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'boletim'])
    assert normalizado.loc[0, 'valor_bm'] == '953'
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'boletim')
    ).any()


def test_origem_lancamento_nao_e_roubado_para_boletim():
    # Achado real na homologação em escala de 2026-08-31: código de
    # coletor de custo/PEP em origem_lancamento também é um número SAP
    # puro de 10 dígitos (mesmo tamanho de um boletim real) -- 9.791
    # linhas indevidamente copiadas pra boletim, apagando o código
    # original. Ficam excluídos como candidato só nessa direção
    # específica (origem_lancamento -> boletim).
    original = pd.DataFrame([{'boletim': None, 'origem_lancamento': '6020018020'}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'boletim'])
    assert normalizado.loc[0, 'origem_lancamento'] == '6020018020'
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'boletim')
    ).any()


def test_periodo_medicao_nao_e_preenchido_por_outra_coluna():
    # Achado real na homologação em escala de 2026-08-31: a assinatura
    # antiga de periodo_medicao usava str.contains frouxo demais -- 100%
    # das 26.667 "correções" vindas de data_envio_bm/origem_lancamento/
    # identificador_agrupamento produziam valor fora do padrão "dd.mm.yyyy
    # a dd.mm.yyyy" (nenhuma dessas colunas contém um período de medição
    # de verdade; um "1900-01-01" -- sentinela de data vazia do Excel --
    # chegou a vazar assim).
    original = pd.DataFrame([{
        'periodo_medicao': None,
        'data_envio_bm': '1900-01-01',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'periodo_medicao'])
    # '1900-01-01' é a sentinela de data vazia do Excel -- desde
    # 2026-09-01 o passo final de conversão pro formato americano
    # também converte a sentinela em nulo real (não fabrica data).
    assert pd.isna(normalizado.loc[0, 'data_envio_bm'])
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'periodo_medicao')
    ).any()


def test_parceiro_nao_e_roubado_para_processo():
    # Achado real na homologação de 2026-09-01: leiautes sem coluna
    # "Processo" (ex. "BOLETIM DE MEDIÇÃO") têm "parceiro" formatado como
    # "texto - texto" (ex. "ELINSA - ELETROTÉCNICA INDUSTR..."), mesma
    # assinatura de `processo` -- 86.054 linhas (22% do dataset) tiveram o
    # nome da parceira indevidamente copiado pra dentro de `processo`, um
    # valor que nunca existiu na origem pra esse campo.
    original = pd.DataFrame([{
        'processo': None,
        'parceiro': 'ELINSA - ELETROTECNICA INDUSTR',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'processo'])
    assert normalizado.loc[0, 'parceiro'] == 'ELINSA - ELETROTECNICA INDUSTR'
    assert 'coluna_destino' not in auditoria.columns or not (
        (auditoria['tipo'] == 'CORRIGIDA')
        & (auditoria['coluna_destino'] == 'processo')
    ).any()


def test_periodo_medicao_fora_do_padrao_vira_nulo():
    # "data_envio_bm" já vem preenchido de propósito: numa linha com só
    # periodo_medicao populado, esse valor (formato de data ISO) seria
    # roubado pra dentro de data_envio_bm (vazio, "precisando de correção"),
    # o que testaria outro mecanismo, não a validação final de
    # periodo_medicao que este teste quer cobrir.
    original = pd.DataFrame([{
        'periodo_medicao': '2025-04-16',
        'data_envio_bm': '01.01.2026',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'periodo_medicao'])
    assert (
        (auditoria['tipo'] == 'INVALIDO_NAO_NORMALIZADO')
        & (auditoria['coluna_destino'] == 'periodo_medicao')
    ).any()


def test_sentinelas_da_origem_viram_nulo():
    # Achado real na homologação em escala de 2026-08-31: "NÃO INFORMADO"
    # e "(EM BRANCO)" já chegam assim da própria planilha de origem (não
    # é o pipeline fabricando) -- 14.297 linhas em "estrutura", 984+728+11
    # em "regional", 11 em "tipo_medicao"/"municipio"/"desc_nota_fiscal".
    # Mesma doutrina de nulo real do projeto: sentinela de apresentação
    # não é dado de negócio, vira nulo.
    original = pd.DataFrame([{
        'estrutura': 'NÃO INFORMADO',
        'regional': '  não informado  ',
        'tipo_medicao': '( EM BRANCO )',
        'municipio': '(Em branco)',
        'desc_nota_fiscal': 'Serviço real de verdade',
    }])

    normalizado, _ = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'estrutura'])
    assert pd.isna(normalizado.loc[0, 'regional'])
    assert pd.isna(normalizado.loc[0, 'tipo_medicao'])
    assert pd.isna(normalizado.loc[0, 'municipio'])
    assert normalizado.loc[0, 'desc_nota_fiscal'] == 'Serviço real de verdade'


def test_contrato_garbage_vira_nulo_sem_afetar_outras_colunas():
    # Achado real na homologação em escala de 2026-08-31: 11 linhas com o
    # placeholder "( EM BRANCO )" da própria planilha de origem em
    # "contrato" -- validado só no final (`_validar_contrato_final`), sem
    # passar pelo mecanismo de busca-e-realocação entre colunas.
    original = pd.DataFrame([{
        'contrato': '( EM BRANCO )',
        'valor_bm': '150,00',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'contrato'])
    assert normalizado.loc[0, 'valor_bm'] == '150,00'
    assert (
        (auditoria['tipo'] == 'INVALIDO_NAO_NORMALIZADO')
        & (auditoria['coluna_destino'] == 'contrato')
    ).any()


def test_processo_aceita_codigo_curto_conhecido():
    # Achado real na homologação em escala de 2026-08-31: além do formato
    # composto "texto - texto", processo também aparece como um código
    # curto isolado (53.728 linhas rejeitadas indevidamente no corpus
    # histórico, só 8 valores distintos, sem cauda longa).
    original = pd.DataFrame([{'processo': 'GEOM'}])

    normalizado, auditoria = normalizar_dataframe(original)

    assert normalizado.loc[0, 'processo'] == 'GEOM'
    assert auditoria.empty


def test_valor_com_mais_de_um_destino_vira_nulo():
    original = pd.DataFrame([{
        'cp': 'FORMATO-INVALIDO',
        'processo': 'AB12',
        'contrato': 'CD34',
    }])

    normalizado, auditoria = normalizar_dataframe(original)

    assert pd.isna(normalizado.loc[0, 'cp'])
    assert (auditoria['tipo'] == 'INVALIDO_AMBIGUO_NAO_NORMALIZADO').any()
