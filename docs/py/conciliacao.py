"""
Núcleo da conciliação financeira (CIB / BRB).

Reproduz, de forma automática, a planilha
"Conciliação Boleto c. movimento":

    1. Extrato BRB (TXT)
         CRED PIX QR CODE DINAMICO  -> recebimentos Bolepix
         CREDITO COBRANCA BRB       -> recebimentos Bolebarras
         soma = Total pagamentos QR Code + Código de Barras

    2. Francesinha Bolepix (XLS/XLSX/TXT)
         detalhe dos PIX recebidos, com nome do pagador

    3. DCB - arquivo de retorno da cobrança (CNAB 400)
         títulos liquidados = Total arquivo de retorno

    4. Divergência
         Total pagamentos - Total arquivo de retorno
         = Bolepix não baixado, identificado item a item

Este módulo não depende de interface nem de pandas: roda igual
no CLI (main.py), no servidor (app.py) e no navegador, via
Pyodide (docs/).

As funções de leitura aceitam caminho (str/Path), bytes ou
objeto com .read(), para funcionar com upload de arquivos.
"""

from collections import Counter
from decimal import Decimal
import re

from conciliacao_terracap.planilhas import (
    celula,
    e_planilha,
    ler_planilha,
    nome_arquivo,
    para_texto,
    texto_da_celula,
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Descrições reconhecidas no extrato BRB.
DESCRICAO_PIX = "CRED PIX QR CODE DINAMICO"
DESCRICAO_COBRANCA = "CREDITO COBRANCA BRB"

ZERO = Decimal("0.00")
UM_CENTAVO = Decimal("0.01")
CENTAVOS = Decimal("100")


# ------------------------------------------------------------
# Layout do DCB (arquivo de retorno CNAB 400 do BRB)
#
# Posições confirmadas contra o arquivo real
# DCB_1219001012_28092026_030540.txt e contra a planilha de
# conciliação manual de 28/09/2026:
#
#   481 títulos com valor pago > 0  ->  R$ 4.319.046,36
#   que é exatamente o "Total arquivo de retorno".
#
# Posições em base 1, inclusivas, como na documentação CNAB.
# ------------------------------------------------------------

CAMPOS_DCB = {
    "tipo_registro":   (1, 1),
    "tipo_inscricao":  (2, 3),
    "doc_pagador":     (4, 17),
    "agencia_conta":   (21, 37),
    "uso_empresa":     (38, 62),
    "nosso_numero":    (71, 82),
    "conta":           (83, 92),
    "seu_numero":      (93, 107),
    "ocorrencia":      (108, 110),
    "data_ocorrencia": (111, 118),
    "valor_pago":      (258, 270),
    "data_credito":    (300, 307),
    "valor_liquido":   (308, 320),
    "sequencial":      (395, 400),
}

TAMANHO_REGISTRO_DCB = 400

# Ocorrências observadas no retorno do BRB.
#
# ATENÇÃO: o critério de "título liquidado" usado pela
# conciliação NÃO é o código, e sim valor pago > 0. Isso porque
# 005 e 105 representam liquidação, e somar apenas 005 deixaria
# 2 títulos (R$ 5.740,60) de fora do total do arquivo.
OCORRENCIAS_DCB = {
    "000": "Sem ocorrência / título em carteira",
    "002": "Entrada confirmada",
    "005": "Liquidação",
    "100": "Baixa",
    "105": "Liquidação (segunda via)",
}


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def moeda_para_decimal(valor):
    """
    Converte para Decimal:

        1.234,56      R$ 1.234,56
        1234,56       1234.56
        1.234,56-     (negativo)
        444.47        (float vindo de planilha)
    """

    if valor is None:
        return ZERO

    if isinstance(valor, Decimal):
        return valor

    if isinstance(valor, bool):
        return ZERO

    if isinstance(valor, int):
        return Decimal(valor)

    if isinstance(valor, float):
        return Decimal(str(valor)).quantize(UM_CENTAVO)

    texto = str(valor).strip()

    if not texto or texto.lower() in ("nan", "none", "-"):
        return ZERO

    negativo = texto.startswith("-") or texto.endswith("-")

    for lixo in ("R$", "+", "-", " ", "\xa0", "\t"):
        texto = texto.replace(lixo, "")

    if not texto:
        return ZERO

    # Formato brasileiro: a vírgula é o separador decimal.
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        numero = Decimal(texto)
    except Exception:
        return ZERO

    return -numero if negativo else numero


def centavos_para_decimal(texto):
    """
    Converte um campo CNAB (inteiro em centavos) para Decimal.
    """

    texto = (texto or "").strip()

    if not texto or not texto.isdigit():
        return ZERO

    return Decimal(int(texto)) / CENTAVOS


def formatar_moeda(valor):
    """
    Formata no padrão brasileiro: R$ 1.234,56
    """

    texto = f"{Decimal(valor):,.2f}"

    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")

    return f"R$ {texto}"


def normalizar_nome(nome):
    """
    Maiúsculas, sem acentos e sem espaços duplicados.
    """

    if nome is None:
        return ""

    nome = str(nome).upper().strip()

    substituicoes = {
        "Á": "A", "À": "A", "Â": "A", "Ã": "A", "Ä": "A",
        "É": "E", "È": "E", "Ê": "E", "Ë": "E",
        "Í": "I", "Ì": "I", "Î": "I", "Ï": "I",
        "Ó": "O", "Ò": "O", "Ô": "O", "Õ": "O", "Ö": "O",
        "Ú": "U", "Ù": "U", "Û": "U", "Ü": "U",
        "Ç": "C", "Ñ": "N",
    }

    for origem, destino in substituicoes.items():
        nome = nome.replace(origem, destino)

    return re.sub(r"\s+", " ", nome)


def somar(registros, campo="valor"):
    """
    Soma um campo Decimal de uma lista de dicionários.
    """

    return sum((r[campo] for r in registros), ZERO)


def mascarar_documento(documento):
    """
    Mostra apenas o miolo do CPF/CNPJ, como na planilha manual:

        06351597005  ->  ***.159.700-**
    """

    digitos = re.sub(r"\D", "", str(documento or "")).lstrip("0")

    if len(digitos) == 11:
        return f"***.{digitos[3:6]}.{digitos[6:9]}-**"

    if len(digitos) == 14:
        return f"**.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-**"

    return digitos


# ============================================================
# 1. EXTRATO BRB
# ============================================================

# 28/09/2026     CRED PIX QR CODE DINAMICO     000000     1.777,26+     1.401.531,78+
REGEX_EXTRATO = re.compile(
    r"""
    (?P<data>\d{2}/\d{2}/\d{4})
    \s+
    (?P<descricao>.*?)
    \s+
    (?P<doc>\d{6})
    \s+
    (?P<valor>[\d\.]+,\d{2})\+
    """,
    re.VERBOSE,
)


def ler_extrato_brb(origem):
    """
    Lê o TXT do extrato e separa os créditos em:

        pix       -> CRED PIX QR CODE DINAMICO
        cobranca  -> CREDITO COBRANCA BRB
        outros    -> demais créditos
    """

    texto = para_texto(origem)

    pix = []
    cobranca = []
    outros = []

    for numero_linha, linha in enumerate(texto.splitlines(), start=1):

        match = REGEX_EXTRATO.search(linha)

        if not match:
            continue

        descricao = re.sub(r"\s+", " ", match.group("descricao").strip())

        registro = {
            "linha": numero_linha,
            "data": match.group("data"),
            "descricao": descricao,
            "doc": match.group("doc"),
            "valor": moeda_para_decimal(match.group("valor")),
        }

        if descricao == DESCRICAO_PIX:
            pix.append(registro)

        elif descricao == DESCRICAO_COBRANCA:
            cobranca.append(registro)

        else:
            outros.append(registro)

    total_pix = somar(pix)
    total_cobranca = somar(cobranca)

    return {
        "pix": pix,
        "cobranca": cobranca,
        "outros_creditos": outros,

        "quantidade_pix": len(pix),
        "quantidade_cobranca": len(cobranca),
        "quantidade_outros": len(outros),

        "total_pix": total_pix,
        "total_cobranca": total_cobranca,
        "total_outros": somar(outros),

        # "Total pagamentos QR Code + Código de Barras"
        "total_pagamentos": total_pix + total_cobranca,
    }


# ============================================================
# 2. DCB - ARQUIVO DE RETORNO DA COBRANÇA
# ============================================================

def _fatia(linha, campo):
    """
    Extrai um campo do registro CNAB pelo mapa CAMPOS_DCB.
    """

    inicio, fim = CAMPOS_DCB[campo]

    return linha[inicio - 1:fim]


def _formatar_data(texto):
    """
    DDMMAAAA -> DD/MM/AAAA
    """

    texto = (texto or "").strip()

    if len(texto) == 8 and texto.isdigit() and texto != "00000000":
        return f"{texto[0:2]}/{texto[2:4]}/{texto[4:8]}"

    return ""


def _ler_dcb_cnab(texto):
    """
    Lê o DCB no formato CNAB 400 (um registro por linha).
    """

    registros = []

    for numero_linha, linha in enumerate(texto.splitlines(), start=1):

        linha = linha.rstrip("\r\n")

        if len(linha) < TAMANHO_REGISTRO_DCB:
            continue

        # Tipo 1 = detalhe. 0/02 = header, 9 = trailer.
        if not linha.startswith("1"):
            continue

        valor_pago = centavos_para_decimal(_fatia(linha, "valor_pago"))
        ocorrencia = _fatia(linha, "ocorrencia").strip()

        registros.append({
            "linha": numero_linha,
            "ocorrencia": ocorrencia,
            "ocorrencia_descricao": OCORRENCIAS_DCB.get(
                ocorrencia, f"Ocorrência {ocorrencia}"
            ),
            "nosso_numero": _fatia(linha, "nosso_numero").strip(),
            "seu_numero": _fatia(linha, "seu_numero").strip(),
            "uso_empresa": _fatia(linha, "uso_empresa").strip(),
            "doc_pagador": _fatia(linha, "doc_pagador").strip(),
            "data_ocorrencia": _formatar_data(
                _fatia(linha, "data_ocorrencia")
            ),
            "data_credito": _formatar_data(
                _fatia(linha, "data_credito")
            ),
            "valor": valor_pago,
            "valor_liquido": centavos_para_decimal(
                _fatia(linha, "valor_liquido")
            ),
            "liquidado": valor_pago > ZERO,
        })

    return registros


def _ler_dcb_planilha(origem, nome=""):
    """
    Compatibilidade com o DCB exportado em planilha: procura em
    todas as abas as células cujo conteúdo começa em "005"
    (código de liquidação), preservando zeros à esquerda.
    """

    abas = ler_planilha(origem, nome=nome)

    registros = []

    for nome_aba, linhas in abas.items():

        for indice_linha, linha in enumerate(linhas, start=1):

            for indice_coluna, valor in enumerate(linha, start=1):

                texto = texto_da_celula(valor)

                if not texto.startswith("005"):
                    continue

                registros.append({
                    "linha": indice_linha,
                    "aba": nome_aba,
                    "coluna": indice_coluna,
                    "ocorrencia": "005",
                    "ocorrencia_descricao": OCORRENCIAS_DCB["005"],
                    "nosso_numero": "",
                    "seu_numero": "",
                    "uso_empresa": "",
                    "doc_pagador": "",
                    "data_ocorrencia": "",
                    "data_credito": "",
                    "registro_005": texto,
                    "valor": ZERO,
                    "valor_liquido": ZERO,
                    "liquidado": True,
                })

    return registros


def ler_dcb(origem):
    """
    Lê o DCB em CNAB 400 (TXT) ou em planilha, detectando
    automaticamente o formato.

    Um título é considerado liquidado quando o valor pago é
    maior que zero — critério validado contra o
    "Total arquivo de retorno" da conciliação manual.
    """

    nome = nome_arquivo(origem, padrao="dcb.txt")

    if e_planilha(origem, nome):
        registros = _ler_dcb_planilha(origem, nome=nome)
        formato = "planilha"
    else:
        registros = _ler_dcb_cnab(para_texto(origem))
        formato = "cnab400"

    liquidados = [r for r in registros if r["liquidado"]]

    ocorrencias = Counter(r["ocorrencia"] for r in registros)

    return {
        "formato": formato,
        "registros": registros,
        "liquidados": liquidados,

        "quantidade_registros": len(registros),
        "quantidade_liquidados": len(liquidados),

        # "Total arquivo de retorno"
        "total_liquidado": somar(liquidados),

        "ocorrencias": [
            {
                "codigo": codigo,
                "descricao": OCORRENCIAS_DCB.get(
                    codigo, f"Ocorrência {codigo}"
                ),
                "quantidade": quantidade,
            }
            for codigo, quantidade in sorted(ocorrencias.items())
        ],
    }


# ============================================================
# 3. FRANCESINHA
# ============================================================

COLUNAS_FRANCESINHA = {
    "data": ("data e hora", "data/hora", "data"),
    "tipo": ("tipo transacao", "tipo"),
    "nome": ("nome contraparte", "contraparte", "nome", "cliente"),
    "documento": ("cnpj / cpf", "cnpj/cpf", "cpf/cnpj", "documento"),
    "valor": ("valor",),
}

TIPOS_RECEBIMENTO = ("RECEBIMENTO", "CREDITO")


def _localizar_cabecalho(linhas):
    """
    Encontra a linha de cabeçalho do relatório e devolve
    (índice da linha, {campo: índice da coluna}).
    """

    for indice_linha, linha in enumerate(linhas[:40]):

        celulas = {}

        for indice_coluna, valor in enumerate(linha):

            texto = normalizar_nome(texto_da_celula(valor)).lower()

            if texto:
                celulas[indice_coluna] = texto

        if not celulas:
            continue

        mapa = {}

        for campo, apelidos in COLUNAS_FRANCESINHA.items():

            esperados = [normalizar_nome(a).lower() for a in apelidos]

            for indice_coluna, texto in celulas.items():
                if texto in esperados:
                    mapa[campo] = indice_coluna
                    break

        if "valor" in mapa and "nome" in mapa:
            return indice_linha, mapa

    return None, {}


def _ler_francesinha_planilha(origem, nome=""):
    """
    Lê o relatório "Extrato/Devolução" do Bolepix exportado pelo
    BRB (colunas Data e Hora / Tipo Transação / Nome Contraparte
    / Valor).
    """

    abas = ler_planilha(origem, nome=nome)

    registros = []

    for nome_aba, linhas in abas.items():

        indice_cabecalho, mapa = _localizar_cabecalho(linhas)

        if not mapa:
            continue

        for indice_linha in range(indice_cabecalho + 1, len(linhas)):

            linha = linhas[indice_linha]

            bruto = celula(linha, mapa["valor"])

            if bruto is None:
                continue

            tipo = texto_da_celula(celula(linha, mapa.get("tipo")))
            nome_pagador = texto_da_celula(celula(linha, mapa["nome"]))

            # A última linha do relatório é o total: tem valor,
            # mas não tem contraparte nem tipo de transação.
            if not nome_pagador and not tipo:
                continue

            if tipo and normalizar_nome(tipo) not in TIPOS_RECEBIMENTO:
                continue

            valor = moeda_para_decimal(bruto)

            if valor == ZERO:
                continue

            registros.append({
                "linha": indice_linha + 1,
                "aba": nome_aba,
                "data": texto_da_celula(celula(linha, mapa.get("data"))),
                "tipo": tipo,
                "nome": nome_pagador,
                "nome_normalizado": normalizar_nome(nome_pagador),
                "documento": texto_da_celula(
                    celula(linha, mapa.get("documento"))
                ),
                "valor": valor,
            })

    return registros


DELIMITADORES = (";", "\t", "|", ",")


def _detectar_delimitador(texto):
    """
    Descobre o separador mais consistente entre as linhas.
    """

    linhas = [l for l in texto.splitlines() if l.strip()]

    melhor = ";"
    melhor_pontuacao = 0

    for simbolo in DELIMITADORES:

        pontuacao = sum(1 for linha in linhas if linha.count(simbolo) >= 2)

        if pontuacao > melhor_pontuacao:
            melhor, melhor_pontuacao = simbolo, pontuacao

    return melhor


def _ler_francesinha_texto(
    texto,
    delimitador=None,
    coluna_documento=0,
    coluna_nome=1,
    coluna_valor=2,
):
    """
    Lê a francesinha em formato delimitado:

        documento ; nome ; valor
    """

    delimitador = delimitador or _detectar_delimitador(texto)

    registros = []
    ignoradas = 0

    maior = max(coluna_nome, coluna_valor, coluna_documento)

    for numero_linha, linha in enumerate(texto.splitlines(), start=1):

        if not linha.strip():
            continue

        partes = [p.strip() for p in linha.split(delimitador)]

        if len(partes) <= maior:
            ignoradas += 1
            continue

        valor = moeda_para_decimal(partes[coluna_valor])

        if valor == ZERO:
            ignoradas += 1
            continue

        registros.append({
            "linha": numero_linha,
            "aba": "",
            "data": "",
            "tipo": "",
            "nome": partes[coluna_nome],
            "nome_normalizado": normalizar_nome(partes[coluna_nome]),
            "documento": partes[coluna_documento],
            "valor": valor,
        })

    return registros, ignoradas


def ler_francesinha(origem, delimitador=None):
    """
    Lê a francesinha em planilha (relatório do BRB) ou em texto
    delimitado, detectando o formato automaticamente.
    """

    nome = nome_arquivo(origem, padrao="francesinha.xls")

    ignoradas = 0

    if e_planilha(origem, nome):
        registros = _ler_francesinha_planilha(origem, nome=nome)
        formato = "planilha"
    else:
        registros, ignoradas = _ler_francesinha_texto(
            para_texto(origem),
            delimitador=delimitador,
        )
        formato = "texto"

    return {
        "formato": formato,
        "registros": registros,
        "quantidade": len(registros),
        "linhas_ignoradas": ignoradas,
        "total": somar(registros),
    }


# ============================================================
# 4. CONCILIAÇÃO POR VALOR
# ============================================================

def _indice_por_valor(registros):
    """
    Agrupa os índices dos registros por valor, preservando a
    ordem, para casar cada valor no máximo uma vez.
    """

    indice = {}

    for posicao, registro in enumerate(registros):
        indice.setdefault(registro["valor"], []).append(posicao)

    return indice


def casar_por_valor(esquerda, direita):
    """
    Casa duas listas de registros pelo valor (conjunto múltiplo).

    Cada registro da direita é consumido no máximo uma vez.

    Devolve (casados, sem_par_esquerda, sem_par_direita), onde
    casados é uma lista de pares
    (registro_esquerda, registro_direita).
    """

    disponiveis = _indice_por_valor(direita)
    usados = set()

    casados = []
    sem_par_esquerda = []

    for registro in esquerda:

        fila = disponiveis.get(registro["valor"])

        if fila:
            posicao = fila.pop(0)
            usados.add(posicao)
            casados.append((registro, direita[posicao]))
        else:
            sem_par_esquerda.append(registro)

    sem_par_direita = [
        registro
        for posicao, registro in enumerate(direita)
        if posicao not in usados
    ]

    return casados, sem_par_esquerda, sem_par_direita


# ============================================================
# 5. CONCILIAÇÃO COMPLETA
# ============================================================

def conciliar(extrato=None, dcb=None, francesinha=None, bolebarras=None):
    """
    Executa a conciliação com os arquivos disponíveis.

    Todos os arquivos são opcionais: o que for enviado é
    processado, o que faltar é omitido do resultado.

    Devolve um dicionário com os blocos:

        extrato        totais do extrato BRB
        francesinha    detalhe dos PIX recebidos (Bolepix)
        bolebarras     detalhe da cobrança por código de barras
        dcb            títulos do arquivo de retorno
        pix_detalhado  PIX do extrato com o nome do pagador
        divergencia    Total pagamentos - Total retorno
        avisos         pontos de atenção para o usuário
    """

    avisos = []

    resultado = {
        "extrato": None,
        "francesinha": None,
        "bolebarras": None,
        "dcb": None,
        "pix_detalhado": None,
        "divergencia": None,
        "avisos": avisos,
    }

    # ---------- 1. extrato ----------

    dados_extrato = None

    if extrato is not None:

        dados_extrato = ler_extrato_brb(extrato)
        resultado["extrato"] = dados_extrato

        if not dados_extrato["pix"] and not dados_extrato["cobranca"]:
            avisos.append(
                "Nenhum crédito PIX ou de cobrança foi reconhecido no "
                "extrato. Confirme se o arquivo é o TXT original do "
                "BRB Internet Banking."
            )

    # ---------- 2. francesinha bolepix ----------

    dados_francesinha = None

    if francesinha is not None:

        dados_francesinha = ler_francesinha(francesinha)
        resultado["francesinha"] = dados_francesinha

        if not dados_francesinha["registros"]:
            avisos.append(
                "Nenhum recebimento foi lido na francesinha Bolepix. "
                "Confirme se o arquivo é o relatório "
                "Extrato/Devolução exportado pelo BRB."
            )

        if dados_extrato is not None:

            diferenca = (
                dados_extrato["total_pix"] - dados_francesinha["total"]
            )

            dados_francesinha["conferencia_extrato"] = {
                "total_extrato": dados_extrato["total_pix"],
                "total_francesinha": dados_francesinha["total"],
                "diferenca": diferenca,
                "confere": diferenca == ZERO,
            }

            if diferenca != ZERO:
                avisos.append(
                    "O total da francesinha Bolepix não fecha com o "
                    "total de CRED PIX QR CODE DINAMICO do extrato "
                    f"(diferença de {formatar_moeda(diferenca)})."
                )

    # ---------- 3. francesinha bolebarras (opcional) ----------

    if bolebarras is not None:

        dados_bolebarras = ler_francesinha(bolebarras)
        resultado["bolebarras"] = dados_bolebarras

        if dados_extrato is not None:

            diferenca = (
                dados_extrato["total_cobranca"] - dados_bolebarras["total"]
            )

            dados_bolebarras["conferencia_extrato"] = {
                "total_extrato": dados_extrato["total_cobranca"],
                "total_francesinha": dados_bolebarras["total"],
                "diferenca": diferenca,
                "confere": diferenca == ZERO,
            }

            if diferenca != ZERO:
                avisos.append(
                    "O total da francesinha Bolebarras não fecha com o "
                    "total de CREDITO COBRANCA BRB do extrato "
                    f"(diferença de {formatar_moeda(diferenca)})."
                )

    # ---------- 4. DCB ----------

    dados_dcb = None

    if dcb is not None:

        dados_dcb = ler_dcb(dcb)
        resultado["dcb"] = dados_dcb

        if not dados_dcb["registros"]:
            avisos.append(
                "Nenhum título foi lido no DCB. Confirme se o arquivo "
                "é o retorno da cobrança (CNAB 400) do BRB."
            )

    # ---------- 5. PIX do extrato com nome do pagador ----------

    # A planilha manual preenche à mão as colunas Cliente /
    # CPF / Data-Hora do pagamento. Aqui isso é feito casando o
    # valor do crédito no extrato com o recebimento do Bolepix.
    if dados_extrato is not None and dados_francesinha is not None:

        casados, extrato_sem_pix, pix_sem_extrato = casar_por_valor(
            dados_extrato["pix"],
            dados_francesinha["registros"],
        )

        linhas = []

        for credito, recebimento in casados:
            linhas.append({
                "data": credito["data"],
                "descricao": credito["descricao"],
                "valor": credito["valor"],
                "nome": recebimento["nome"],
                "documento": mascarar_documento(
                    recebimento.get("documento")
                ),
                "data_pagamento": recebimento["data"],
                "identificado": True,
            })

        for credito in extrato_sem_pix:
            linhas.append({
                "data": credito["data"],
                "descricao": credito["descricao"],
                "valor": credito["valor"],
                "nome": "",
                "documento": "",
                "data_pagamento": "",
                "identificado": False,
            })

        resultado["pix_detalhado"] = {
            "linhas": linhas,
            "quantidade": len(linhas),
            "quantidade_identificados": len(casados),
            "quantidade_nao_identificados": len(extrato_sem_pix),
            "pix_sem_credito_no_extrato": pix_sem_extrato,
        }

    # ---------- 6. divergência ----------

    if dados_extrato is not None and dados_dcb is not None:

        total_pagamentos = dados_extrato["total_pagamentos"]
        total_retorno = dados_dcb["total_liquidado"]

        divergencia = total_pagamentos - total_retorno

        # Identifica, item a item, os PIX que entraram na conta
        # mas não têm título liquidado no arquivo de retorno.
        nao_baixados = []

        if dados_francesinha is not None:

            _, nao_baixados, _ = casar_por_valor(
                dados_francesinha["registros"],
                dados_dcb["liquidados"],
            )

        resultado["divergencia"] = {
            "total_pagamentos": total_pagamentos,
            "total_retorno": total_retorno,
            "divergencia": divergencia,

            "nao_baixados": nao_baixados,
            "quantidade_nao_baixados": len(nao_baixados),
            "total_nao_baixados": somar(nao_baixados),

            "explicada": (
                somar(nao_baixados) == divergencia
                if dados_francesinha is not None
                else None
            ),
        }

        if dados_francesinha is not None and not resultado[
            "divergencia"
        ]["explicada"]:
            avisos.append(
                "A divergência não foi totalmente explicada pelos "
                "Bolepix não baixados. Verifique também a cobrança "
                "por código de barras."
            )

    return resultado
