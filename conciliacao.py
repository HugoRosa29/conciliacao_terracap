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

from planilhas import (
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


def boleto_do_nosso_numero(nosso_numero):
    r"""
    Extrai do nosso número do BRB o número do boleto — que é o
    que o GAI imprime, nos relatórios dele, como "Nosso Número".

    O nosso número tem 12 dígitos:

        6 8 3 4 3 4 2 0 7 0 1 5
        |  \____ ____/  \_/ \_/
        |       |        |   |
        |       |        |   dígitos verificadores
        |       |        código do banco (070 = BRB)
        |       boleto, como o GAI chama de nosso número
        prefixo da carteira

    É por esse número que a cobrança da francesinha encontra o
    seu par no GAI, sem depender do valor. Conferido contra os
    arquivos de retorno reais (3.291 títulos, todos com 070 na
    mesma posição) e, item a item, contra a francesinha
    Bolebarra de 30/09/2026.
    """

    digitos = re.sub(r"\D", "", str(nosso_numero or ""))

    if len(digitos) != 12 or digitos[7:10] != "070":
        return ""

    return digitos[1:7]


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

# O Bolepix sai no relatório "Extrato/Devolução".
COLUNAS_FRANCESINHA = {
    "data": ("data e hora", "data/hora", "data"),
    "tipo": ("tipo transacao", "tipo"),
    "nome": ("nome contraparte", "contraparte", "nome", "cliente"),
    "documento": ("cnpj / cpf", "cnpj/cpf", "cpf/cnpj", "documento"),
    "valor": ("valor",),
}

# A Bolebarra sai no relatório "Francesinha movimento", que é
# título por título: traz o nosso número, que amarra cada
# recebimento ao DCB e ao GAI sem depender do valor.
COLUNAS_MOVIMENTO = {
    "nosso_numero": ("nosso numero",),
    "documento": ("nº documento", "n° documento", "no documento"),
    "vencimento": ("vencimento",),
    "valor": ("valor",),
    "situacao": ("situacao",),
    "data": ("dt. liquid.", "dt liquid", "data liquidacao"),
    "valor_liquidado": ("vl. liquid.", "vl liquid", "valor liquidacao"),
    "nome": ("sacado", "nome contraparte", "contraparte", "nome", "cliente"),
}

TIPOS_RECEBIMENTO = ("RECEBIMENTO", "CREDITO")

# Layouts conhecidos, na ordem em que são tentados. Cada um traz
# o mapa de colunas e os campos que precisam existir para o
# cabeçalho ser aceito.
LAYOUTS_FRANCESINHA = (
    ("movimento", COLUNAS_MOVIMENTO, ("nosso_numero", "valor")),
    ("extrato", COLUNAS_FRANCESINHA, ("nome", "valor")),
)


def _localizar_cabecalho(linhas):
    """
    Encontra a linha de cabeçalho do relatório e devolve
    (layout, índice da linha, {campo: índice da coluna}).
    """

    for indice_linha, linha in enumerate(linhas[:40]):

        celulas = {}

        for indice_coluna, valor in enumerate(linha):

            texto = normalizar_nome(texto_da_celula(valor)).lower()

            if texto:
                celulas[indice_coluna] = texto

        if not celulas:
            continue

        for layout, colunas, obrigatorios in LAYOUTS_FRANCESINHA:

            mapa = {}

            for campo, apelidos in colunas.items():

                esperados = [normalizar_nome(a).lower() for a in apelidos]

                for indice_coluna, texto in celulas.items():
                    if texto in esperados:
                        mapa[campo] = indice_coluna
                        break

            if all(campo in mapa for campo in obrigatorios):
                return layout, indice_linha, mapa

    return "", None, {}


def _ler_francesinha_planilha(origem, nome=""):
    """
    Lê a francesinha exportada pelo BRB, nos dois formatos:

        Extrato/Devolução      Bolepix, por recebimento PIX
        Francesinha movimento  Bolebarra, título por título
    """

    abas = ler_planilha(origem, nome=nome)

    registros = []
    layouts = set()

    for nome_aba, linhas in abas.items():

        layout, indice_cabecalho, mapa = _localizar_cabecalho(linhas)

        if not mapa:
            continue

        layouts.add(layout)

        ler_linha = (
            _linha_do_movimento if layout == "movimento" else _linha_do_extrato
        )

        for indice_linha in range(indice_cabecalho + 1, len(linhas)):

            registro = ler_linha(linhas[indice_linha], mapa)

            if registro is None:
                continue

            registro["linha"] = indice_linha + 1
            registro["aba"] = nome_aba

            registros.append(registro)

    return registros, ("movimento" if "movimento" in layouts else "extrato")


def _linha_do_extrato(linha, mapa):
    """
    Uma linha do relatório Extrato/Devolução (Bolepix).
    """

    bruto = celula(linha, mapa["valor"])

    if bruto is None:
        return None

    tipo = texto_da_celula(celula(linha, mapa.get("tipo")))
    nome_pagador = texto_da_celula(celula(linha, mapa["nome"]))

    # A última linha do relatório é o total: tem valor, mas não
    # tem contraparte nem tipo de transação.
    if not nome_pagador and not tipo:
        return None

    if tipo and normalizar_nome(tipo) not in TIPOS_RECEBIMENTO:
        return None

    valor = moeda_para_decimal(bruto)

    if valor == ZERO:
        return None

    return {
        "data": texto_da_celula(celula(linha, mapa.get("data"))),
        "tipo": tipo,
        "nome": nome_pagador,
        "nome_normalizado": normalizar_nome(nome_pagador),
        "documento": texto_da_celula(celula(linha, mapa.get("documento"))),
        "valor": valor,
        "nosso_numero": "",
        "boleto": "",
    }


def _linha_do_movimento(linha, mapa):
    """
    Uma linha do relatório Francesinha movimento (Bolebarra).

    O nosso número é o crivo: as linhas de total e as anotações
    feitas à mão depois da última cobrança não têm um, e ficam
    de fora.
    """

    nosso_numero = re.sub(
        r"\D", "", texto_da_celula(celula(linha, mapa["nosso_numero"]))
    )

    if len(nosso_numero) != 12:
        return None

    valor = moeda_para_decimal(celula(linha, mapa["valor"]))

    if valor == ZERO:
        return None

    nome_sacado = texto_da_celula(celula(linha, mapa.get("nome")))

    documento = texto_da_celula(celula(linha, mapa.get("documento")))

    return {
        "data": _somente_data(
            texto_da_celula(celula(linha, mapa.get("data")))
        ),
        "tipo": texto_da_celula(celula(linha, mapa.get("situacao"))),
        "nome": nome_sacado,
        "nome_normalizado": normalizar_nome(nome_sacado),
        # No movimento da Bolebarra esta coluna traz o número da
        # alienação, que é o "seu número" do DCB.
        "documento": "" if documento in ("0", "") else documento,
        "valor": valor,
        "vencimento": _somente_data(
            texto_da_celula(celula(linha, mapa.get("vencimento")))
        ),
        "valor_liquidado": moeda_para_decimal(
            celula(linha, mapa.get("valor_liquidado"))
        ),
        "nosso_numero": nosso_numero,
        "boleto": boleto_do_nosso_numero(nosso_numero),
    }


def _somente_data(texto):
    """
    "30/09/2026 00:00:00" -> "30/09/2026"
    """

    return (texto or "").split(" ")[0]


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
            "nosso_numero": "",
            "boleto": "",
        })

    return registros, ignoradas


def ler_francesinha(origem, delimitador=None):
    """
    Lê a francesinha em planilha (relatório do BRB) ou em texto
    delimitado, detectando o formato automaticamente.
    """

    nome = nome_arquivo(origem, padrao="francesinha.xls")

    ignoradas = 0
    layout = ""

    if e_planilha(origem, nome):
        registros, layout = _ler_francesinha_planilha(origem, nome=nome)
        formato = "planilha"
    else:
        registros, ignoradas = _ler_francesinha_texto(
            para_texto(origem),
            delimitador=delimitador,
        )
        formato = "texto"

    com_boleto = [r for r in registros if r.get("boleto")]

    return {
        "formato": formato,
        "layout": layout,
        "registros": registros,
        "quantidade": len(registros),
        "linhas_ignoradas": ignoradas,
        "total": somar(registros),

        # Quando a francesinha traz o nosso número, a
        # conciliação deixa de depender do valor e passa a casar
        # título por título.
        "quantidade_com_boleto": len(com_boleto),
        "tem_nosso_numero": len(com_boleto) == len(registros) and bool(
            registros
        ),
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
# 5. BOLEBARRA x DCB
# ============================================================

def _indice_por_chave(registros, chave):
    """
    Agrupa registros por uma chave de texto, preservando a
    ordem de cada grupo.
    """

    indice = {}

    for registro in registros:

        valor = registro.get(chave)

        if valor:
            indice.setdefault(valor, []).append(registro)

    return indice


def conferir_bolebarras_dcb(bolebarras, dcb):
    """
    Confere a francesinha Bolebarra contra o arquivo de retorno,
    título por título, pelo nosso número.

    É uma conferência mais firme que a do Bolepix: o Bolepix só
    pode ser casado pelo valor, porque o relatório de PIX não
    traz o nosso número. Aqui cada cobrança da francesinha tem
    um, e ou ela aparece liquidada no DCB com o mesmo valor, ou
    a diferença é apontada.

    Devolve os quatro desencontros possíveis:

        valor_divergente   está nos dois, com valores diferentes
        sem_baixa          está na francesinha, não liquidou no DCB
        fora_da_francesinha  liquidou no DCB, não está na francesinha
                             (normalmente são os Bolepix do dia)
        sem_nosso_numero   linha da francesinha sem nosso número
    """

    titulos = bolebarras["registros"]

    # Um mesmo nosso número aparece mais de uma vez no retorno:
    # uma na entrada do título e outra na liquidação. Vale a
    # liquidação.
    por_nosso_numero = _indice_por_chave(dcb["registros"], "nosso_numero")

    casados = []
    valor_divergente = []
    sem_baixa = []
    sem_nosso_numero = []

    usados = set()

    for titulo in titulos:

        nosso_numero = titulo.get("nosso_numero")

        if not nosso_numero:
            sem_nosso_numero.append(titulo)
            continue

        encontrados = por_nosso_numero.get(nosso_numero, [])

        liquidados = [r for r in encontrados if r["liquidado"]]

        if not liquidados:
            sem_baixa.append({
                "titulo": titulo,
                "no_arquivo": bool(encontrados),
                "ocorrencias": [r["ocorrencia"] for r in encontrados],
            })
            continue

        registro = liquidados[0]

        usados.add(nosso_numero)

        par = {
            "titulo": titulo,
            "registro": registro,
            "diferenca": titulo["valor"] - registro["valor"],
        }

        if par["diferenca"] == ZERO:
            casados.append(par)
        else:
            valor_divergente.append(par)

    fora_da_francesinha = [
        registro
        for registro in dcb["liquidados"]
        if registro["nosso_numero"] not in usados
    ]

    total_conferido = sum(
        (par["titulo"]["valor"] for par in casados), ZERO
    )

    return {
        "casados": casados,
        "valor_divergente": valor_divergente,
        "sem_baixa": sem_baixa,
        "sem_nosso_numero": sem_nosso_numero,
        "fora_da_francesinha": fora_da_francesinha,

        "quantidade_casados": len(casados),
        "quantidade_valor_divergente": len(valor_divergente),
        "quantidade_sem_baixa": len(sem_baixa),
        "quantidade_sem_nosso_numero": len(sem_nosso_numero),
        "quantidade_fora_da_francesinha": len(fora_da_francesinha),

        "total_francesinha": bolebarras["total"],
        "total_conferido": total_conferido,
        "total_sem_baixa": somar(
            [item["titulo"] for item in sem_baixa]
        ),
        "total_fora_da_francesinha": somar(fora_da_francesinha),

        "confere": not (
            valor_divergente or sem_baixa or sem_nosso_numero
        ),
    }


# ============================================================
# 6. QUEM É O PAGADOR DE UM REGISTRO DO GAI
# ============================================================

def _nome_comum(candidatos):
    """
    Escolhe o nome entre as cobranças de uma mesma alienação.

    Nomes iguais, um nome só. Nomes que diferem porque o
    relatório corta o texto em larguras diferentes contam como o
    mesmo nome, e vale o mais completo. Nomes de verdade
    diferentes: nenhum deles, para não pendurar um pagamento em
    quem não o fez.
    """

    nomes = sorted(
        {c["nome"] for c in candidatos if c.get("nome")},
        key=len,
        reverse=True,
    )

    if not nomes:
        return ""

    maior = normalizar_nome(nomes[0])

    if all(maior.startswith(normalizar_nome(nome)) for nome in nomes):
        return nomes[0]

    return ""


def identificar_recusados(nao_efetivadas, bolebarras):
    """
    Põe o nome do sacado em cada registro recusado pelo GAI.

    O relatório de Baixas Não Efetivadas não traz nome: traz a
    alienação, o boleto e a parcela. O nome vem da francesinha
    Bolebarra, procurado primeiro pelo boleto e, se ali não
    estiver, pela alienação — que é a coluna "Nº Documento" da
    francesinha.

    A busca pela alienação é o que resolve a segunda via: ela é
    recusada com um número de boleto novo, que não existe na
    francesinha. Foi assim que o pagamento de R$ 1.921,07 de
    30/09/2026 ganhou dono — o boleto 827696 não está na
    francesinha, mas a alienação 111268 está, paga pelo boleto
    825978, de EDLEUZA GONCALVES DOS REIS.

    Quem não é encontrado fica sem nome, e com razão: a maior
    parte dos registros recusados é entrada de título, boleto
    que ninguém pagou no dia.
    """

    por_boleto = {}
    por_alienacao = {}

    for titulo in bolebarras["registros"]:

        if titulo.get("boleto"):
            por_boleto.setdefault(titulo["boleto"], titulo)

        if titulo.get("documento"):
            por_alienacao.setdefault(titulo["documento"], []).append(titulo)

    identificados = 0

    for registro in nao_efetivadas["registros"]:

        titulo = por_boleto.get(registro["boleto"])

        if titulo is not None:
            registro["nome"] = titulo["nome"]
            registro["nosso_numero"] = titulo["nosso_numero"]
            registro["origem_do_nome"] = "boleto"

        else:
            registro["nome"] = _nome_comum(
                por_alienacao.get(registro["alienacao"], [])
            )
            registro["nosso_numero"] = ""
            registro["origem_do_nome"] = (
                "alienacao" if registro["nome"] else ""
            )

        if registro["nome"]:
            identificados += 1

    nao_efetivadas["quantidade_identificados"] = identificados

    return identificados


# ============================================================
# 7. BOLEBARRA x RELATÓRIOS DO GAI
# ============================================================

# Por que o nosso número e não o valor: a Relação de Parcelas
# Lidas quebra um boleto nas parcelas que ele pagou, então o
# valor de uma linha do relatório quase nunca é o valor da
# cobrança. Somadas por boleto, as parcelas fecham com o valor
# da francesinha — e o boleto é justamente o que o GAI imprime
# como nosso número.

SITUACAO_BAIXADO = "baixado"
SITUACAO_LIDO_SEM_BAIXA = "lido_sem_baixa"
SITUACAO_RECUSADO = "recusado"
SITUACAO_AUSENTE = "ausente"

EXPLICACAO_SITUACAO = {
    SITUACAO_LIDO_SEM_BAIXA: (
        "O GAI leu o pagamento e registrou valor 0,00 na Relação de "
        "Parcelas Lidas: a baixa não foi aplicada."
    ),
    SITUACAO_RECUSADO: (
        "O GAI recusou a baixa e informou o motivo no relatório de "
        "Baixas de Pagamentos Não Efetivadas."
    ),
    SITUACAO_AUSENTE: (
        "A cobrança não aparece em nenhum dos dois relatórios do GAI. "
        "Em geral é cobrança de outra gerência (GIR, GGR, GOP), que "
        "não passa pelo GAI."
    ),
}


def conferir_bolebarras_gai(bolebarras, lidas=None, nao_efetivadas=None):
    """
    Confere a francesinha Bolebarra contra os relatórios do GAI
    e diz, cobrança por cobrança, o que o GAI fez com ela.

    Cada cobrança cai em uma de quatro situações:

        baixado          o GAI baixou, e o valor fecha
        lido_sem_baixa   o GAI leu e registrou 0,00
        recusado         o GAI recusou, com motivo
        ausente          não está em nenhum relatório do GAI

    Aponta também o caminho inverso: os boletos que o GAI baixou
    e não estão na francesinha Bolebarra — que são, no dia a
    dia, os recebimentos do Bolepix.
    """

    titulos = bolebarras["registros"]

    boletos_baixados = {}
    lidos_sem_baixa = {}

    if lidas is not None:
        boletos_baixados = {
            grupo["boleto"]: grupo for grupo in lidas["boletos"]
        }

        for registro in lidas["lidos_sem_baixa"]:
            lidos_sem_baixa.setdefault(registro["boleto"], registro)

    recusados = {}

    if nao_efetivadas is not None:
        for registro in nao_efetivadas["registros"]:
            recusados.setdefault(registro["boleto"], registro)

    linhas = []
    usados = set()

    for titulo in titulos:

        boleto = titulo.get("boleto")

        grupo = boletos_baixados.get(boleto) if boleto else None

        if grupo is not None:
            usados.add(boleto)

            linhas.append({
                "titulo": titulo,
                "situacao": SITUACAO_BAIXADO,
                "valor_gai": grupo["valor"],
                "diferenca": titulo["valor"] - grupo["valor"],
                "alienacao": grupo["alienacao"],
                "quantidade_parcelas": len(grupo["parcelas"]),
                "motivo": "",
            })

            continue

        recusa = recusados.get(boleto) if boleto else None

        if boleto and boleto in lidos_sem_baixa:
            situacao = SITUACAO_LIDO_SEM_BAIXA
        elif recusa is not None:
            situacao = SITUACAO_RECUSADO
        else:
            situacao = SITUACAO_AUSENTE

        linhas.append({
            "titulo": titulo,
            "situacao": situacao,
            "valor_gai": ZERO,
            "diferenca": titulo["valor"],
            "alienacao": (
                (recusa or lidos_sem_baixa.get(boleto) or {}).get("alienacao")
                or titulo.get("documento")
                or ""
            ),
            "quantidade_parcelas": 0,
            "motivo": (recusa or {}).get("motivo", ""),
        })

    def por_situacao(situacao):
        return [linha for linha in linhas if linha["situacao"] == situacao]

    baixados = por_situacao(SITUACAO_BAIXADO)

    # Um boleto que o GAI baixou por um valor diferente do que a
    # francesinha recebeu é um caso à parte: o dinheiro entrou,
    # mas a baixa saiu torta.
    valor_divergente = [
        linha for linha in baixados if linha["diferenca"] != ZERO
    ]

    nao_baixados = [
        linha for linha in linhas if linha["situacao"] != SITUACAO_BAIXADO
    ]

    fora_da_francesinha = [
        grupo
        for grupo in boletos_baixados.values()
        if grupo["boleto"] not in usados
    ]

    def total(colecao, chave="valor"):
        return sum((item[chave] for item in colecao), ZERO)

    total_baixado = sum(
        (linha["titulo"]["valor"] for linha in baixados), ZERO
    )

    return {
        "linhas": linhas,
        "baixados": baixados,
        "valor_divergente": valor_divergente,
        "nao_baixados": nao_baixados,
        "fora_da_francesinha": fora_da_francesinha,

        "quantidade_baixados": len(baixados),
        "quantidade_valor_divergente": len(valor_divergente),
        "quantidade_nao_baixados": len(nao_baixados),
        "quantidade_fora_da_francesinha": len(fora_da_francesinha),

        "total_francesinha": bolebarras["total"],
        "total_baixado": total_baixado,
        "total_nao_baixado": sum(
            (linha["titulo"]["valor"] for linha in nao_baixados), ZERO
        ),
        "total_fora_da_francesinha": total(fora_da_francesinha),

        "situacoes": [
            {
                "situacao": situacao,
                "quantidade": len(por_situacao(situacao)),
                "total": sum(
                    (
                        linha["titulo"]["valor"]
                        for linha in por_situacao(situacao)
                    ),
                    ZERO,
                ),
                "explicacao": EXPLICACAO_SITUACAO.get(situacao, ""),
            }
            for situacao in (
                SITUACAO_LIDO_SEM_BAIXA,
                SITUACAO_RECUSADO,
                SITUACAO_AUSENTE,
            )
            if por_situacao(situacao)
        ],

        "confere": not nao_baixados and not valor_divergente,
    }


# ============================================================
# 8. CONCILIAÇÃO COMPLETA
# ============================================================

def _ler_relatorios_gai(lidas, nao_baixadas, avisos):
    """
    Lê os dois PDFs do GAI, sem exigir que venham no campo
    certo: cada relatório se identifica pelo título impresso na
    primeira página, então trocar um pelo outro na tela não
    muda o resultado.
    """

    import gai

    encontrados = {}

    for campo, origem in (
        ("Parcelas lidas", lidas),
        ("Baixas não efetivadas", nao_baixadas),
    ):
        if origem is None:
            continue

        if not gai.e_pdf(origem):
            avisos.append(
                f"O arquivo enviado no campo {campo} do GAI não é um "
                "PDF. Os dois relatórios do GAI são PDF."
            )
            continue

        try:
            relatorio = gai.ler_relatorio_gai(origem)
        except Exception as erro:
            avisos.append(
                f"Não foi possível ler o PDF enviado no campo {campo} "
                f"do GAI: {erro}"
            )
            continue

        if relatorio["tipo"] in encontrados:
            avisos.append(
                "Os dois PDFs enviados são o mesmo relatório do GAI; "
                "o segundo foi ignorado."
            )
            continue

        encontrados[relatorio["tipo"]] = relatorio

    return (
        encontrados.get("parcelas_lidas"),
        encontrados.get("baixas_nao_efetivadas"),
    )

def conciliar(
    extrato=None,
    dcb=None,
    francesinha=None,
    bolebarras=None,
    gai_lidas=None,
    gai_nao_baixadas=None,
    gir=None,
    ggr=None,
    gop=None,
    benner=None,
):
    """
    Executa a conciliação com os arquivos disponíveis.

    Todos os arquivos são opcionais: o que for enviado é
    processado, o que faltar é omitido do resultado. Em
    particular, a conciliação da Bolebarra com o DCB e com o GAI
    não depende do extrato — o extrato serve para fechar os
    totais do dia, não para casar título por título.

    Devolve um dicionário com os blocos:

        extrato         totais do extrato BRB
        francesinha     detalhe dos PIX recebidos (Bolepix)
        bolebarras      detalhe da cobrança por código de barras
        dcb             títulos do arquivo de retorno
        gai_lidas       Relação de Parcelas Lidas do GAI
        gai_nao_baixadas  Baixas Não Efetivadas do GAI
        pix_detalhado   PIX do extrato com o nome do pagador
        bolepix_dcb     Bolepix sem baixa no arquivo de retorno
        bolebarras_dcb  Bolebarra conferida com o DCB
        bolebarras_gai  Bolebarra conferida com os relatórios do GAI
        divergencia     Total pagamentos - Total retorno
        avisos          pontos de atenção para o usuário
    """

    avisos = []

    resultado = {
        "extrato": None,
        "francesinha": None,
        "bolebarras": None,
        "dcb": None,
        "gai_lidas": None,
        "gai_nao_baixadas": None,
        "pix_detalhado": None,
        "bolepix_dcb": None,
        "bolepix_gai": None,
        "bolebarras_dcb": None,
        "bolebarras_gai": None,
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

    dados_bolebarras = None

    if bolebarras is not None:

        dados_bolebarras = ler_francesinha(bolebarras)
        resultado["bolebarras"] = dados_bolebarras

        if not dados_bolebarras["registros"]:
            avisos.append(
                "Nenhuma cobrança foi lida na francesinha Bolebarra. "
                "Confirme se o arquivo é o relatório Francesinha "
                "movimento exportado pelo BRB."
            )

        elif not dados_bolebarras["tem_nosso_numero"]:
            avisos.append(
                f"{dados_bolebarras['quantidade'] - dados_bolebarras['quantidade_com_boleto']}"
                " cobrança(s) da francesinha Bolebarra estão sem nosso "
                "número: essas não podem ser casadas com o DCB nem com "
                "o GAI título por título."
            )

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

    # ---------- 5. relatórios do GAI ----------

    dados_lidas, dados_nao_baixadas = _ler_relatorios_gai(
        gai_lidas, gai_nao_baixadas, avisos
    )

    resultado["gai_lidas"] = dados_lidas
    resultado["gai_nao_baixadas"] = dados_nao_baixadas

    for relatorio in (dados_lidas, dados_nao_baixadas):
        if relatorio is not None:
            avisos.extend(relatorio["avisos"])

    # O relatório de recusas vem só com códigos. A francesinha
    # Bolebarra dá nome a cada um que ela conhece.
    if dados_nao_baixadas is not None and dados_bolebarras is not None:
        identificar_recusados(dados_nao_baixadas, dados_bolebarras)

    if dados_lidas is not None and dados_dcb is not None:

        esperado = (dados_lidas["arquivo"] or "").upper()
        recebido = (nome_arquivo(dcb, padrao="") or "").upper()

        # O relatório do GAI diz de qual arquivo de retorno ele
        # saiu. Conferir evita conciliar o dia errado.
        if esperado and recebido and not recebido.endswith(esperado):
            avisos.append(
                "A Relação de Parcelas Lidas do GAI foi gerada a partir "
                f"de {dados_lidas['arquivo']}, que não é o DCB enviado. "
                "Confirme se os arquivos são do mesmo dia."
            )

    # ---------- 6. PIX do extrato com nome do pagador ----------

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

    # ---------- 7. Bolepix x DCB ----------

    # Quais PIX recebidos não têm título liquidado no arquivo de
    # retorno. O Bolepix só pode ser casado pelo valor: o
    # relatório de PIX do banco não traz nosso número.
    #
    # Não depende do extrato — o extrato só entra depois, para
    # fechar o total do dia.
    nao_baixados = []

    if dados_francesinha is not None and dados_dcb is not None:

        _, nao_baixados, _ = casar_por_valor(
            dados_francesinha["registros"],
            dados_dcb["liquidados"],
        )

        resultado["bolepix_dcb"] = {
            "total_francesinha": dados_francesinha["total"],
            "nao_baixados": nao_baixados,
            "quantidade_nao_baixados": len(nao_baixados),
            "total_nao_baixados": somar(nao_baixados),
            "confere": not nao_baixados,
        }

    # ---------- 8. Bolebarra x DCB ----------

    if dados_bolebarras is not None and dados_dcb is not None:

        conferencia = conferir_bolebarras_dcb(dados_bolebarras, dados_dcb)

        resultado["bolebarras_dcb"] = conferencia

        if conferencia["valor_divergente"]:
            avisos.append(
                f"{conferencia['quantidade_valor_divergente']} cobrança(s) "
                "da francesinha Bolebarra estão no arquivo de retorno com "
                "valor diferente."
            )

        if conferencia["sem_baixa"]:
            avisos.append(
                f"{conferencia['quantidade_sem_baixa']} cobrança(s) da "
                "francesinha Bolebarra não têm título liquidado no "
                "arquivo de retorno "
                f"({formatar_moeda(conferencia['total_sem_baixa'])})."
            )

    # ---------- 9. Bolebarra x GAI ----------

    if dados_bolebarras is not None and (
        dados_lidas is not None or dados_nao_baixadas is not None
    ):
        conferencia = conferir_bolebarras_gai(
            dados_bolebarras,
            lidas=dados_lidas,
            nao_efetivadas=dados_nao_baixadas,
        )

        resultado["bolebarras_gai"] = conferencia

        if dados_lidas is None:
            avisos.append(
                "Sem a Relação de Parcelas Lidas do GAI não é possível "
                "dizer o que o GAI baixou; envie também esse PDF."
            )

        elif conferencia["nao_baixados"]:
            avisos.append(
                f"{conferencia['quantidade_nao_baixados']} cobrança(s) da "
                "francesinha Bolebarra não foram baixadas no GAI "
                f"({formatar_moeda(conferencia['total_nao_baixado'])})."
            )

        if conferencia["valor_divergente"]:
            avisos.append(
                f"{conferencia['quantidade_valor_divergente']} cobrança(s) "
                "foram baixadas no GAI por valor diferente do recebido."
            )

    # ---------- 10. Bolepix x GAI ----------

    # Os boletos que o GAI baixou e não estão na francesinha
    # Bolebarra são, no dia a dia, os recebimentos do Bolepix.
    # Comparados com a francesinha Bolepix, mostram quanto do
    # PIX do dia o GAI deixou de baixar.
    if (
        resultado["bolebarras_gai"] is not None
        and dados_francesinha is not None
        and dados_lidas is not None
    ):
        do_gai = [
            dict(grupo, nome="", data="")
            for grupo in resultado["bolebarras_gai"]["fora_da_francesinha"]
        ]

        _, pix_sem_gai, gai_sem_pix = casar_por_valor(
            dados_francesinha["registros"],
            do_gai,
        )

        resultado["bolepix_gai"] = {
            "total_francesinha": dados_francesinha["total"],
            "total_no_gai": somar(do_gai),
            "nao_baixados": pix_sem_gai,
            "quantidade_nao_baixados": len(pix_sem_gai),
            "total_nao_baixados": somar(pix_sem_gai),
            "boletos_sem_pix": gai_sem_pix,
            "quantidade_boletos_sem_pix": len(gai_sem_pix),
            "confere": not pix_sem_gai and not gai_sem_pix,
        }

    # ---------- 11. divergência do dia ----------

    if dados_extrato is not None and dados_dcb is not None:

        total_pagamentos = dados_extrato["total_pagamentos"]
        total_retorno = dados_dcb["total_liquidado"]

        divergencia = total_pagamentos - total_retorno

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

    from sistemas import ler_sistema, cruzar_sistemas

    relatorios = [
        ler_sistema(origem, esperado=sistema)
        for sistema, origem in (("GIR", gir), ("GGR", ggr), ("GOP", gop))
        if origem is not None
    ]
    resultado["sistemas"] = relatorios
    resultado["bolebarras_sistemas"] = (
        cruzar_sistemas(relatorios, dados_dcb, dados_bolebarras, dados_lidas)
        if relatorios else None
    )
    for relatorio in relatorios:
        avisos.extend(relatorio["avisos"])

    from benner import ler_benner, conferir_benner

    resultado["benner"] = ler_benner(benner) if benner is not None else None
    resultado["gai_benner"] = None
    if resultado["benner"] is not None:
        resultado["gai_benner"] = conferir_benner(dados_lidas, resultado["benner"])
        avisos.extend(resultado["benner"]["avisos"])
        avisos.extend(resultado["gai_benner"]["avisos"])

    return resultado
