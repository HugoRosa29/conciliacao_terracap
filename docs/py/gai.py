"""
Leitura dos relatórios do GAI em PDF.

O GAI (Gerenciamento de Alienações de Imóveis) processa o mesmo
arquivo de retorno que o banco manda no DCB e devolve dois
relatórios, que no dia a dia são salvos com o valor no nome:

    Relação de Parcelas Lidas
        o que o GAI conseguiu baixar
        ("VALORES BAIXADOS NO GAI (AG. 121) R$ ...")

    Baixas de Pagamentos Não Efetivadas
        o que o GAI leu e recusou, com o motivo
        ("VALORES NÃO BAIXADOS NO GAI (AG. 121) R$ ...")

Os dois são relatórios de texto posicionado: cada campo é
desenhado numa coordenada fixa da página. A leitura, então,
agrupa os pedaços de texto por linha (coordenada vertical) e
decide a coluna de cada um pela coordenada horizontal — as
faixas estão em COLUNAS_LIDAS e COLUNAS_NAO_EFETIVADAS, e foram
conferidas contra todas as páginas dos relatórios reais de
30/09/2026.

Usa pypdf, que é Python puro, para o mesmo código rodar no CLI,
no servidor e no navegador via Pyodide.
"""

from decimal import Decimal
import logging
import re

from planilhas import e_caminho, nome_arquivo, para_bytes


ZERO = Decimal("0.00")

ASSINATURA_PDF = b"%PDF"


# O pypdf avisa, a cada fonte Type1 do relatório, que o
# fontTools não está instalado. Ele só precisaria do fontTools
# para desenhar a fonte; para ler o texto, não. O aviso seria
# centenas de linhas de ruído no terminal e no console do
# navegador, então fica silenciado.
logging.getLogger("pypdf").setLevel(logging.ERROR)


# ============================================================
# LAYOUT DOS RELATÓRIOS
#
# Faixas horizontais (em pontos, como o PDF guarda) de cada
# coluna. O começo de um campo varia dentro da faixa porque os
# números são alinhados à direita: "2.573,02" começa em 484 e
# "1.095.559,00" começa em 468, e os dois são a coluna Valor.
# ============================================================

# Relação de Parcelas Lidas
#
# Alienação  Imóvel  Modalidade  Parcela  Data Pagamento
# Data Crédito  Código Baixa  Nosso Número  Valor  IPTU
COLUNAS_LIDAS = (
    ("alienacao", 0, 55),
    ("imovel", 55, 110),
    ("modalidade", 110, 160),
    ("parcela", 160, 197),
    ("data_pagamento", 197, 270),
    ("data_credito", 270, 330),
    ("codigo_baixa", 330, 390),
    ("nosso_numero", 390, 460),
    ("valor", 460, 530),
    ("iptu", 530, 10000),
)

# Baixas de Pagamentos Não Efetivadas, primeira linha do
# registro:
#
# Alienação  Imóvel  Banco  Agência  Nr. Pag.  Tipo  Parcela
# Data Pagamento  Data Vencimento
#
# A coluna impressa como "Imóvel" traz, neste relatório, o
# número do boleto — é por ele que o pagamento é reconhecido.
COLUNAS_NAO_EFETIVADAS = (
    ("alienacao", 0, 60),
    ("boleto", 60, 135),
    ("banco", 135, 175),
    ("agencia", 175, 240),
    ("numero_pagamento", 240, 285),
    ("tipo", 285, 340),
    ("parcela", 340, 400),
    ("data_pagamento", 400, 480),
    ("data_vencimento", 480, 10000),
)

# Baixas de Pagamentos Não Efetivadas, segunda linha do
# registro:
#
# Parcela  Multa  Mora  Correção  Isenção  Total Devido
# Total Pago  Diferença
COLUNAS_VALORES = (
    ("parcela_valor", 0, 70),
    ("multa", 70, 170),
    ("mora", 170, 215),
    ("correcao", 215, 268),
    ("isencao", 268, 320),
    ("total_devido", 320, 410),
    ("total_pago", 410, 480),
    ("diferenca", 480, 10000),
)

# Distância vertical máxima, em pontos, entre pedaços de texto
# da mesma linha. O relatório desenha os campos de uma linha com
# meio ponto de diferença entre eles e deixa doze pontos até a
# linha seguinte.
TOLERANCIA_LINHA = 4.0

# "99464 - 0" -> alienação 99464, dígito verificador 0
CODIGO_COM_DIGITO = re.compile(r"(\d+)\s*-\s*(\d)")

# Títulos que identificam cada relatório.
TITULO_LIDAS = "relacao de parcelas lidas"
TITULO_NAO_EFETIVADAS = "baixas de pagamentos nao efetivadas"


# ============================================================
# APOIO
# ============================================================

def e_pdf(origem, nome=""):
    """
    Detecta um PDF pela extensão ou pela assinatura do arquivo.
    """

    if (nome or nome_arquivo(origem, padrao="")).lower().endswith(".pdf"):
        return True

    return para_bytes(origem)[:4] == ASSINATURA_PDF


def _sem_acento(texto):
    """
    Deixa o texto comparável: minúsculas, sem acento e sem
    espaços repetidos.
    """

    substituicoes = {
        "á": "a", "à": "a", "â": "a", "ã": "a", "ä": "a",
        "é": "e", "è": "e", "ê": "e", "ë": "e",
        "í": "i", "ì": "i", "î": "i", "ï": "i",
        "ó": "o", "ò": "o", "ô": "o", "õ": "o", "ö": "o",
        "ú": "u", "ù": "u", "û": "u", "ü": "u",
        "ç": "c", "ñ": "n",
    }

    texto = str(texto or "").lower()

    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)

    return re.sub(r"\s+", " ", texto).strip()


def _valor_brasileiro(texto):
    """
    Converte "1.095.559,00" em Decimal. Devolve None quando o
    texto não é um valor, para distinguir "campo vazio" de
    "campo igual a zero" — a diferença entre um boleto que o
    GAI não listou e um que ele listou com 0,00.
    """

    texto = str(texto or "").strip()

    if not re.fullmatch(r"-?[\d\.]*\d,\d{2}", texto):
        return None

    try:
        return Decimal(texto.replace(".", "").replace(",", "."))
    except Exception:
        return None


def _numero(texto):
    """
    Lê a parte numérica de um campo "99464 - 0", sem o dígito
    verificador, e devolve "" quando o campo não é desse tipo.
    """

    encontrado = CODIGO_COM_DIGITO.fullmatch(str(texto or "").strip())

    return encontrado.group(1) if encontrado else ""


def _coluna(itens, colunas):
    """
    Distribui os pedaços de texto de uma linha nas colunas,
    pela coordenada horizontal.

    Devolve ({campo: texto}, repetidos), onde repetidos lista os
    campos que receberam mais de um pedaço — sinal de que o
    layout do relatório mudou e as faixas precisam de revisão.
    """

    campos = {}
    repetidos = []

    for x, texto in itens:

        for nome, inicio, fim in colunas:

            if not inicio <= x < fim:
                continue

            if nome in campos:
                campos[nome] = f"{campos[nome]} {texto}"
                repetidos.append(nome)
            else:
                campos[nome] = texto

            break

    return campos, repetidos


# Dois ou mais espaços separam campos dentro de um pedaço de
# texto que veio colado (ver _partes).
ESPACO_LARGO = re.compile(r"\s{2,}")


def _partes(texto):
    """
    Separa os campos de um pedaço de texto que veio colado.

    O PDF desenha cada campo numa coordenada, mas de vez em
    quando manda vários deles numa tacada só, e aí o extrator
    entrega o conjunto como um pedaço único — com a coordenada
    do primeiro campo, não de cada um.
    """

    return [parte for parte in ESPACO_LARGO.split(texto.strip()) if parte]


def _linhas_pdf(origem):
    """
    Percorre o PDF e devolve, em ordem de leitura,

        (página, [(x, texto), ...])

    com um item por linha impressa. O pypdf entrega cada pedaço
    de texto junto da matriz que o posiciona na página; daí
    saem o x (coluna) e o y (linha).
    """

    from pypdf import PdfReader

    if e_caminho(origem):
        leitor = PdfReader(str(origem))
    else:
        from io import BytesIO

        leitor = PdfReader(BytesIO(para_bytes(origem)))

    for numero_pagina, pagina in enumerate(leitor.pages, start=1):

        pedacos = []

        def guardar(texto, matriz_corrente, matriz_texto, fonte, tamanho):
            texto = texto.strip()

            if texto:
                pedacos.append((matriz_texto[5], matriz_texto[4], texto))

        pagina.extract_text(visitor_text=guardar)

        # De cima para baixo e, dentro da linha, da esquerda
        # para a direita.
        pedacos.sort(key=lambda pedaco: (-pedaco[0], pedaco[1]))

        linhas = []

        for y, x, texto in pedacos:

            if linhas and linhas[-1][0] - y <= TOLERANCIA_LINHA:
                linhas[-1][1].append((x, texto))
            else:
                linhas.append((y, [(x, texto)]))

        for _, itens in linhas:
            yield numero_pagina, sorted(itens)


def _texto_da_linha(itens):
    """
    Junta a linha num texto só, para procurar títulos e totais.
    """

    return " ".join(texto for _, texto in itens)


# ============================================================
# IDENTIFICAÇÃO DO RELATÓRIO
# ============================================================

def identificar(origem):
    """
    Diz qual dos dois relatórios do GAI é o PDF:

        "parcelas_lidas" | "baixas_nao_efetivadas" | ""

    Serve para o usuário poder soltar os dois arquivos sem se
    preocupar com a ordem.
    """

    for numero_pagina, itens in _linhas_pdf(origem):

        texto = _sem_acento(_texto_da_linha(itens))

        if TITULO_LIDAS in texto:
            return "parcelas_lidas"

        if TITULO_NAO_EFETIVADAS in texto:
            return "baixas_nao_efetivadas"

        # O título está no cabeçalho da primeira página; não
        # vale varrer o relatório inteiro atrás dele.
        if numero_pagina > 1:
            break

    return ""


# ============================================================
# RELAÇÃO DE PARCELAS LIDAS  ("valores baixados no GAI")
# ============================================================

# O relatório tem duas espécies de linha:
#
#   parcela paga      valor maior que zero. Várias parcelas
#                     podem sair no mesmo boleto, e o número do
#                     boleto é impresso só na última delas.
#
#   parcela lida      valor 0,00 e nenhum número de boleto na
#   sem baixa         coluna própria: o boleto aparece na
#                     coluna "Imóvel". É o pagamento que o GAI
#                     leu do arquivo e não aplicou.
#
# O agrupamento das parcelas pagas por boleto foi conferido
# contra o rodapé do relatório real ("TOTAL DE BOLETOS
# BAIXADOS: 336") e contra a francesinha Bolebarra, título por
# título.

REGEX_TOTAL_PAGAMENTOS = re.compile(
    r"total de pagamentos:\s*(\d+)"
)
REGEX_TOTAL_RECEBIDO = re.compile(
    r"total geral recebido:\s*([\d\.]+,\d{2})"
)
REGEX_TOTAL_BOLETOS = re.compile(
    r"total de boletos baixados:\s*(\d+)"
)
REGEX_LOTE = re.compile(r"lote:\s*(\S+)")
REGEX_ARQUIVO = re.compile(r"arquivo:\s*(\S+)")


def ler_parcelas_lidas(origem):
    """
    Lê a "Relação de Parcelas Lidas" do GAI.

    Devolve as parcelas pagas agrupadas por boleto — que é a
    unidade que a francesinha do banco também usa — e, à parte,
    os boletos que o GAI leu sem dar baixa.
    """

    parcelas = []
    sem_baixa = []
    rodape = {}
    cabecalho = {}
    avisos = []

    for numero_pagina, itens in _linhas_pdf(origem):

        texto = _sem_acento(_texto_da_linha(itens))

        if "lote:" in texto and not cabecalho:
            lote = REGEX_LOTE.search(texto)
            arquivo = REGEX_ARQUIVO.search(texto)

            cabecalho = {
                "lote": lote.group(1) if lote else "",
                "arquivo": arquivo.group(1).upper() if arquivo else "",
            }

        for regex, chave, converter in (
            (REGEX_TOTAL_PAGAMENTOS, "total_pagamentos", int),
            (REGEX_TOTAL_BOLETOS, "total_boletos_baixados", int),
            (REGEX_TOTAL_RECEBIDO, "total_recebido", _valor_brasileiro),
        ):
            encontrado = regex.search(texto)

            # O "TOTAL GERAL RECEBIDO" aparece duas vezes: na
            # relação e, depois, no resumo por conta contábil,
            # que não soma o mesmo. Vale o primeiro.
            if encontrado and chave not in rodape:
                rodape[chave] = converter(encontrado.group(1))

        campos, repetidos = _coluna(itens, COLUNAS_LIDAS)

        alienacao = _numero(campos.get("alienacao"))

        if not alienacao:
            continue

        if repetidos:
            avisos.append(
                f"Página {numero_pagina} da Relação de Parcelas Lidas "
                f"trouxe mais de um campo na coluna {repetidos[0]}."
            )

        valor = _valor_brasileiro(campos.get("valor"))

        registro = {
            "pagina": numero_pagina,
            "alienacao": alienacao,
            "imovel": _numero(campos.get("imovel")),
            "modalidade": (campos.get("modalidade") or "").strip(),
            "parcela": (campos.get("parcela") or "").strip(),
            "data_pagamento": (campos.get("data_pagamento") or "").strip(),
            "data_credito": (campos.get("data_credito") or "").strip(),
            "codigo_baixa": (campos.get("codigo_baixa") or "").strip(),
            "boleto": (campos.get("nosso_numero") or "").strip(),
            "valor": valor if valor is not None else ZERO,
            "iptu": _valor_brasileiro(campos.get("iptu")) or ZERO,
        }

        if registro["valor"] > ZERO:
            parcelas.append(registro)
        else:
            # Boleto lido e não aplicado: o número dele está na
            # coluna "Imóvel".
            registro["boleto"] = registro["imovel"]
            sem_baixa.append(registro)

    # Cada parcela paga herda o boleto impresso na última
    # parcela do seu grupo.
    pendentes = []

    for parcela in parcelas:

        pendentes.append(parcela)

        if parcela["boleto"]:
            for anterior in pendentes:
                anterior["boleto"] = parcela["boleto"]

            pendentes = []

    if pendentes:
        avisos.append(
            f"{len(pendentes)} parcela(s) paga(s) na Relação de Parcelas "
            "Lidas ficaram sem número de boleto."
        )

    boletos = _agrupar_por_boleto(parcelas)

    total_baixado = sum((p["valor"] for p in parcelas), ZERO)

    esperado = rodape.get("total_boletos_baixados")

    if esperado is not None and esperado != len(boletos):
        avisos.append(
            f"A Relação de Parcelas Lidas diz {esperado} boletos "
            f"baixados, mas a leitura agrupou {len(boletos)}."
        )

    esperado = rodape.get("total_recebido")

    if esperado is not None and esperado != total_baixado:
        avisos.append(
            "O total lido na Relação de Parcelas Lidas não fecha com o "
            "TOTAL GERAL RECEBIDO do próprio relatório."
        )

    return {
        "tipo": "parcelas_lidas",
        "lote": cabecalho.get("lote", ""),
        "arquivo": cabecalho.get("arquivo", ""),

        "parcelas": parcelas,
        "boletos": boletos,
        "lidos_sem_baixa": sem_baixa,

        "quantidade_parcelas": len(parcelas),
        "quantidade_boletos": len(boletos),
        "quantidade_lidos_sem_baixa": len(
            {registro["boleto"] for registro in sem_baixa}
        ),

        "total_baixado": total_baixado,

        "rodape": rodape,
        "avisos": avisos,
    }


def _agrupar_por_boleto(parcelas):
    """
    Junta as parcelas pagas de cada boleto, somando o valor.
    """

    ordem = []
    grupos = {}

    for parcela in parcelas:

        boleto = parcela["boleto"]

        if boleto not in grupos:
            ordem.append(boleto)

            grupos[boleto] = {
                "boleto": boleto,
                "alienacao": parcela["alienacao"],
                "imovel": parcela["imovel"],
                "data_pagamento": parcela["data_pagamento"],
                "parcelas": [],
                "valor": ZERO,
            }

        grupos[boleto]["parcelas"].append(parcela)
        grupos[boleto]["valor"] += parcela["valor"]

    return [grupos[boleto] for boleto in ordem]


# ============================================================
# BAIXAS DE PAGAMENTOS NÃO EFETIVADAS
# ============================================================

# Cada registro ocupa três linhas: os códigos, os valores e o
# motivo da recusa.

MARCA_REGISTROS_NAO_BAIXADOS = "total de registros nao baixados"
MARCA_VALORES_NAO_BAIXADOS = "total de valores nao baixados"

REGEX_INTEIRO = re.compile(r"\b(\d+)\b")
REGEX_MOEDA = re.compile(r"-?[\d\.]*\d,\d{2}")

# Uma palavra de verdade, para separar a linha do motivo das
# linhas só de números e códigos.
REGEX_PALAVRA = re.compile(r"[A-Za-zÀ-ÿ]{4}")


def ler_baixas_nao_efetivadas(origem):
    """
    Lê o relatório "Baixas de Pagamentos Não Efetivadas" do GAI.

    Cada registro traz o boleto recusado e o motivo. A maior
    parte deles vem com total pago 0,00: são os registros de
    entrada de título do arquivo de retorno, que não carregam
    dinheiro. Os que vêm com valor são os pagamentos que o GAI
    recebeu e não conseguiu aplicar.
    """

    registros = []
    rodape = {}
    atual = None

    for numero_pagina, itens in _linhas_pdf(origem):

        texto = _texto_da_linha(itens)
        comparavel = _sem_acento(texto)

        # O rodapé do relatório. Vem antes do resto porque o
        # número às vezes é desenhado na frente do rótulo, e a
        # linha não pode ser confundida com um motivo de recusa.
        if MARCA_REGISTROS_NAO_BAIXADOS in comparavel:
            encontrado = REGEX_INTEIRO.search(texto)

            if encontrado:
                rodape["quantidade"] = int(encontrado.group(1))

            continue

        if MARCA_VALORES_NAO_BAIXADOS in comparavel:
            encontrado = REGEX_MOEDA.search(texto)

            if encontrado:
                rodape["total"] = _valor_brasileiro(encontrado.group())

            continue

        campos, _ = _coluna(itens, COLUNAS_NAO_EFETIVADAS)

        alienacao = _numero(campos.get("alienacao"))

        # ---------- primeira linha: os códigos ----------

        if alienacao:
            atual = {
                "pagina": numero_pagina,
                "alienacao": alienacao,
                "boleto": _numero(campos.get("boleto")),
                "banco": (campos.get("banco") or "").strip(),
                "agencia": (campos.get("agencia") or "").strip(),
                "tipo": (campos.get("tipo") or "").strip(),
                "parcela": (campos.get("parcela") or "").strip(),
                "data_pagamento": (campos.get("data_pagamento") or "").strip(),
                "data_vencimento": (
                    campos.get("data_vencimento") or ""
                ).strip(),
                "total_pago": None,
                "motivo": "",

                # O relatório não diz quem pagou. Quem preenche
                # é conciliacao.identificar_recusados, quando a
                # francesinha Bolebarra vem junto.
                "nome": "",
                "nosso_numero": "",
                "origem_do_nome": "",
            }

            registros.append(atual)
            continue

        if atual is None:
            continue

        # ---------- segunda linha: os valores ----------

        if atual["total_pago"] is None:

            valores = _ler_valores(itens)

            if valores.get("total_pago") is not None:
                atual.update(valores)
                continue

        # ---------- terceira linha: o motivo ----------

        # Só a primeira linha com texto depois dos valores. O
        # que vem adiante é o cabeçalho da página seguinte, que
        # se repete entre um registro e o próximo.
        if atual["motivo"]:
            continue

        if REGEX_PALAVRA.search(texto):
            atual["motivo"] = texto.strip()

    for registro in registros:
        if registro["total_pago"] is None:
            registro["total_pago"] = ZERO

        registro.setdefault("total_devido", ZERO)
        registro.setdefault("diferenca", ZERO)

    com_valor = [r for r in registros if r["total_pago"] > ZERO]

    total = sum((r["total_pago"] for r in registros), ZERO)

    avisos = []

    if rodape.get("quantidade") not in (None, len(registros)):
        avisos.append(
            f"O relatório de Baixas Não Efetivadas diz "
            f"{rodape['quantidade']} registros, mas a leitura "
            f"encontrou {len(registros)}."
        )

    if rodape.get("total") not in (None, total):
        avisos.append(
            "O total lido no relatório de Baixas Não Efetivadas não "
            "fecha com o total impresso no próprio relatório."
        )

    return {
        "tipo": "baixas_nao_efetivadas",
        "registros": registros,
        "com_valor": com_valor,

        "quantidade": len(registros),
        "quantidade_com_valor": len(com_valor),
        "quantidade_identificados": 0,

        "total": total,

        "motivos": _contar_motivos(registros),
        "rodape": rodape,
        "avisos": avisos,
    }


def _ler_valores(itens):
    """
    Lê a linha de valores de um registro recusado.

    Cada campo costuma vir num pedaço de texto próprio, e aí a
    coluna sai da coordenada. Quando o PDF manda vários campos
    colados, só o primeiro deles tem coordenada; nesse caso vale
    a ordem da faixa, em que o Total Pago é o último campo antes
    da Diferença — e a Diferença, por ficar na borda direita da
    página, chega sempre separada.

    A leitura é conferida no fim contra o total impresso no
    rodapé do próprio relatório.
    """

    campos = {}
    colados = []

    for x, texto in itens:

        partes = _partes(texto)

        if len(partes) > 1:
            colados = partes
            continue

        for nome, inicio, fim in COLUNAS_VALORES:
            if inicio <= x < fim:
                campos.setdefault(nome, texto.strip())
                break

    if colados and "total_pago" not in campos:
        campos["total_pago"] = colados[-1]

    return {
        "total_pago": _valor_brasileiro(campos.get("total_pago")),
        "total_devido": _valor_brasileiro(campos.get("total_devido")) or ZERO,
        "diferenca": _valor_brasileiro(campos.get("diferenca")) or ZERO,
    }


def _contar_motivos(registros):
    """
    Agrupa os registros recusados por motivo, para a tela
    mostrar o retrato do dia em poucas linhas.
    """

    ordem = []
    grupos = {}

    for registro in registros:

        motivo = registro["motivo"] or "Sem motivo informado"

        if motivo not in grupos:
            ordem.append(motivo)
            grupos[motivo] = {
                "motivo": motivo,
                "quantidade": 0,
                "total": ZERO,
            }

        grupos[motivo]["quantidade"] += 1
        grupos[motivo]["total"] += registro["total_pago"]

    return [grupos[motivo] for motivo in ordem]


# ============================================================
# LEITURA COM DETECÇÃO AUTOMÁTICA
# ============================================================

def ler_relatorio_gai(origem):
    """
    Lê qualquer um dos dois relatórios, descobrindo qual é pelo
    título impresso na primeira página.
    """

    tipo = identificar(origem)

    if tipo == "parcelas_lidas":
        return ler_parcelas_lidas(origem)

    if tipo == "baixas_nao_efetivadas":
        return ler_baixas_nao_efetivadas(origem)

    raise ValueError(
        "O PDF não é a Relação de Parcelas Lidas nem o relatório de "
        "Baixas de Pagamentos Não Efetivadas do GAI."
    )
