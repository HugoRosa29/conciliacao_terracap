"""
Exportação da conciliação para Excel.

Usa xlsxwriter direto (sem pandas), para que o mesmo código
gere o arquivo no servidor e no navegador, via Pyodide.
"""

from decimal import Decimal
from io import BytesIO


AZUL = "#1f3864"

ABAS = (
    "Resumo",
    "PIX detalhado",
    "Bolepix nao baixado",
    "DCB liquidados",
    "Extrato creditos",
)


def _valor(celula):
    """
    Converte Decimal em float para o xlsxwriter escrever como
    número, mantendo o resto como está.
    """

    if isinstance(celula, Decimal):
        return float(celula)

    return celula


def _escrever_aba(livro, titulo, colunas, linhas, colunas_moeda=()):
    """
    Cria uma aba com cabeçalho fixo, largura automática e
    formato de moeda nas colunas indicadas.
    """

    pagina = livro.add_worksheet(titulo[:31])

    formato_cabecalho = livro.add_format({
        "bold": True,
        "bg_color": AZUL,
        "font_color": "white",
        "border": 1,
    })

    formato_moeda = livro.add_format({"num_format": "R$ #,##0.00"})
    formato_negrito = livro.add_format({"bold": True})

    for indice, coluna in enumerate(colunas):
        pagina.write(0, indice, coluna, formato_cabecalho)

    larguras = [len(coluna) for coluna in colunas]

    for numero_linha, linha in enumerate(linhas, start=1):

        for indice, celula in enumerate(linha):

            valor = _valor(celula)

            formato = None

            if indice in colunas_moeda and isinstance(valor, float):
                formato = formato_moeda

            elif isinstance(valor, str) and valor and not any(
                item for item in linha[indice + 1:]
            ) and indice == 0:
                # Linha de título dentro do resumo.
                formato = formato_negrito

            pagina.write(numero_linha, indice, valor, formato)

            larguras[indice] = max(larguras[indice], len(str(valor or "")))

    for indice, largura in enumerate(larguras):
        pagina.set_column(
            indice,
            indice,
            min(max(largura + 2, 10), 45),
            formato_moeda if indice in colunas_moeda else None,
        )

    pagina.freeze_panes(1, 0)

    if linhas:
        pagina.autofilter(0, 0, len(linhas), len(colunas) - 1)

    return pagina


def montar_planilha(resultado):
    """
    Monta o .xlsx da conciliação em memória e devolve os bytes.
    """

    import xlsxwriter

    extrato = resultado.get("extrato")
    francesinha = resultado.get("francesinha")
    bolebarras = resultado.get("bolebarras")
    dcb = resultado.get("dcb")
    divergencia = resultado.get("divergencia")
    detalhado = resultado.get("pix_detalhado")

    buffer = BytesIO()

    livro = xlsxwriter.Workbook(buffer, {"in_memory": True})

    # ---------- Resumo ----------

    resumo = []

    if extrato is not None:
        resumo += [
            ["Extrato BRB", "", ""],
            [
                "CRED PIX QR CODE DINAMICO (Bolepix)",
                extrato["quantidade_pix"],
                extrato["total_pix"],
            ],
            [
                "CREDITO COBRANCA BRB (Bolebarras)",
                extrato["quantidade_cobranca"],
                extrato["total_cobranca"],
            ],
            [
                "Total pagamentos QR Code + C.Barras",
                extrato["quantidade_pix"] + extrato["quantidade_cobranca"],
                extrato["total_pagamentos"],
            ],
            ["", "", ""],
        ]

    if francesinha is not None:
        resumo += [
            ["Francesinha Bolepix", "", ""],
            [
                "Recebimentos PIX",
                francesinha["quantidade"],
                francesinha["total"],
            ],
        ]

        conferencia = francesinha.get("conferencia_extrato")

        if conferencia:
            resumo.append([
                "Diferença contra o PIX do extrato",
                "OK" if conferencia["confere"] else "DIVERGENTE",
                conferencia["diferenca"],
            ])

        resumo.append(["", "", ""])

    if bolebarras is not None:
        resumo += [
            ["Francesinha Bolebarras", "", ""],
            [
                "Recebimentos",
                bolebarras["quantidade"],
                bolebarras["total"],
            ],
            ["", "", ""],
        ]

    if dcb is not None:
        resumo += [
            ["DCB - arquivo de retorno", "", ""],
            ["Títulos no arquivo", dcb["quantidade_registros"], ""],
            [
                "Total arquivo de retorno",
                dcb["quantidade_liquidados"],
                dcb["total_liquidado"],
            ],
            ["", "", ""],
        ]

    if divergencia is not None:
        resumo += [
            ["Divergência", "", ""],
            ["Total pagamentos", "", divergencia["total_pagamentos"]],
            [
                "Total arquivo de retorno",
                "",
                divergencia["total_retorno"],
            ],
            [
                "Divergência (Bolepix não baixado)",
                divergencia["quantidade_nao_baixados"],
                divergencia["divergencia"],
            ],
        ]

    _escrever_aba(
        livro,
        "Resumo",
        ["Item", "Quantidade", "Valor"],
        resumo or [["Sem dados", "", ""]],
        colunas_moeda={2},
    )

    # ---------- PIX detalhado ----------

    if detalhado is not None:
        _escrever_aba(
            livro,
            "PIX detalhado",
            [
                "Data",
                "Descrição",
                "Valor",
                "Cliente / Contraparte",
                "CNPJ / CPF",
                "Data/Hora Pagamento",
            ],
            [
                [
                    linha["data"],
                    linha["descricao"],
                    linha["valor"],
                    linha["nome"],
                    linha["documento"],
                    linha["data_pagamento"],
                ]
                for linha in detalhado["linhas"]
            ],
            colunas_moeda={2},
        )

    # ---------- Bolepix não baixado ----------

    if divergencia is not None:
        _escrever_aba(
            livro,
            "Bolepix nao baixado",
            ["Data/Hora Pagamento", "Cliente / Contraparte", "Valor"],
            [
                [registro["data"], registro["nome"], registro["valor"]]
                for registro in divergencia["nao_baixados"]
            ],
            colunas_moeda={2},
        )

    # ---------- DCB liquidados ----------

    if dcb is not None:
        _escrever_aba(
            livro,
            "DCB liquidados",
            [
                "Ocorrência",
                "Descrição",
                "Nosso número",
                "Seu número",
                "CNPJ / CPF pagador",
                "Data ocorrência",
                "Data crédito",
                "Valor pago",
            ],
            [
                [
                    registro["ocorrencia"],
                    registro["ocorrencia_descricao"],
                    registro["nosso_numero"],
                    registro["seu_numero"],
                    registro["doc_pagador"],
                    registro["data_ocorrencia"],
                    registro["data_credito"],
                    registro["valor"],
                ]
                for registro in dcb["liquidados"]
            ],
            colunas_moeda={7},
        )

    # ---------- Extrato completo ----------

    if extrato is not None:
        creditos = (
            extrato["pix"]
            + extrato["cobranca"]
            + extrato["outros_creditos"]
        )

        _escrever_aba(
            livro,
            "Extrato creditos",
            ["Data", "Descrição", "DOC", "Valor"],
            [
                [
                    registro["data"],
                    registro["descricao"],
                    registro["doc"],
                    registro["valor"],
                ]
                for registro in creditos
            ],
            colunas_moeda={3},
        )

    livro.close()

    buffer.seek(0)

    return buffer.getvalue()
