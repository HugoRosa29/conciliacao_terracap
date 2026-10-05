"""
Exportação da conciliação para Excel.

Usa xlsxwriter direto (sem pandas), para que o mesmo código
gere o arquivo no servidor e no navegador, via Pyodide.
"""

from decimal import Decimal
from io import BytesIO


AZUL = "#1f3864"

# Abas geradas, na ordem. Cada uma só aparece se os arquivos
# dela foram enviados.
ABAS = (
    "Resumo",
    "Bolebarra x GAI",
    "Bolebarra nao baixada",
    "Bolebarra x DCB",
    "GAI nao efetivadas",
    "PIX detalhado",
    "Bolepix nao baixado",
    "DCB liquidados",
    "Extrato creditos",
)

# Como cada situação da conferência com o GAI aparece na
# planilha.
SITUACAO_GAI = {
    "baixado": "Baixado no GAI",
    "lido_sem_baixa": "Lido pelo GAI com valor 0,00",
    "recusado": "Recusado pelo GAI",
    "ausente": "Fora dos relatorios do GAI",
}


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
    lidas = resultado.get("gai_lidas")
    nao_efetivadas = resultado.get("gai_nao_baixadas")
    com_dcb = resultado.get("bolebarras_dcb")
    com_gai = resultado.get("bolebarras_gai")
    bolepix_dcb = resultado.get("bolepix_dcb")

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

    if com_dcb is not None:
        resumo += [
            ["Bolebarra x DCB", "", ""],
            [
                "Cobranças conferidas título por título",
                com_dcb["quantidade_casados"],
                com_dcb["total_conferido"],
            ],
            [
                "Cobranças com valor diferente no DCB",
                com_dcb["quantidade_valor_divergente"],
                "",
            ],
            [
                "Cobranças sem baixa no DCB",
                com_dcb["quantidade_sem_baixa"],
                com_dcb["total_sem_baixa"],
            ],
            [
                "Liquidados no DCB fora da Bolebarra (Bolepix)",
                com_dcb["quantidade_fora_da_francesinha"],
                com_dcb["total_fora_da_francesinha"],
            ],
            ["", "", ""],
        ]

    if lidas is not None:
        resumo += [
            ["GAI - Relação de Parcelas Lidas", "", ""],
            [
                "Boletos baixados no GAI",
                lidas["quantidade_boletos"],
                lidas["total_baixado"],
            ],
            [
                "Boletos lidos com valor 0,00",
                lidas["quantidade_lidos_sem_baixa"],
                "",
            ],
            ["", "", ""],
        ]

    if nao_efetivadas is not None:
        resumo += [
            ["GAI - Baixas de Pagamentos Não Efetivadas", "", ""],
            [
                "Registros recusados",
                nao_efetivadas["quantidade"],
                nao_efetivadas["total"],
            ],
            ["", "", ""],
        ]

    if com_gai is not None:
        resumo += [
            ["Bolebarra x GAI", "", ""],
            [
                "Cobranças baixadas no GAI",
                com_gai["quantidade_baixados"],
                com_gai["total_baixado"],
            ],
            [
                "Cobranças NÃO baixadas no GAI",
                com_gai["quantidade_nao_baixados"],
                com_gai["total_nao_baixado"],
            ],
        ]

        for situacao in com_gai["situacoes"]:
            resumo.append([
                f"   {SITUACAO_GAI.get(situacao['situacao'], situacao['situacao'])}",
                situacao["quantidade"],
                situacao["total"],
            ])

        resumo += [
            [
                "Baixados no GAI fora da Bolebarra (Bolepix)",
                com_gai["quantidade_fora_da_francesinha"],
                com_gai["total_fora_da_francesinha"],
            ],
            ["", "", ""],
        ]

    if bolepix_dcb is not None:
        resumo += [
            ["Bolepix x DCB", "", ""],
            [
                "Bolepix sem baixa no arquivo de retorno",
                bolepix_dcb["quantidade_nao_baixados"],
                bolepix_dcb["total_nao_baixados"],
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

    for relatorio in resultado.get("sistemas", []):
        resumo.append([relatorio["sistema"] + " — pagamentos positivos",
                       relatorio["quantidade"], relatorio["total"]])

    benner_resumo = resultado.get("gai_benner")
    if benner_resumo is not None:
        estados = {"aguardando_gai": "Envie Parcelas Lidas do GAI",
                   "datas_incompativeis": "Datas incompatíveis", "comparado": "Comparado"}
        resumo.append(["GAI x BENNER: " + estados[benner_resumo["status"]], "", ""])
        if benner_resumo["status"] == "comparado":
            resumo.append(["GAI não localizado no BENNER", benner_resumo["quantidade_ausentes"], benner_resumo["total_ausente"]])
            resumo.append(["GAI x BENNER: valor divergente", benner_resumo["quantidade_divergentes"], ""])

    _escrever_aba(
        livro,
        "Resumo",
        ["Item", "Quantidade", "Valor"],
        resumo or [["Sem dados", "", ""]],
        colunas_moeda={2},
    )

    # ---------- GAI x BENNER ----------
    conferencia_benner = resultado.get("gai_benner")
    if conferencia_benner is not None:
        colunas_benner = ["Alienação", "Imóvel", "Parcela", "Boleto GAI", "Data pagamento", "Valor GAI"]
        campos_benner = ("alienacao", "imovel", "parcela", "boleto", "data_pagamento", "valor")
        if conferencia_benner["status"] == "comparado":
            _escrever_aba(livro, "GAI ausente BENNER", colunas_benner,
                [[r.get(k, "") for k in campos_benner] for r in conferencia_benner["ausentes"]], colunas_moeda={5})
            _escrever_aba(livro, "GAI BENNER valores", colunas_benner + ["Valor BENNER", "Diferença", "Aba BENNER", "Linha BENNER"],
                [[r.get(k, "") for k in campos_benner + ("valor_benner", "diferenca", "aba_benner", "linha_benner")]
                 for r in conferencia_benner["divergentes"]], colunas_moeda={5, 6, 7})
        if conferencia_benner["nao_comparados"]:
            _escrever_aba(livro, "GAI BENNER nao comparado", colunas_benner,
                [[r.get(k, "") for k in campos_benner] for r in conferencia_benner["nao_comparados"]], colunas_moeda={5})
        revisar = resultado["benner"]["nao_identificados"]
        if revisar:
            _escrever_aba(livro, "BENNER revisar", ["Aba", "Linha", "Data", "Valor", "Histórico"],
                [[r[k] for k in ("aba", "linha", "data_pagamento", "valor", "historico")] for r in revisar], colunas_moeda={3})

    # ---------- GIR / GGR / GOP ----------
    for relatorio in resultado.get("sistemas", []):
        _escrever_aba(livro, relatorio["sistema"] + " x banco", [
            "Sistema", "Documento / Processo", "Contrato", "Boleto", "Cliente / Ocupante",
            "Data pagamento", "Valor pago", "Página PDF", "DCB", "Valor DCB",
            "Bolebarra", "Valor Bolebarra",
        ], [[r.get(k) for k in (
            "sistema", "documento", "contrato", "boleto", "nome", "data_pagamento",
            "valor", "pagina", "situacao_dcb", "valor_dcb", "situacao_bolebarra", "valor_bolebarra",
        )] for r in relatorio["linhas"]], colunas_moeda={6, 9, 11})

    sistemas = resultado.get("bolebarras_sistemas")
    if sistemas is not None:
        _escrever_aba(livro, "Bolebarra x sistemas", [
            "Boleto", "Sacado", "Data pagamento", "Sistema", "Valor Bolebarra",
            "Valor nos sistemas", "Situação",
        ], [[r[k] for k in ("boleto", "nome", "data_pagamento", "sistema", "valor",
                            "valor_sistemas", "situacao")]
            for r in sistemas["linhas"]], colunas_moeda={4, 5})

    # ---------- Bolebarra x GAI ----------

    if com_gai is not None:

        colunas_gai = [
            "Nosso número",
            "Boleto (GAI)",
            "Alienação",
            "Sacado",
            "Dt. liquidação",
            "Valor recebido",
            "Valor baixado no GAI",
            "Diferença",
            "Situação",
            "Motivo informado pelo GAI",
        ]

        def linha_gai(linha):
            titulo = linha["titulo"]

            return [
                titulo["nosso_numero"],
                titulo["boleto"],
                linha["alienacao"],
                titulo["nome"],
                titulo["data"],
                titulo["valor"],
                linha["valor_gai"],
                linha["diferenca"],
                SITUACAO_GAI.get(linha["situacao"], linha["situacao"]),
                linha["motivo"],
            ]

        _escrever_aba(
            livro,
            "Bolebarra x GAI",
            colunas_gai,
            [linha_gai(linha) for linha in com_gai["linhas"]],
            colunas_moeda={5, 6, 7},
        )

        # A aba que o usuário abre primeiro: só o que não fechou.
        _escrever_aba(
            livro,
            "Bolebarra nao baixada",
            colunas_gai,
            [
                linha_gai(linha)
                for linha in com_gai["nao_baixados"]
                + com_gai["valor_divergente"]
            ],
            colunas_moeda={5, 6, 7},
        )

    # ---------- Bolebarra x DCB ----------

    if com_dcb is not None:

        linhas_dcb = []

        for par in com_dcb["casados"] + com_dcb["valor_divergente"]:
            linhas_dcb.append([
                par["titulo"]["nosso_numero"],
                par["titulo"]["nome"],
                par["titulo"]["data"],
                par["titulo"]["valor"],
                par["registro"]["valor"],
                par["diferenca"],
                par["registro"]["ocorrencia"],
                par["registro"]["ocorrencia_descricao"],
            ])

        for item in com_dcb["sem_baixa"]:
            linhas_dcb.append([
                item["titulo"]["nosso_numero"],
                item["titulo"]["nome"],
                item["titulo"]["data"],
                item["titulo"]["valor"],
                "",
                item["titulo"]["valor"],
                ", ".join(item["ocorrencias"]),
                "Sem título liquidado no arquivo de retorno",
            ])

        for registro in com_dcb["fora_da_francesinha"]:
            linhas_dcb.append([
                registro["nosso_numero"],
                "",
                registro["data_ocorrencia"],
                "",
                registro["valor"],
                -registro["valor"],
                registro["ocorrencia"],
                "Liquidado no DCB, fora da francesinha Bolebarra",
            ])

        _escrever_aba(
            livro,
            "Bolebarra x DCB",
            [
                "Nosso número",
                "Sacado",
                "Dt. liquidação",
                "Valor na francesinha",
                "Valor pago no DCB",
                "Diferença",
                "Ocorrência",
                "Situação",
            ],
            linhas_dcb,
            colunas_moeda={3, 4, 5},
        )

    # ---------- GAI: baixas não efetivadas ----------

    if nao_efetivadas is not None:
        _escrever_aba(
            livro,
            "GAI nao efetivadas",
            [
                "Alienação",
                "Sacado",
                "Boleto",
                "Nosso número",
                "Agência",
                "Parcela",
                "Data pagamento",
                "Data vencimento",
                "Total pago",
                "Motivo",
            ],
            [
                [
                    registro["alienacao"],
                    registro["nome"],
                    registro["boleto"],
                    registro["nosso_numero"],
                    registro["agencia"],
                    registro["parcela"],
                    registro["data_pagamento"],
                    registro["data_vencimento"],
                    registro["total_pago"],
                    registro["motivo"],
                ]
                # As recusas que trouxeram dinheiro primeiro, e
                # depois as que a francesinha soube nomear.
                for registro in sorted(
                    nao_efetivadas["registros"],
                    key=lambda r: (-r["total_pago"], not r["nome"]),
                )
            ],
            colunas_moeda={8},
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

    # Não depende do extrato: basta a francesinha Bolepix e o
    # arquivo de retorno.
    if bolepix_dcb is not None:
        _escrever_aba(
            livro,
            "Bolepix nao baixado",
            ["Data/Hora Pagamento", "Cliente / Contraparte", "Valor"],
            [
                [registro["data"], registro["nome"], registro["valor"]]
                for registro in bolepix_dcb["nao_baixados"]
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
