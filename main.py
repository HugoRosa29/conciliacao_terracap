"""
Conciliação financeira pela linha de comando.

Mesma conciliação da aplicação web, para quem prefere rodar
direto no terminal ou automatizar o fechamento do dia.

Uso:

    python main.py --extrato "28 09 2026 m.txt" \\
                   --francesinha "Francesinha bolepix.xls" \\
                   --dcb "DCB_1219001012.txt"

    python main.py --pasta ../PROMPT

    python main.py --pasta ../PROMPT --excel conciliacao.xlsx

Com --pasta, os arquivos são reconhecidos pelo nome:

    extrato      .txt que contém "Extrato" do BRB
    dcb          .txt que começa com DCB
    francesinha  planilha com "bolepix" no nome
    bolebarras   planilha com "bolebarra" no nome
    GAI          os .pdf da pasta, na ordem em que aparecem
                 (cada relatório se identifica pelo próprio
                 título, então a ordem não importa)

O extrato é opcional: a conferência da Bolebarra com o DCB e com
o GAI é feita título por título e não precisa dele.
"""

from pathlib import Path
import argparse
import sys

from conciliacao import conciliar, formatar_moeda


LARGURA = 72


# ============================================================
# DESCOBERTA DE ARQUIVOS
# ============================================================

def descobrir(pasta):
    """
    Encontra os arquivos do dia dentro de uma pasta, pelo nome.
    """

    pasta = Path(pasta)

    achados = {
        "extrato": None,
        "dcb": None,
        "francesinha": None,
        "bolebarras": None,
        "gai_lidas": None,
        "gai_nao_baixadas": None,
        "gir": None,
        "ggr": None,
        "gop": None,
        "benner": None,
    }

    for caminho in sorted(pasta.iterdir()):

        if not caminho.is_file():
            continue

        nome = caminho.name.lower()

        if nome.startswith("dcb") and caminho.suffix.lower() == ".txt":
            achados["dcb"] = achados["dcb"] or caminho

        elif "benner" in nome and caminho.suffix.lower() in (".xlt", ".xls", ".xlsx", ".xltx"):
            if achados["benner"] is not None:
                raise ValueError("Mais de um arquivo BENNER na pasta; selecione com --benner.")
            achados["benner"] = caminho

        elif "bolepix" in nome:
            achados["francesinha"] = achados["francesinha"] or caminho

        elif "bolebarra" in nome:
            achados["bolebarras"] = achados["bolebarras"] or caminho

        elif caminho.suffix.lower() == ".pdf":
            from pypdf import PdfReader
            from sistemas import identificar_sistema

            sistema = identificar_sistema(" ".join(
                pagina.extract_text() or "" for pagina in PdfReader(caminho).pages
            ))
            if sistema:
                campo = sistema.lower()
                if achados[campo] is not None:
                    raise ValueError(f"Mais de um PDF de {sistema} na pasta; selecione com --{campo}.")
                achados[campo] = caminho
                continue
            # Qual dos dois relatórios do GAI é cada PDF sai do
            # título impresso dentro dele, não do nome.
            if achados["gai_lidas"] is None:
                achados["gai_lidas"] = caminho
            elif achados["gai_nao_baixadas"] is None:
                achados["gai_nao_baixadas"] = caminho

        elif caminho.suffix.lower() == ".txt":
            achados["extrato"] = achados["extrato"] or caminho

    return achados


# ============================================================
# RELATÓRIO
# ============================================================

def titulo(texto):
    print()
    print("=" * LARGURA)
    print(texto)
    print("=" * LARGURA)


def linha(rotulo, valor, largura=44):
    print(f"  {rotulo:<{largura}} {valor:>22}")


SITUACAO_GAI = {
    "lido_sem_baixa": "lido pelo GAI com valor 0,00",
    "recusado": "recusado pelo GAI",
    "ausente": "fora dos relatórios do GAI",
}


def _relatar_gai(resultado):
    """
    Os totais dos dois relatórios do GAI, como eles mesmos
    declaram no rodapé.
    """

    lidas = resultado["gai_lidas"]
    nao_baixadas = resultado["gai_nao_baixadas"]

    if lidas is None and nao_baixadas is None:
        return

    print("\nGAI - RELATÓRIOS")
    print("-" * LARGURA)

    if lidas is not None:
        linha(
            f"Boletos baixados ({lidas['quantidade_boletos']})",
            formatar_moeda(lidas["total_baixado"]),
        )
        linha(
            "Boletos lidos com valor 0,00",
            str(lidas["quantidade_lidos_sem_baixa"]),
        )

        if lidas["arquivo"]:
            print(f"\n  Gerado a partir de {lidas['arquivo']}")

    if nao_baixadas is not None:
        linha(
            f"Baixas não efetivadas ({nao_baixadas['quantidade']})",
            formatar_moeda(nao_baixadas["total"]),
        )

        print("\n  Motivos:")

        for motivo in nao_baixadas["motivos"]:
            print(
                f"    {motivo['quantidade']:>4}  "
                f"{formatar_moeda(motivo['total']):>16}  "
                f"{motivo['motivo'][:40]}"
            )

        # As recusas que trouxeram dinheiro, com dono e
        # alienação — o resto é entrada de título.
        recusas = [
            registro
            for registro in nao_baixadas["registros"]
            if registro["total_pago"] > 0 or registro["nome"]
        ]

        if recusas:
            print("\n  Recusas identificadas:")

            for registro in sorted(
                recusas, key=lambda r: (-r["total_pago"], r["alienacao"])
            ):
                print(
                    f"    alienação {registro['alienacao']:<8}"
                    f" boleto {registro['boleto']:<8}"
                    f" parcela {registro['parcela']:<5}"
                    f"{formatar_moeda(registro['total_pago']):>14}"
                    f"   {registro['nome'][:32] or '—'}"
                )


def _relatar_bolebarras_dcb(conferencia):
    """
    A Bolebarra conferida com o arquivo de retorno, título por
    título.
    """

    if conferencia is None:
        return

    print("\nFRANCESINHA BOLEBARRA x DCB")
    print("-" * LARGURA)

    linha(
        f"Conferidas pelo nosso número ({conferencia['quantidade_casados']})",
        formatar_moeda(conferencia["total_conferido"]),
    )

    if conferencia["confere"]:
        print("\n  Toda a francesinha Bolebarra está liquidada no DCB.")
    else:
        for rotulo, chave, total in (
            ("Com valor diferente no DCB", "valor_divergente", None),
            ("Sem baixa no DCB", "sem_baixa", "total_sem_baixa"),
            ("Sem nosso número na francesinha", "sem_nosso_numero", None),
        ):
            quantidade = conferencia[f"quantidade_{chave}"]

            if quantidade:
                linha(
                    f"{rotulo} ({quantidade})",
                    formatar_moeda(conferencia[total]) if total else "",
                )

        for par in conferencia["valor_divergente"]:
            print(
                f"    {par['titulo']['nosso_numero']}  "
                f"{par['titulo']['nome'][:28]:<30}"
                f"{formatar_moeda(par['titulo']['valor']):>16} contra "
                f"{formatar_moeda(par['registro']['valor'])} no DCB"
            )

        for item in conferencia["sem_baixa"]:
            print(
                f"    {item['titulo']['nosso_numero']}  "
                f"{item['titulo']['nome'][:28]:<30}"
                f"{formatar_moeda(item['titulo']['valor']):>16}"
                f"   ocorrências: {', '.join(item['ocorrencias']) or '—'}"
            )

    linha(
        "Liquidados no DCB fora da Bolebarra "
        f"({conferencia['quantidade_fora_da_francesinha']})",
        formatar_moeda(conferencia["total_fora_da_francesinha"]),
    )

    print("  (são os recebimentos do Bolepix, que não têm francesinha")
    print("   de cobrança própria)")


def _relatar_bolebarras_gai(conferencia):
    """
    O que o GAI fez com cada cobrança da Bolebarra.
    """

    if conferencia is None:
        return

    titulo("FRANCESINHA BOLEBARRA x GAI")

    linha(
        f"Baixadas no GAI ({conferencia['quantidade_baixados']})",
        formatar_moeda(conferencia["total_baixado"]),
    )
    linha(
        f"NÃO baixadas no GAI ({conferencia['quantidade_nao_baixados']})",
        formatar_moeda(conferencia["total_nao_baixado"]),
    )

    for situacao in conferencia["situacoes"]:
        linha(
            "   "
            + SITUACAO_GAI.get(situacao["situacao"], situacao["situacao"])
            + f" ({situacao['quantidade']})",
            formatar_moeda(situacao["total"]),
        )

    linha(
        "Baixados no GAI fora da Bolebarra "
        f"({conferencia['quantidade_fora_da_francesinha']})",
        formatar_moeda(conferencia["total_fora_da_francesinha"]),
    )

    if conferencia["valor_divergente"]:
        print("\nBAIXADAS NO GAI POR VALOR DIFERENTE")
        print("-" * LARGURA)

        for item in conferencia["valor_divergente"]:
            print(
                f"  {item['titulo']['nosso_numero']}  "
                f"{item['titulo']['nome'][:26]:<28}"
                f"{formatar_moeda(item['titulo']['valor']):>16} contra "
                f"{formatar_moeda(item['valor_gai'])} no GAI"
            )

    if not conferencia["nao_baixados"]:
        print("\n  Toda a francesinha Bolebarra foi baixada no GAI.")
        return

    print("\nBOLEBARRA NÃO BAIXADA NO GAI")
    print("-" * LARGURA)

    for numero, item in enumerate(conferencia["nao_baixados"], start=1):
        print(
            f"  {numero:>3}. {item['titulo']['nome'][:30]:<32}"
            f"{formatar_moeda(item['titulo']['valor']):>16}"
            f"   boleto {item['titulo']['boleto']}"
            f"   alienação {item['alienacao'] or '—'}"
        )
        print(
            f"       {SITUACAO_GAI.get(item['situacao'], item['situacao'])}"
            + (f" — {item['motivo'][:70]}" if item["motivo"] else "")
        )


def _relatar_bolepix(bolepix):
    """
    Os PIX recebidos que não têm baixa no arquivo de retorno.
    """

    if bolepix is None or not bolepix["nao_baixados"]:
        return

    print("\nBOLEPIX NÃO BAIXADO")
    print("-" * LARGURA)

    for numero, registro in enumerate(bolepix["nao_baixados"], start=1):
        print(
            f"  {numero:>3}. {registro['nome'][:38]:<40}"
            f"{formatar_moeda(registro['valor']):>16}"
            f"   {registro['data']}"
        )

    linha(
        f"Total não baixado ({bolepix['quantidade_nao_baixados']})",
        formatar_moeda(bolepix["total_nao_baixados"]),
    )


def gerar_relatorio(resultado):
    """
    Imprime a conciliação no terminal.
    """

    extrato = resultado["extrato"]
    francesinha = resultado["francesinha"]
    bolebarras = resultado["bolebarras"]
    dcb = resultado["dcb"]
    divergencia = resultado["divergencia"]
    detalhado = resultado["pix_detalhado"]

    titulo("CONCILIAÇÃO FINANCEIRA")

    benner = resultado.get("gai_benner")
    if benner is not None:
        print("\nGAI x BENNER")
        if benner["status"] == "aguardando_gai":
            print("  Envie a Relação de Parcelas Lidas do GAI.")
        elif benner["status"] == "datas_incompativeis":
            print("  Datas incompatíveis: envie arquivos do mesmo período.")
        else:
            linha("Parcelas não localizadas no BENNER", str(benner["quantidade_ausentes"]))
            linha("Total não localizado", formatar_moeda(benner["total_ausente"]))
            linha("Parcelas com valor divergente", str(benner["quantidade_divergentes"]))
            for item in benner["ausentes"]:
                print(f"  Alienação {item['alienacao']} | imóvel {item['imovel']}"
                      f" | parcela {item['parcela']} | {item['data_pagamento']}"
                      f" | {formatar_moeda(item['valor'])}")

    for relatorio in resultado.get("sistemas", []):
        print(f"\n{relatorio['sistema']} — {relatorio['quantidade']} pagamentos")
        linha("Total pago", formatar_moeda(relatorio["total"]))
        for item in relatorio["linhas"]:
            print(f"  {item['boleto']}  {formatar_moeda(item['valor']):>16}"
                  f"  DCB: {item['situacao_dcb']} | Bolebarra: {item['situacao_bolebarra']}")

    if extrato is not None:
        print("\nEXTRATO BRB")
        print("-" * LARGURA)

        linha(
            f"CRED PIX QR CODE DINAMICO ({extrato['quantidade_pix']})",
            formatar_moeda(extrato["total_pix"]),
        )
        linha(
            f"CREDITO COBRANCA BRB ({extrato['quantidade_cobranca']})",
            formatar_moeda(extrato["total_cobranca"]),
        )
        linha(
            "Total pagamentos QR Code + C.Barras",
            formatar_moeda(extrato["total_pagamentos"]),
        )

    for rotulo, dados in (
        ("FRANCESINHA BOLEPIX", francesinha),
        ("FRANCESINHA BOLEBARRAS", bolebarras),
    ):
        if dados is None:
            continue

        print(f"\n{rotulo}")
        print("-" * LARGURA)

        linha(
            f"Recebimentos ({dados['quantidade']})",
            formatar_moeda(dados["total"]),
        )

        conferencia = dados.get("conferencia_extrato")

        if conferencia:
            linha(
                "Diferença contra o extrato",
                formatar_moeda(conferencia["diferenca"])
                + ("  OK" if conferencia["confere"] else "  !!"),
            )

    if dcb is not None:
        print("\nDCB - ARQUIVO DE RETORNO")
        print("-" * LARGURA)

        linha("Títulos no arquivo", str(dcb["quantidade_registros"]))
        linha(
            f"Títulos liquidados ({dcb['quantidade_liquidados']})",
            formatar_moeda(dcb["total_liquidado"]),
        )

        print("\n  Ocorrências:")

        for ocorrencia in dcb["ocorrencias"]:
            print(
                f"    {ocorrencia['codigo']}  "
                f"{ocorrencia['descricao']:<40}"
                f"{ocorrencia['quantidade']:>6}"
            )

    _relatar_gai(resultado)
    _relatar_bolebarras_dcb(resultado["bolebarras_dcb"])
    _relatar_bolebarras_gai(resultado["bolebarras_gai"])
    _relatar_bolepix(resultado["bolepix_dcb"])

    if divergencia is not None:
        titulo("DIVERGÊNCIA")

        linha(
            "Total de pagamentos",
            formatar_moeda(divergencia["total_pagamentos"]),
        )
        linha(
            "Total do arquivo de retorno",
            formatar_moeda(divergencia["total_retorno"]),
        )
        linha(
            "Divergência",
            formatar_moeda(divergencia["divergencia"]),
        )

        if divergencia["nao_baixados"]:
            linha(
                "Bolepix não baixado "
                f"({divergencia['quantidade_nao_baixados']})",
                formatar_moeda(divergencia["total_nao_baixados"]),
            )

        if divergencia["dcb_sem_francesinha"]:
            linha(
                "Liquidado no DCB sem francesinha "
                f"({divergencia['quantidade_dcb_sem_francesinha']})",
                formatar_moeda(-divergencia["total_dcb_sem_francesinha"]),
            )

            for registro in divergencia["dcb_sem_francesinha"]:
                print(
                    f"    nosso número {registro['nosso_numero']}"
                    f"   alienação {registro['seu_numero']:<8}"
                    f"{formatar_moeda(registro['valor']):>16}"
                )

        if divergencia["nao_baixados"] or divergencia["dcb_sem_francesinha"]:
            print(
                "  Explicada pelos itens acima".ljust(46)
                + ("OK" if divergencia["explicada"] else "!!").rjust(23)
            )

            if not divergencia["explicada"]:
                print(
                    "\n  ATENÇÃO: os itens acima não explicam toda a "
                    "divergência. Veja a conferência da Bolebarra."
                )

    if detalhado is not None:
        nao_identificados = detalhado["quantidade_nao_identificados"]

        if nao_identificados:
            print(
                f"\n  {nao_identificados} crédito(s) PIX do extrato sem "
                "pagador identificado na francesinha."
            )

    if resultado["avisos"]:
        print("\nAVISOS")
        print("-" * LARGURA)

        for aviso in resultado["avisos"]:
            print("  -", aviso)

    print()


# ============================================================
# EXECUÇÃO
# ============================================================

def main(argv=None):
    analisador = argparse.ArgumentParser(
        description="Conciliação do extrato BRB com o Bolepix e o DCB.",
    )

    analisador.add_argument("--extrato", help="TXT do extrato BRB")
    analisador.add_argument("--francesinha", help="Francesinha Bolepix")
    analisador.add_argument("--dcb", help="Arquivo de retorno (CNAB 400)")
    analisador.add_argument("--bolebarras", help="Francesinha Bolebarra")
    analisador.add_argument(
        "--gai-lidas",
        help="PDF da Relação de Parcelas Lidas do GAI",
    )
    analisador.add_argument(
        "--gai-nao-baixadas",
        help="PDF das Baixas de Pagamentos Não Efetivadas do GAI",
    )
    analisador.add_argument(
        "--pasta",
        help="Pasta com os arquivos do dia, reconhecidos pelo nome",
    )
    analisador.add_argument(
        "--excel",
        help="Grava o resultado em um arquivo .xlsx",
    )

    for sistema in ("gir", "ggr", "gop"):
        analisador.add_argument(f"--{sistema}", help=f"PDF do {sistema.upper()}")

    analisador.add_argument("--benner", help="Planilha de integração GAI para BENNER (XLT/XLS/XLSX)")
    argumentos = analisador.parse_args(argv)

    if argumentos.pasta:
        arquivos = descobrir(argumentos.pasta)

        print(f"Pasta: {argumentos.pasta}")

        for rotulo, caminho in arquivos.items():
            print(f"  {rotulo:<13} {caminho.name if caminho else '—'}")
    else:
        arquivos = {
            "extrato": argumentos.extrato,
            "francesinha": argumentos.francesinha,
            "dcb": argumentos.dcb,
            "bolebarras": argumentos.bolebarras,
            "gai_lidas": argumentos.gai_lidas,
            "gai_nao_baixadas": argumentos.gai_nao_baixadas,
            "gir": argumentos.gir,
            "ggr": argumentos.ggr,
            "gop": argumentos.gop,
            "benner": argumentos.benner,
        }

    arquivos = {
        rotulo: caminho
        for rotulo, caminho in arquivos.items()
        if caminho
    }

    if not arquivos:
        analisador.error(
            "Informe ao menos um arquivo, ou use --pasta."
        )

    faltando = [
        str(caminho)
        for caminho in arquivos.values()
        if not Path(caminho).exists()
    ]

    if faltando:
        print("Arquivo não encontrado:", file=sys.stderr)

        for caminho in faltando:
            print("   -", caminho, file=sys.stderr)

        return 2

    resultado = conciliar(**arquivos)

    gerar_relatorio(resultado)

    if argumentos.excel:
        from exportar import montar_planilha

        destino = Path(argumentos.excel)
        destino.write_bytes(montar_planilha(resultado))

        print(f"Excel gravado em {destino}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
