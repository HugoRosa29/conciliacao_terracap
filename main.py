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
"""

from pathlib import Path
import argparse
import sys

from conciliacao_terracap.conciliacao import conciliar, formatar_moeda


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
    }

    for caminho in sorted(pasta.iterdir()):

        if not caminho.is_file():
            continue

        nome = caminho.name.lower()

        if nome.startswith("dcb") and caminho.suffix.lower() == ".txt":
            achados["dcb"] = achados["dcb"] or caminho

        elif "bolepix" in nome:
            achados["francesinha"] = achados["francesinha"] or caminho

        elif "bolebarra" in nome:
            achados["bolebarras"] = achados["bolebarras"] or caminho

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
            print("\nBOLEPIX NÃO BAIXADO")
            print("-" * LARGURA)

            for numero, registro in enumerate(
                divergencia["nao_baixados"], start=1
            ):
                print(
                    f"  {numero:>3}. {registro['nome'][:38]:<40}"
                    f"{formatar_moeda(registro['valor']):>16}"
                    f"   {registro['data']}"
                )

            if divergencia["explicada"]:
                print(
                    "\n  A divergência está totalmente explicada pelos "
                    "itens acima."
                )
            else:
                print(
                    "\n  ATENÇÃO: os itens acima somam "
                    f"{formatar_moeda(divergencia['total_nao_baixados'])} "
                    "e não explicam toda a divergência."
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
    analisador.add_argument("--bolebarras", help="Francesinha Bolebarras")
    analisador.add_argument(
        "--pasta",
        help="Pasta com os arquivos do dia, reconhecidos pelo nome",
    )
    analisador.add_argument(
        "--excel",
        help="Grava o resultado em um arquivo .xlsx",
    )

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
        from conciliacao_terracap.exportar import montar_planilha

        destino = Path(argumentos.excel)
        destino.write_bytes(montar_planilha(resultado))

        print(f"Excel gravado em {destino}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
