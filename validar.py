"""
Validação da conciliação contra os arquivos reais de 28/09/2026.

Os valores esperados foram tirados da planilha de conciliação
feita à mão ("Planilha Conciliação Boleto c. movimento - 28 09 2026"),
aba 28.09.26:

    Francesinha Bolepix .................    687.438,67
    Francesinha Bolebarras ..............  3.645.052,09
    Total pagamentos QR Code + C.Barras .  4.332.490,76
    Total arquivo de retorno ............  4.319.046,36
    Divergência (Bolepix não baixado) ...     13.444,40

Uso:
    python validar.py [pasta]

A pasta padrão é ../PROMPT.
"""

from pathlib import Path
from decimal import Decimal
import sys

from conciliacao_terracap.conciliacao import conciliar, formatar_moeda


PASTA_PADRAO = Path(__file__).resolve().parent.parent / "PROMPT"

ARQUIVOS = {
    "extrato": "28 09 2026 m.txt",
    "dcb": "DCB_1219001012_28092026_030540.txt",
    "francesinha": "Francesinha movimento bolepix - 28.09.26.xlt.xls",
}

ESPERADO = {
    "quantidade_pix": 135,
    "quantidade_cobranca": 4,
    "total_pix": Decimal("687438.67"),
    "total_cobranca": Decimal("3645052.09"),
    "total_pagamentos": Decimal("4332490.76"),
    "quantidade_francesinha": 135,
    "total_francesinha": Decimal("687438.67"),
    "quantidade_liquidados": 481,
    "total_retorno": Decimal("4319046.36"),
    "divergencia": Decimal("13444.40"),
    "quantidade_nao_baixados": 5,
}


def comparar(titulo, obtido, esperado, moeda=False):
    """
    Imprime uma linha de comparação e devolve True se confere.
    """

    confere = obtido == esperado

    mostrar = formatar_moeda if moeda else str

    print(
        f"  {'OK ' if confere else 'ERRO'}  {titulo:<42}"
        f" obtido {mostrar(obtido):>16}"
        f"   esperado {mostrar(esperado):>16}"
    )

    return confere


def main(pasta=PASTA_PADRAO):
    pasta = Path(pasta)

    faltando = [
        nome
        for nome in ARQUIVOS.values()
        if not (pasta / nome).exists()
    ]

    if faltando:
        print(f"Arquivos não encontrados em {pasta}:")

        for nome in faltando:
            print("   -", nome)

        return 2

    print(f"Pasta: {pasta}\n")

    resultado = conciliar(
        extrato=pasta / ARQUIVOS["extrato"],
        dcb=pasta / ARQUIVOS["dcb"],
        francesinha=pasta / ARQUIVOS["francesinha"],
    )

    extrato = resultado["extrato"]
    francesinha = resultado["francesinha"]
    dcb = resultado["dcb"]
    divergencia = resultado["divergencia"]

    verificacoes = []

    print("EXTRATO BRB")
    verificacoes += [
        comparar(
            "Quantidade de créditos PIX",
            extrato["quantidade_pix"],
            ESPERADO["quantidade_pix"],
        ),
        comparar(
            "Total PIX (Bolepix)",
            extrato["total_pix"],
            ESPERADO["total_pix"],
            moeda=True,
        ),
        comparar(
            "Quantidade de créditos cobrança",
            extrato["quantidade_cobranca"],
            ESPERADO["quantidade_cobranca"],
        ),
        comparar(
            "Total cobrança (Bolebarras)",
            extrato["total_cobranca"],
            ESPERADO["total_cobranca"],
            moeda=True,
        ),
        comparar(
            "Total pagamentos QR Code + C.Barras",
            extrato["total_pagamentos"],
            ESPERADO["total_pagamentos"],
            moeda=True,
        ),
    ]

    print("\nFRANCESINHA BOLEPIX")
    verificacoes += [
        comparar(
            "Quantidade de recebimentos",
            francesinha["quantidade"],
            ESPERADO["quantidade_francesinha"],
        ),
        comparar(
            "Total",
            francesinha["total"],
            ESPERADO["total_francesinha"],
            moeda=True,
        ),
        comparar(
            "Fecha com o PIX do extrato",
            francesinha["conferencia_extrato"]["confere"],
            True,
        ),
    ]

    print("\nDCB - ARQUIVO DE RETORNO")
    verificacoes += [
        comparar(
            "Títulos liquidados",
            dcb["quantidade_liquidados"],
            ESPERADO["quantidade_liquidados"],
        ),
        comparar(
            "Total arquivo de retorno",
            dcb["total_liquidado"],
            ESPERADO["total_retorno"],
            moeda=True,
        ),
    ]

    print("\nDIVERGÊNCIA")
    verificacoes += [
        comparar(
            "Divergência",
            divergencia["divergencia"],
            ESPERADO["divergencia"],
            moeda=True,
        ),
        comparar(
            "Itens Bolepix não baixados",
            divergencia["quantidade_nao_baixados"],
            ESPERADO["quantidade_nao_baixados"],
        ),
        comparar(
            "Soma dos não baixados = divergência",
            divergencia["total_nao_baixados"],
            ESPERADO["divergencia"],
            moeda=True,
        ),
    ]

    print("\nBOLEPIX NÃO BAIXADO (detalhe)")

    for numero, registro in enumerate(
        divergencia["nao_baixados"], start=1
    ):
        print(
            f"  {numero}. {registro['nome'][:40]:<42}"
            f"{formatar_moeda(registro['valor']):>16}"
            f"   {registro['data']}"
        )

    print("\nPIX IDENTIFICADOS")

    detalhado = resultado["pix_detalhado"]

    verificacoes.append(
        comparar(
            "PIX do extrato com pagador identificado",
            detalhado["quantidade_identificados"],
            ESPERADO["quantidade_pix"],
        )
    )

    if resultado["avisos"]:
        print("\nAVISOS")

        for aviso in resultado["avisos"]:
            print("  -", aviso)

    total = len(verificacoes)
    ok = sum(1 for v in verificacoes if v)

    print(f"\n{'=' * 70}")
    print(f"{ok}/{total} verificações passaram")
    print("=" * 70)

    return 0 if ok == total else 1


if __name__ == "__main__":
    pasta = sys.argv[1] if len(sys.argv) > 1 else PASTA_PADRAO

    sys.exit(main(pasta))
