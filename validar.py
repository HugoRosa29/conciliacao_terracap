"""
Validação da conciliação contra os arquivos reais.

Os valores esperados saem das planilhas de conciliação feitas à
mão, que são o gabarito de cada dia.

28/09/2026 — "Planilha Conciliação Boleto c. movimento -
28 09 2026", aba 28.09.26. Fecha o dia pelo extrato:

    Francesinha Bolepix .................    687.438,67
    Francesinha Bolebarras ..............  3.645.052,09
    Total pagamentos QR Code + C.Barras .  4.332.490,76
    Total arquivo de retorno ............  4.319.046,36
    Divergência (Bolepix não baixado) ...     13.444,40

30/09/2026 — "Planilha Conciliação Boleto c. movimento -
30 09 2026 - 2", aba "30 09 2026". Não tem extrato: é o dia em
que a conferência passa pelos relatórios do GAI.

    Francesinha Bolebarra ...............  9.740.777,55
    Francesinha Bolepix .................    363.466,24
    Total arquivo de retorno ............ 10.103.006,88
    GAI — total baixado ................. 10.013.491,52
    GAI — baixas não efetivadas .........      1.921,07
    Bolepix não baixado .................      1.236,91

    A planilha feita à mão parava numa "DIVERGENCIA GAI" de
    4.213,39 sem saber de quem era — a anotação ao lado dizia
    "Estimativa GAI, é preciso especificar". É a cobrança do
    boleto 834036, que o GAI leu com valor 0,00.

Uso:
    python validar.py [pasta]

A pasta padrão é ../PROMPT. Os arquivos de 30/09 são procurados
em [pasta]/TESTE 2.
"""

from pathlib import Path
from decimal import Decimal
import sys

from conciliacao import conciliar, formatar_moeda


PASTA_PADRAO = Path(__file__).resolve().parent.parent / "PROMPT"

SUBPASTA_30 = "TESTE 2"


# ============================================================
# OS DOIS DIAS
# ============================================================

ARQUIVOS_28 = {
    "extrato": "28 09 2026 m.txt",
    "dcb": "DCB_1219001012_28092026_030540.txt",
    "francesinha": "Francesinha movimento bolepix - 28.09.26.xlt.xls",
}

ESPERADO_28 = {
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

ARQUIVOS_30 = {
    "dcb": "DCB_1219001012_30092026_031844.txt",
    "francesinha": "Francesinha movimento Bolepix 30 09 2026.xlt",
    "bolebarras": "Francesinha movimento Bolebarra 30 09 2026.xlt",
    "gai_lidas": "VALORES BAIXADOS NO GAI (AG. 121) R$ 10.013.491,52.pdf",
    "gai_nao_baixadas": "VALORES NÃO BAIXADOS NO GAI (AG. 121) R$ 1.921,07.pdf",
}

ESPERADO_30 = {
    "quantidade_bolebarras": 270,
    "total_bolebarras": Decimal("9740777.55"),
    "total_bolepix": Decimal("363466.24"),

    "quantidade_liquidados": 365,
    "total_retorno": Decimal("10103006.88"),

    "quantidade_casados_dcb": 270,
    "total_conferido_dcb": Decimal("9740777.55"),
    "quantidade_fora_da_francesinha": 95,
    "total_fora_da_francesinha": Decimal("362229.33"),

    "quantidade_boletos_gai": 336,
    "total_baixado_gai": Decimal("10013491.52"),
    "quantidade_nao_efetivadas": 157,
    "total_nao_efetivadas": Decimal("1921.07"),

    "quantidade_baixados": 241,
    "total_baixado": Decimal("9651262.19"),
    "quantidade_nao_baixados": 29,
    "total_nao_baixado": Decimal("89515.36"),

    # A cobrança que o GAI leu e não baixou: a divergência que a
    # planilha manual não conseguia nomear.
    "boleto_lido_sem_baixa": "834036",
    "total_lido_sem_baixa": Decimal("4213.39"),

    # As cobranças de outras gerências (GIR, GGR), que não
    # passam pelo GAI.
    "quantidade_ausentes": 28,
    "total_ausentes": Decimal("85301.97"),

    "quantidade_bolepix_nao_baixado": 2,
    "total_bolepix_nao_baixado": Decimal("1236.91"),

    # A recusa que trouxe dinheiro: segunda via do boleto da
    # parcela 19 da alienação 111268. O número do boleto
    # recusado não está na francesinha — o nome vem pela
    # alienação.
    "recusa_com_valor": Decimal("1921.07"),
    "alienacao_da_recusa": "111268",
    "nome_da_recusa": "EDLEUZA GONCALVES DOS REIS",
    "recusados_identificados": 38,
}


# ============================================================
# APOIO
# ============================================================

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


def localizar(pasta, arquivos):
    """
    Devolve os caminhos do dia, ou a lista do que falta.
    """

    faltando = [
        nome for nome in arquivos.values() if not (pasta / nome).exists()
    ]

    if faltando:
        return None, faltando

    return {
        rotulo: pasta / nome for rotulo, nome in arquivos.items()
    }, []


# ============================================================
# 28/09/2026 — O DIA FECHADO PELO EXTRATO
# ============================================================

def validar_28(pasta):
    """
    Confere o fechamento pelo extrato: PIX, cobrança, arquivo de
    retorno e a divergência explicada item a item.
    """

    caminhos, faltando = localizar(pasta, ARQUIVOS_28)

    if faltando:
        return None, faltando

    resultado = conciliar(**caminhos)

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
            ESPERADO_28["quantidade_pix"],
        ),
        comparar(
            "Total PIX (Bolepix)",
            extrato["total_pix"],
            ESPERADO_28["total_pix"],
            moeda=True,
        ),
        comparar(
            "Quantidade de créditos cobrança",
            extrato["quantidade_cobranca"],
            ESPERADO_28["quantidade_cobranca"],
        ),
        comparar(
            "Total cobrança (Bolebarras)",
            extrato["total_cobranca"],
            ESPERADO_28["total_cobranca"],
            moeda=True,
        ),
        comparar(
            "Total pagamentos QR Code + C.Barras",
            extrato["total_pagamentos"],
            ESPERADO_28["total_pagamentos"],
            moeda=True,
        ),
    ]

    print("\nFRANCESINHA BOLEPIX")
    verificacoes += [
        comparar(
            "Quantidade de recebimentos",
            francesinha["quantidade"],
            ESPERADO_28["quantidade_francesinha"],
        ),
        comparar(
            "Total",
            francesinha["total"],
            ESPERADO_28["total_francesinha"],
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
            ESPERADO_28["quantidade_liquidados"],
        ),
        comparar(
            "Total arquivo de retorno",
            dcb["total_liquidado"],
            ESPERADO_28["total_retorno"],
            moeda=True,
        ),
    ]

    print("\nDIVERGÊNCIA")
    verificacoes += [
        comparar(
            "Divergência",
            divergencia["divergencia"],
            ESPERADO_28["divergencia"],
            moeda=True,
        ),
        comparar(
            "Itens Bolepix não baixados",
            divergencia["quantidade_nao_baixados"],
            ESPERADO_28["quantidade_nao_baixados"],
        ),
        comparar(
            "Soma dos não baixados = divergência",
            divergencia["total_nao_baixados"],
            ESPERADO_28["divergencia"],
            moeda=True,
        ),
    ]

    print("\nBOLEPIX NÃO BAIXADO (detalhe)")

    for numero, registro in enumerate(
        resultado["bolepix_dcb"]["nao_baixados"], start=1
    ):
        print(
            f"  {numero}. {registro['nome'][:40]:<42}"
            f"{formatar_moeda(registro['valor']):>16}"
            f"   {registro['data']}"
        )

    print("\nPIX IDENTIFICADOS")

    verificacoes.append(
        comparar(
            "PIX do extrato com pagador identificado",
            resultado["pix_detalhado"]["quantidade_identificados"],
            ESPERADO_28["quantidade_pix"],
        )
    )

    return verificacoes, []


# ============================================================
# 30/09/2026 — O DIA CONFERIDO COM O GAI, SEM EXTRATO
# ============================================================

def validar_30(pasta):
    """
    Confere a Bolebarra contra o DCB e contra os relatórios do
    GAI, sem extrato nenhum.
    """

    caminhos, faltando = localizar(pasta, ARQUIVOS_30)

    if faltando:
        return None, faltando

    resultado = conciliar(**caminhos)

    bolebarras = resultado["bolebarras"]
    dcb = resultado["dcb"]
    lidas = resultado["gai_lidas"]
    nao_efetivadas = resultado["gai_nao_baixadas"]
    com_dcb = resultado["bolebarras_dcb"]
    com_gai = resultado["bolebarras_gai"]

    verificacoes = []

    print("FRANCESINHAS")
    verificacoes += [
        comparar(
            "Cobranças na Bolebarra",
            bolebarras["quantidade"],
            ESPERADO_30["quantidade_bolebarras"],
        ),
        comparar(
            "Total Bolebarra",
            bolebarras["total"],
            ESPERADO_30["total_bolebarras"],
            moeda=True,
        ),
        comparar(
            "Toda a Bolebarra tem nosso número",
            bolebarras["tem_nosso_numero"],
            True,
        ),
        comparar(
            "Total Bolepix",
            resultado["francesinha"]["total"],
            ESPERADO_30["total_bolepix"],
            moeda=True,
        ),
    ]

    print("\nDCB - ARQUIVO DE RETORNO")
    verificacoes += [
        comparar(
            "Títulos liquidados",
            dcb["quantidade_liquidados"],
            ESPERADO_30["quantidade_liquidados"],
        ),
        comparar(
            "Total arquivo de retorno",
            dcb["total_liquidado"],
            ESPERADO_30["total_retorno"],
            moeda=True,
        ),
    ]

    print("\nRELATÓRIOS DO GAI")
    verificacoes += [
        comparar(
            "Boletos baixados no GAI",
            lidas["quantidade_boletos"],
            ESPERADO_30["quantidade_boletos_gai"],
        ),
        comparar(
            "Total baixado no GAI",
            lidas["total_baixado"],
            ESPERADO_30["total_baixado_gai"],
            moeda=True,
        ),
        comparar(
            "Total = rodapé do próprio relatório",
            lidas["total_baixado"],
            lidas["rodape"]["total_recebido"],
            moeda=True,
        ),
        comparar(
            "Registros de baixa não efetivada",
            nao_efetivadas["quantidade"],
            ESPERADO_30["quantidade_nao_efetivadas"],
        ),
        comparar(
            "Total não efetivado",
            nao_efetivadas["total"],
            ESPERADO_30["total_nao_efetivadas"],
            moeda=True,
        ),
        comparar(
            "Total = rodapé do próprio relatório",
            nao_efetivadas["total"],
            nao_efetivadas["rodape"]["total"],
            moeda=True,
        ),
    ]

    recusa = (nao_efetivadas["com_valor"] or [{}])[0]

    print("\nRECUSAS COM NOME")
    verificacoes += [
        comparar(
            "Recusas com dinheiro",
            nao_efetivadas["quantidade_com_valor"],
            1,
        ),
        comparar(
            "Valor da recusa",
            recusa.get("total_pago"),
            ESPERADO_30["recusa_com_valor"],
            moeda=True,
        ),
        comparar(
            "Alienação da recusa",
            recusa.get("alienacao"),
            ESPERADO_30["alienacao_da_recusa"],
        ),
        comparar(
            "Nome do pagador (vem pela alienação)",
            recusa.get("nome"),
            ESPERADO_30["nome_da_recusa"],
        ),
        comparar(
            "Recusados com nome na francesinha",
            nao_efetivadas["quantidade_identificados"],
            ESPERADO_30["recusados_identificados"],
        ),
    ]

    print("\nBOLEBARRA x DCB")
    verificacoes += [
        comparar(
            "Cobranças casadas pelo nosso número",
            com_dcb["quantidade_casados"],
            ESPERADO_30["quantidade_casados_dcb"],
        ),
        comparar(
            "Total conferido",
            com_dcb["total_conferido"],
            ESPERADO_30["total_conferido_dcb"],
            moeda=True,
        ),
        comparar("Nenhum desencontro", com_dcb["confere"], True),
        comparar(
            "Liquidados no DCB fora da Bolebarra",
            com_dcb["quantidade_fora_da_francesinha"],
            ESPERADO_30["quantidade_fora_da_francesinha"],
        ),
        comparar(
            "Total desses liquidados (é o Bolepix)",
            com_dcb["total_fora_da_francesinha"],
            ESPERADO_30["total_fora_da_francesinha"],
            moeda=True,
        ),
    ]

    por_situacao = {
        situacao["situacao"]: situacao for situacao in com_gai["situacoes"]
    }

    lidos_sem_baixa = [
        linha
        for linha in com_gai["nao_baixados"]
        if linha["situacao"] == "lido_sem_baixa"
    ]

    print("\nBOLEBARRA x GAI")
    verificacoes += [
        comparar(
            "Cobranças baixadas no GAI",
            com_gai["quantidade_baixados"],
            ESPERADO_30["quantidade_baixados"],
        ),
        comparar(
            "Total baixado",
            com_gai["total_baixado"],
            ESPERADO_30["total_baixado"],
            moeda=True,
        ),
        comparar(
            "Nenhuma baixada por valor diferente",
            com_gai["quantidade_valor_divergente"],
            0,
        ),
        comparar(
            "Cobranças NÃO baixadas no GAI",
            com_gai["quantidade_nao_baixados"],
            ESPERADO_30["quantidade_nao_baixados"],
        ),
        comparar(
            "Total não baixado",
            com_gai["total_nao_baixado"],
            ESPERADO_30["total_nao_baixado"],
            moeda=True,
        ),
        comparar(
            "Lidas pelo GAI com valor 0,00",
            por_situacao["lido_sem_baixa"]["total"],
            ESPERADO_30["total_lido_sem_baixa"],
            moeda=True,
        ),
        comparar(
            "Boleto da cobrança lida sem baixa",
            lidos_sem_baixa[0]["titulo"]["boleto"] if lidos_sem_baixa else "",
            ESPERADO_30["boleto_lido_sem_baixa"],
        ),
        comparar(
            "Fora dos relatórios do GAI (outras gerências)",
            por_situacao["ausente"]["quantidade"],
            ESPERADO_30["quantidade_ausentes"],
        ),
        comparar(
            "Total dessas cobranças",
            por_situacao["ausente"]["total"],
            ESPERADO_30["total_ausentes"],
            moeda=True,
        ),
        comparar(
            "Baixados no GAI fora da Bolebarra",
            com_gai["total_fora_da_francesinha"],
            ESPERADO_30["total_fora_da_francesinha"],
            moeda=True,
        ),
    ]

    print("\nBOLEPIX NÃO BAIXADO (sem extrato)")
    verificacoes += [
        comparar(
            "Itens sem baixa no arquivo de retorno",
            resultado["bolepix_dcb"]["quantidade_nao_baixados"],
            ESPERADO_30["quantidade_bolepix_nao_baixado"],
        ),
        comparar(
            "Total",
            resultado["bolepix_dcb"]["total_nao_baixados"],
            ESPERADO_30["total_bolepix_nao_baixado"],
            moeda=True,
        ),
        comparar(
            "Mesmo total pelos relatórios do GAI",
            resultado["bolepix_gai"]["total_nao_baixados"],
            ESPERADO_30["total_bolepix_nao_baixado"],
            moeda=True,
        ),
    ]

    print("\nA DIVERGÊNCIA QUE A PLANILHA MANUAL NÃO NOMEAVA")

    for linha in lidos_sem_baixa:
        print(
            f"  boleto {linha['titulo']['boleto']}"
            f"   alienação {linha['alienacao']}"
            f"   {formatar_moeda(linha['titulo']['valor']):>14}"
            f"   {linha['titulo']['nome'][:34]}"
        )
        print(f"    {linha['motivo'][:66]}")

    return verificacoes, []


# ============================================================
# EXECUÇÃO
# ============================================================

CASOS = (
    ("28/09/2026 — fechamento pelo extrato", validar_28, ""),
    ("30/09/2026 — conferência com o GAI", validar_30, SUBPASTA_30),
)


def main(pasta=PASTA_PADRAO):
    pasta = Path(pasta)

    verificacoes = []
    pulados = []

    for titulo, validador, subpasta in CASOS:

        onde = pasta / subpasta if subpasta else pasta

        print(f"\n{'=' * 70}")
        print(titulo)
        print(f"{onde}")
        print("=" * 70)

        if not onde.is_dir():
            print(f"  pasta não encontrada — caso pulado")
            pulados.append(titulo)
            continue

        resultado, faltando = validador(onde)

        if faltando:
            print("  arquivos não encontrados — caso pulado:")

            for nome in faltando:
                print("   -", nome)

            pulados.append(titulo)
            continue

        verificacoes += resultado

    total = len(verificacoes)
    ok = sum(1 for v in verificacoes if v)

    print(f"\n{'=' * 70}")
    print(f"{ok}/{total} verificações passaram")

    for titulo in pulados:
        print(f"  pulado: {titulo}")

    print("=" * 70)

    if not total:
        return 2

    return 0 if ok == total else 1


if __name__ == "__main__":
    pasta = sys.argv[1] if len(sys.argv) > 1 else PASTA_PADRAO

    sys.exit(main(pasta))
