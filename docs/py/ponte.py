"""
Ponte entre o navegador e o núcleo da conciliação.

Carregado pelo Pyodide (ver docs/app.js). O JavaScript entrega
os arquivos como bytes e recebe o resultado como texto JSON,
evitando conversões automáticas entre Python e JavaScript.

O resultado completo fica guardado aqui para que a exportação
em Excel não precise reprocessar os arquivos.
"""

from datetime import datetime
from decimal import Decimal
import json

from conciliacao_terracap.conciliacao import conciliar
from conciliacao_terracap.exportar import montar_planilha


# Último resultado processado, usado pela exportação.
ULTIMO = {"resultado": None}


class Arquivo:
    """
    Embala o conteúdo enviado pelo navegador no formato que as
    funções de leitura esperam (nome + .read()).
    """

    def __init__(self, nome, conteudo):
        self.filename = nome or ""
        self._conteudo = bytes(conteudo)

    def read(self):
        return self._conteudo


def _json_seguro(objeto):
    """
    Converte os tipos que o json não conhece.
    """

    if isinstance(objeto, Decimal):
        return float(objeto)

    if isinstance(objeto, datetime):
        return objeto.strftime("%d/%m/%Y %H:%M:%S")

    return str(objeto)


def _para_tela(resultado):
    """
    Monta a versão enxuta que vai para a tela.

    O DCB tem milhares de títulos: a tela recebe só os totais e
    as ocorrências, e o detalhe completo vai para o Excel.
    """

    enxuto = {
        "avisos": resultado["avisos"],
        "extrato": None,
        "francesinha": None,
        "bolebarras": None,
        "dcb": None,
        "pix_detalhado": resultado["pix_detalhado"],
        "divergencia": resultado["divergencia"],
    }

    extrato = resultado["extrato"]

    if extrato is not None:
        enxuto["extrato"] = {
            chave: extrato[chave]
            for chave in (
                "quantidade_pix",
                "quantidade_cobranca",
                "quantidade_outros",
                "total_pix",
                "total_cobranca",
                "total_outros",
                "total_pagamentos",
            )
        }

        enxuto["extrato"]["cobranca"] = extrato["cobranca"]

    for bloco in ("francesinha", "bolebarras"):

        dados = resultado[bloco]

        if dados is None:
            continue

        enxuto[bloco] = {
            "formato": dados["formato"],
            "quantidade": dados["quantidade"],
            "total": dados["total"],
            "conferencia_extrato": dados.get("conferencia_extrato"),
        }

    dcb = resultado["dcb"]

    if dcb is not None:
        enxuto["dcb"] = {
            "formato": dcb["formato"],
            "quantidade_registros": dcb["quantidade_registros"],
            "quantidade_liquidados": dcb["quantidade_liquidados"],
            "total_liquidado": dcb["total_liquidado"],
            "ocorrencias": dcb["ocorrencias"],
        }

    return enxuto


def processar(arquivos):
    """
    Recebe {rotulo: bytes ou None} e devolve o resultado em JSON.

    Em caso de erro de leitura, devolve {"erro": "..."} em vez de
    levantar exceção, para a tela mostrar uma mensagem clara.
    """

    entrada = {}

    for rotulo in ("extrato", "francesinha", "dcb", "bolebarras"):

        item = arquivos.get(rotulo)

        if not item:
            entrada[rotulo] = None
            continue

        entrada[rotulo] = Arquivo(item.get("nome"), item.get("dados"))

    if not any(entrada.values()):
        return json.dumps(
            {"erro": "Envie pelo menos um arquivo para conciliar."}
        )

    try:
        resultado = conciliar(**entrada)
    except Exception as erro:
        return json.dumps({
            "erro": (
                "Não foi possível processar os arquivos: "
                f"{type(erro).__name__}: {erro}"
            )
        })

    ULTIMO["resultado"] = resultado

    return json.dumps(
        _para_tela(resultado),
        default=_json_seguro,
        ensure_ascii=False,
    )


def exportar():
    """
    Devolve os bytes do .xlsx da última conciliação processada.
    """

    resultado = ULTIMO["resultado"]

    if resultado is None:
        raise RuntimeError("Nenhuma conciliação foi processada ainda.")

    return montar_planilha(resultado)
