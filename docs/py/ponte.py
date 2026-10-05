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

from conciliacao import conciliar
from exportar import montar_planilha


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


def _titulo_para_tela(linha):
    """
    Achata uma cobrança conferida com o GAI numa linha de
    tabela.
    """

    titulo = linha["titulo"]

    return {
        "nosso_numero": titulo["nosso_numero"],
        "boleto": titulo["boleto"],
        "alienacao": linha["alienacao"],
        "nome": titulo["nome"],
        "data": titulo["data"],
        "valor": titulo["valor"],
        "valor_gai": linha["valor_gai"],
        "situacao": linha["situacao"],
        "motivo": linha["motivo"],
    }


def para_tela(resultado):
    """
    Monta a versão enxuta que vai para a tela.

    O DCB tem milhares de títulos: a tela recebe só os totais, as
    ocorrências e as listas que o usuário precisa ler item a
    item. O detalhe completo vai para o Excel.
    """

    enxuto = {
        "avisos": resultado["avisos"],
        "sistemas": resultado.get("sistemas", []),
        "benner": resultado.get("benner"),
        "gai_benner": resultado.get("gai_benner"),
        "bolebarras_sistemas": resultado.get("bolebarras_sistemas"),
        "extrato": None,
        "francesinha": None,
        "bolebarras": None,
        "dcb": None,
        "gai_lidas": None,
        "gai_nao_baixadas": None,
        "pix_detalhado": resultado["pix_detalhado"],
        "bolepix_dcb": resultado["bolepix_dcb"],
        "bolepix_gai": None,
        "bolebarras_dcb": None,
        "bolebarras_gai": None,
        "divergencia": resultado["divergencia"],
    }

    lidas = resultado["gai_lidas"]

    if lidas is not None:
        enxuto["gai_lidas"] = {
            "lote": lidas["lote"],
            "arquivo": lidas["arquivo"],
            "quantidade_boletos": lidas["quantidade_boletos"],
            "quantidade_parcelas": lidas["quantidade_parcelas"],
            "quantidade_lidos_sem_baixa": lidas["quantidade_lidos_sem_baixa"],
            "total_baixado": lidas["total_baixado"],
        }

    nao_baixadas = resultado["gai_nao_baixadas"]

    if nao_baixadas is not None:
        enxuto["gai_nao_baixadas"] = {
            "quantidade": nao_baixadas["quantidade"],
            "quantidade_com_valor": nao_baixadas["quantidade_com_valor"],
            "quantidade_identificados": nao_baixadas[
                "quantidade_identificados"
            ],
            "total": nao_baixadas["total"],
            "motivos": nao_baixadas["motivos"],

            # A tela detalha as recusas com valor positivo.
            # Todos os registros permanecem no resumo por motivo
            # e na exportação completa para Excel.
            "registros": [
                {
                    chave: registro[chave]
                    for chave in (
                        "alienacao",
                        "boleto",
                        "parcela",
                        "nome",
                        "nosso_numero",
                        "data_pagamento",
                        "total_pago",
                        "motivo",
                    )
                }
                for registro in sorted(
                    nao_baixadas["com_valor"],
                    key=lambda r: (-r["total_pago"], r["alienacao"]),
                )
            ],
        }

    conferencia = resultado["bolebarras_dcb"]

    if conferencia is not None:
        enxuto["bolebarras_dcb"] = {
            chave: conferencia[chave]
            for chave in (
                "quantidade_casados",
                "quantidade_valor_divergente",
                "quantidade_sem_baixa",
                "quantidade_sem_nosso_numero",
                "quantidade_fora_da_francesinha",
                "total_francesinha",
                "total_conferido",
                "total_sem_baixa",
                "total_fora_da_francesinha",
                "confere",
            )
        }

        # As duas listas que o usuário precisa ler item a item.
        enxuto["bolebarras_dcb"]["sem_baixa"] = [
            {
                "nosso_numero": item["titulo"]["nosso_numero"],
                "nome": item["titulo"]["nome"],
                "data": item["titulo"]["data"],
                "valor": item["titulo"]["valor"],
                "no_arquivo": item["no_arquivo"],
                "ocorrencias": ", ".join(item["ocorrencias"]),
            }
            for item in conferencia["sem_baixa"]
        ]

        enxuto["bolebarras_dcb"]["valor_divergente"] = [
            {
                "nosso_numero": par["titulo"]["nosso_numero"],
                "nome": par["titulo"]["nome"],
                "valor": par["titulo"]["valor"],
                "valor_dcb": par["registro"]["valor"],
                "diferenca": par["diferenca"],
            }
            for par in conferencia["valor_divergente"]
        ]

    conferencia = resultado["bolebarras_gai"]

    if conferencia is not None:
        enxuto["bolebarras_gai"] = {
            chave: conferencia[chave]
            for chave in (
                "quantidade_baixados",
                "quantidade_valor_divergente",
                "quantidade_nao_baixados",
                "quantidade_fora_da_francesinha",
                "total_francesinha",
                "total_baixado",
                "total_nao_baixado",
                "total_fora_da_francesinha",
                "situacoes",
                "confere",
            )
        }

        enxuto["bolebarras_gai"]["nao_baixados"] = [
            _titulo_para_tela(linha)
            for linha in conferencia["nao_baixados"]
        ]

        enxuto["bolebarras_gai"]["valor_divergente"] = [
            _titulo_para_tela(linha)
            for linha in conferencia["valor_divergente"]
        ]

    conferencia = resultado["bolepix_gai"]

    if conferencia is not None:
        enxuto["bolepix_gai"] = {
            chave: conferencia[chave]
            for chave in (
                "total_francesinha",
                "total_no_gai",
                "nao_baixados",
                "quantidade_nao_baixados",
                "total_nao_baixados",
                "quantidade_boletos_sem_pix",
                "confere",
            )
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
            "layout": dados["layout"],
            "quantidade": dados["quantidade"],
            "total": dados["total"],
            "tem_nosso_numero": dados["tem_nosso_numero"],
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

    for rotulo in (
        "extrato",
        "francesinha",
        "dcb",
        "bolebarras",
        "gai_lidas",
        "gai_nao_baixadas",
        "gir",
        "ggr",
        "gop",
        "benner",
    ):

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
        para_tela(resultado),
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
