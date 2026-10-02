"""
Servidor local opcional da conciliação financeira.

A versão publicada no GitHub Pages (pasta docs/) roda tudo no
navegador e não precisa deste servidor. Ele existe para quem
preferir rodar a aplicação na própria máquina ou na intranet:

    python -m uvicorn app:app --reload

e abrir http://127.0.0.1:8000

Os dois caminhos usam exatamente o mesmo núcleo (conciliacao.py),
então o resultado é idêntico.
"""

from collections import OrderedDict
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from conciliacao import conciliar
from exportar import montar_planilha


PASTA = Path(__file__).resolve().parent
ESTATICOS = PASTA / "docs"

# Limite de upload por arquivo. O DCB real tem ~850 KB; o teto
# alto cobre arquivos de fechamento mensal.
TAMANHO_MAXIMO = 64 * 1024 * 1024

# Resultados recentes, guardados apenas para permitir a
# exportação em Excel sem reenviar os arquivos. Ficam na memória
# do processo, o que atende o uso local / intranet.
RESULTADOS = OrderedDict()
MAXIMO_RESULTADOS = 20

TIPO_XLSX = (
    "application/vnd.openxmlformats-officedocument"
    ".spreadsheetml.sheet"
)


app = FastAPI(
    title="Conciliação Financeira",
    description="Conciliação de extrato BRB, Bolepix e DCB",
    version="1.0.0",
)


# ============================================================
# SERIALIZAÇÃO
# ============================================================

def limpar(objeto):
    """
    Converte a estrutura da conciliação para algo que o JSON
    aceite: Decimal vira float, datetime vira texto.
    """

    if isinstance(objeto, Decimal):
        return float(objeto)

    if isinstance(objeto, dict):
        return {chave: limpar(valor) for chave, valor in objeto.items()}

    if isinstance(objeto, (list, tuple)):
        return [limpar(item) for item in objeto]

    if isinstance(objeto, datetime):
        return objeto.strftime("%d/%m/%Y %H:%M:%S")

    return objeto


def resposta_para_tela(resultado):
    """
    Monta a versão enxuta do resultado que vai para o navegador.

    O DCB tem milhares de títulos; a tela recebe só os totais e
    as ocorrências. O resultado completo fica no servidor para a
    exportação em Excel.
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

    return limpar(enxuto)


def guardar(resultado):
    """
    Guarda o resultado completo e devolve o identificador usado
    pela exportação.
    """

    identificador = uuid4().hex

    RESULTADOS[identificador] = resultado

    while len(RESULTADOS) > MAXIMO_RESULTADOS:
        RESULTADOS.popitem(last=False)

    return identificador


# ============================================================
# UPLOAD
# ============================================================

class ArquivoEnviado:
    """
    Embala o upload com nome e conteúdo, que é o que as funções
    de leitura do núcleo esperam.
    """

    def __init__(self, nome, conteudo):
        self.filename = nome
        self._conteudo = conteudo

    def read(self):
        return self._conteudo


async def receber(upload, rotulo):
    """
    Lê o upload na memória, validando presença e tamanho.
    """

    if upload is None or not upload.filename:
        return None

    conteudo = await upload.read()

    if not conteudo:
        raise HTTPException(
            status_code=400,
            detail=f"O arquivo de {rotulo} está vazio.",
        )

    if len(conteudo) > TAMANHO_MAXIMO:
        raise HTTPException(
            status_code=413,
            detail=(
                f"O arquivo de {rotulo} passa do limite de "
                f"{TAMANHO_MAXIMO // (1024 * 1024)} MB."
            ),
        )

    return ArquivoEnviado(upload.filename, conteudo)


# ============================================================
# ROTAS
# ============================================================

@app.post("/api/conciliar")
async def api_conciliar(
    extrato: UploadFile = File(None),
    francesinha: UploadFile = File(None),
    dcb: UploadFile = File(None),
    bolebarras: UploadFile = File(None),
):
    """
    Recebe os arquivos do dia e devolve a conciliação.
    """

    arquivos = {
        "extrato": await receber(extrato, "extrato"),
        "francesinha": await receber(francesinha, "francesinha Bolepix"),
        "dcb": await receber(dcb, "DCB"),
        "bolebarras": await receber(bolebarras, "francesinha Bolebarras"),
    }

    if not any(arquivos.values()):
        raise HTTPException(
            status_code=400,
            detail="Envie pelo menos um arquivo para conciliar.",
        )

    try:
        resultado = conciliar(**arquivos)
    except Exception as erro:
        raise HTTPException(
            status_code=422,
            detail=(
                "Não foi possível processar os arquivos: "
                f"{type(erro).__name__}: {erro}"
            ),
        )

    resposta = resposta_para_tela(resultado)
    resposta["id"] = guardar(resultado)

    return resposta


@app.get("/api/exportar/{identificador}")
def api_exportar(identificador: str):
    """
    Devolve a conciliação em Excel, com uma aba por bloco.
    """

    resultado = RESULTADOS.get(identificador)

    if resultado is None:
        raise HTTPException(
            status_code=404,
            detail=(
                "Esta conciliação não está mais disponível. "
                "Processe os arquivos novamente."
            ),
        )

    conteudo = montar_planilha(resultado)

    hoje = datetime.now().strftime("%Y-%m-%d_%H%M")

    return StreamingResponse(
        BytesIO(conteudo),
        media_type=TIPO_XLSX,
        headers={
            "Content-Disposition": (
                f'attachment; filename="conciliacao_{hoje}.xlsx"'
            )
        },
    )


@app.get("/api/saude")
def api_saude():
    """
    Checagem simples para saber se o servidor está no ar.
    """

    return {
        "status": "ok",
        "resultados_em_memoria": len(RESULTADOS),
    }


# ============================================================
# FRONT-END
# ============================================================

@app.get("/")
def pagina_inicial():
    """
    Entrega a mesma interface publicada no GitHub Pages.
    """

    return FileResponse(ESTATICOS / "index.html")


app.mount("/", StaticFiles(directory=ESTATICOS, html=True), name="docs")
