"""
Leitura de planilhas sem pandas.

O programa precisa apenas percorrer células, então usa xlrd
(.xls) e openpyxl (.xlsx) diretamente. Os dois são bibliotecas
Python puras, o que permite rodar o mesmo código no navegador
via Pyodide, sem baixar pandas e numpy.

Cada planilha é devolvida como:

    {nome_da_aba: [[celula, celula, ...], ...]}

onde cada célula é str, float, datetime ou None.
"""

from datetime import datetime
from io import BytesIO
from pathlib import Path


ASSINATURA_XLS = b"\xd0\xcf\x11\xe0"
ASSINATURA_XLSX = b"PK\x03\x04"

EXTENSOES_PLANILHA = (".xls", ".xlsx", ".xlsm", ".xlt", ".xltx")


def e_caminho(origem):
    """
    Diferencia um caminho de arquivo de um conteúdo de texto.
    """

    if isinstance(origem, Path):
        return True

    if isinstance(origem, str):
        return "\n" not in origem and len(origem) < 1000

    return False


def para_bytes(origem):
    """
    Normaliza qualquer origem suportada para bytes.
    """

    if isinstance(origem, (bytes, bytearray, memoryview)):
        return bytes(origem)

    if hasattr(origem, "read"):
        conteudo = origem.read()

        if isinstance(conteudo, str):
            return conteudo.encode("latin-1", errors="ignore")

        return bytes(conteudo)

    return Path(origem).read_bytes()


def para_texto(origem, encoding="latin-1"):
    """
    Normaliza qualquer origem suportada para texto.

    O extrato do BRB e o DCB vêm em latin-1, mas arquivos
    salvos novamente podem vir em UTF-8.
    """

    if isinstance(origem, str) and not e_caminho(origem):
        return origem

    dados = para_bytes(origem)

    for tentativa in ("utf-8-sig", "utf-8", encoding):
        try:
            return dados.decode(tentativa)
        except UnicodeDecodeError:
            continue

    return dados.decode(encoding, errors="ignore")


def nome_arquivo(origem, padrao=""):
    """
    Descobre o nome do arquivo, usado para escolher o leitor.
    """

    if e_caminho(origem):
        return str(origem)

    for atributo in ("filename", "name"):
        valor = getattr(origem, atributo, None)

        if isinstance(valor, str) and valor:
            return valor

    return padrao


def e_planilha(origem, nome=""):
    """
    Detecta planilha pela extensão ou pela assinatura binária.

    Os relatórios do BRB chegam com nome "....xlt.xls", mas o
    conteúdo é um .xls de verdade.
    """

    extensao = Path(nome or nome_arquivo(origem)).suffix.lower()

    if extensao in EXTENSOES_PLANILHA:
        return True

    inicio = para_bytes(origem)[:8]

    return (
        inicio.startswith(ASSINATURA_XLS)
        or inicio.startswith(ASSINATURA_XLSX)
    )


# ============================================================
# LEITORES
# ============================================================

def _ler_xls(dados):
    """
    Lê um .xls (BIFF) com xlrd.
    """

    import xlrd

    livro = xlrd.open_workbook(file_contents=dados)

    abas = {}

    for aba in livro.sheets():

        linhas = []

        for indice_linha in range(aba.nrows):

            linha = []

            for indice_coluna in range(aba.ncols):

                tipo = aba.cell_type(indice_linha, indice_coluna)
                valor = aba.cell_value(indice_linha, indice_coluna)

                if tipo == xlrd.XL_CELL_EMPTY or tipo == xlrd.XL_CELL_BLANK:
                    valor = None

                elif tipo == xlrd.XL_CELL_DATE:
                    try:
                        valor = xlrd.xldate_as_datetime(
                            valor, livro.datemode
                        )
                    except Exception:
                        valor = None

                elif tipo == xlrd.XL_CELL_TEXT:
                    valor = valor.strip() or None

                elif tipo == xlrd.XL_CELL_ERROR:
                    valor = None

                linha.append(valor)

            linhas.append(linha)

        abas[aba.name] = linhas

    return abas


def _ler_xlsx(dados):
    """
    Lê um .xlsx com openpyxl.
    """

    from openpyxl import load_workbook

    livro = load_workbook(
        BytesIO(dados),
        read_only=True,
        data_only=True,
    )

    abas = {}

    for aba in livro.worksheets:

        linhas = []

        for linha in aba.iter_rows(values_only=True):

            celulas = []

            for valor in linha:

                if isinstance(valor, str):
                    valor = valor.strip() or None

                celulas.append(valor)

            linhas.append(celulas)

        abas[aba.title] = linhas

    livro.close()

    return abas


def ler_planilha(origem, nome=""):
    """
    Lê todas as abas de uma planilha .xls ou .xlsx.

    Devolve {nome_da_aba: [[celula, ...], ...]}.
    """

    dados = para_bytes(origem)

    if dados.startswith(ASSINATURA_XLSX):
        return _ler_xlsx(dados)

    if dados.startswith(ASSINATURA_XLS):
        return _ler_xls(dados)

    # Sem assinatura reconhecida: decide pela extensão.
    extensao = Path(nome or nome_arquivo(origem)).suffix.lower()

    if extensao in (".xlsx", ".xlsm", ".xltx"):
        return _ler_xlsx(dados)

    return _ler_xls(dados)


# ============================================================
# APOIO
# ============================================================

def celula(linha, indice):
    """
    Lê uma célula da linha tolerando linhas curtas.
    """

    if indice is None or indice >= len(linha):
        return None

    return linha[indice]


def texto_da_celula(valor):
    """
    Converte uma célula em texto limpo.
    """

    if valor is None:
        return ""

    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y %H:%M:%S")

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor).strip()
