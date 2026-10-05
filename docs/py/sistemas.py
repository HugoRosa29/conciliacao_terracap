"""Leitura das tabelas GIR/GGR/GOP e cruzamento por boleto e data.

As seções são reconhecidas pelo título, inclusive em cópias sem marca-texto.
GIR: Relação de Parcelas Lidas. GGR/GOP: Baixas Efetivadas Normalmente,
com a Relação de Parcelas Lidas usada apenas para obter o número do boleto.
Usa somente pypdf para funcionar também no navegador via Pyodide.
"""
from collections import defaultdict
from decimal import Decimal
from io import BytesIO
import re

from gai import _sem_acento as normalizar, _valor_brasileiro
from planilhas import para_bytes

ZERO = Decimal('0.00')
DINHEIRO = r'-?[\d.]+,\d{2}'
DATA = r'\d{2}/\d{2}/\d{4}'


def identificar_sistema(texto):
    texto = normalizar(texto)
    if 'gir - gestao de imoveis rurais' in texto:
        return 'GIR'
    if 'do ggr' in texto:
        return 'GGR'
    if 'controle de operacao' in texto:
        return 'GOP'
    return ''


def _pedacos(pagina):
    pedacos = []
    def guardar(texto, cm, tm, fonte, tamanho):
        if texto.strip():
            x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
            y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
            pedacos.append((x, y, texto.strip()))
    pagina.extract_text(visitor_text=guardar)
    return pedacos


def _celula(pedacos, inicio, fim):
    return ' '.join(t for x, y, t in pedacos if inicio <= x < fim)


def _linhas_tabela(pagina, sistema):
    # Os PDFs imprimem cada célula por completo, inclusive texto com várias
    # linhas, antes da célula seguinte. Preservar essa ordem evita misturar
    # nomes longos de linhas de alturas diferentes.
    atual = []
    for x, y, texto in _pedacos(pagina):
        inicio = (x < 105 and re.fullmatch(r'\d+/\d{4}', texto))
        if sistema == 'GIR':
            inicio = x < 100 and re.fullmatch(r'\d{5}-\d+/\d{4}', texto)
        if inicio:
            if atual:
                yield atual
            atual = []
        if normalizar(texto).startswith('total pago:'):
            if atual:
                yield atual
            atual = []
            break
        if atual or inicio:
            atual.append((x, y, texto))
    if atual:
        yield atual


def ler_sistema(origem, esperado=None):
    from pypdf import PdfReader
    leitor = PdfReader(BytesIO(para_bytes(origem)))
    textos = [p.extract_text() or '' for p in leitor.pages]
    sistema = identificar_sistema(' '.join(textos))
    if not sistema or (esperado and sistema != esperado.upper()):
        raise ValueError(f'PDF não corresponde ao relatório {esperado or "GIR/GGR/GOP"}.')
    secao = ('relacao de parcelas lidas' if sistema == 'GIR'
             else 'baixas de pagamentos efetivadas normalmente')
    paginas = [i for i, t in enumerate(textos) if secao in normalizar(t)]
    if not paginas:
        raise ValueError(f'{sistema}: tabela {secao} não encontrada.')

    # A tabela de baixas GGR/GOP omite Nosso Número. O documento liga
    # essa tabela à relação de parcelas lidas dentro do mesmo relatório.
    boletos = defaultdict(set)
    if sistema != 'GIR':
        for texto in textos:
            if 'relacao de parcelas lidas' not in normalizar(texto):
                continue
            for linha in texto.splitlines():
                m = re.match(rf'{DATA}\s+\d{{2}}:\d{{2}}:\d{{2}}\s+(\d+/\d{{4}})\s+(.*)', linha)
                if not m:
                    continue
                partes = m[2].split()
                posicao = 1 if sistema == 'GOP' else 0
                if len(partes) > posicao and partes[posicao].isdigit():
                    boletos[m[1]].add(partes[posicao].zfill(6))

    registros, totais, avisos = [], [], []
    for i in paginas:
        totais.extend(_valor_brasileiro(v) for v in re.findall(
            rf'Total Pago:\s*({DINHEIRO})', textos[i], re.I))
        for pedacos in _linhas_tabela(leitor.pages[i], sistema):
            if sistema == 'GIR':
                registro = dict(documento=_celula(pedacos, 0, 100),
                    contrato=_celula(pedacos, 100, 180),
                    boleto=_celula(pedacos, 180, 257).zfill(6),
                    nome=_celula(pedacos, 257, 338).strip('- '),
                    data_pagamento=_celula(pedacos, 420, 500),
                    valor=_valor_brasileiro(_celula(pedacos, 500, 600)))
            else:
                documento = _celula(pedacos, 0, 105)
                cliente = _celula(pedacos, 157, 214)
                nome = re.search(r'Cliente:\s*(.*?)(?:Processo:|Alienação:|$)', cliente, re.I)
                candidatos = boletos[documento]
                registro = dict(documento=documento, contrato='',
                    boleto=next(iter(candidatos)) if len(candidatos) == 1 else '',
                    nome=nome[1].strip() if nome else '',
                    data_pagamento=_celula(pedacos, 384, 439),
                    valor=_valor_brasileiro(_celula(pedacos, 498, 600)))
            if registro['valor'] is None:
                raise ValueError(f'{sistema}, página {i+1}: valor ilegível em {registro["documento"]}.')
            if registro['valor'] <= ZERO:
                continue
            if not re.fullmatch(DATA, registro['data_pagamento']):
                raise ValueError(f'{sistema}: data de pagamento inválida em {registro["documento"]}.')
            if not registro['boleto']:
                avisos.append(f'{sistema}: boleto ausente ou ambíguo para {registro["documento"]}.')
            registro.update(sistema=sistema, pagina=i+1)
            registros.append(registro)
    total = sum((r['valor'] for r in registros), ZERO)
    total_impresso = totais[-1] if totais else None
    if total_impresso is None or total != total_impresso:
        avisos.append(f'{sistema}: soma dos pagamentos ({total}) não confere com o total impresso ({total_impresso}).')
    return dict(sistema=sistema, secao=secao, paginas=[i+1 for i in paginas],
                registros=registros, quantidade=len(registros), total=total,
                total_impresso=total_impresso, avisos=avisos)


def _boleto(nosso_numero):
    numero = re.sub(r'\D', '', str(nosso_numero or ''))
    return numero[1:7] if len(numero) == 12 and numero[7:10] == '070' else ''


def _indice(registros, campo_data):
    indice = defaultdict(list)
    for r in registros:
        numero = _boleto(r.get('nosso_numero'))
        if numero:
            indice[(numero, r.get(campo_data, ''))].append(r)
    return indice


def _conferir(grupo, indice, valor):
    if indice is None:
        return 'Não enviado', None, []
    candidatos = indice.get(grupo, [])
    if not candidatos:
        return 'Não encontrado', None, []
    # Duas carteiras podem compartilhar os seis dígitos do boleto.
    if len({r['nosso_numero'] for r in candidatos}) > 1:
        return 'Identificação ambígua', None, candidatos
    total = sum((r['valor'] for r in candidatos), ZERO)
    return ('Conferido' if total == valor else 'Valor divergente'), total, candidatos


def cruzar_sistemas(relatorios, dcb=None, bolebarras=None, gai_lidas=None):
    indices = {
        'dcb': _indice(dcb['liquidados'], 'data_ocorrencia') if dcb else None,
        'bolebarra': _indice(bolebarras['registros'], 'data') if bolebarras else None,
    }
    for relatorio in relatorios:
        grupos = defaultdict(list)
        for r in relatorio['registros']:
            grupos[(r['boleto'], r['data_pagamento'])].append(r)
        linhas = []
        for chave, parcelas in grupos.items():
            valor = sum((r['valor'] for r in parcelas), ZERO)
            linha = dict(parcelas[0], valor=valor,
                         documento='; '.join(dict.fromkeys(r['documento'] for r in parcelas)))
            for origem, indice in indices.items():
                status, total, candidatos = _conferir(chave, indice, valor)
                linha['situacao_' + origem] = status
                linha['valor_' + origem] = total
                if origem == 'bolebarra' and not linha['nome'] and len(candidatos) == 1:
                    linha['nome'] = candidatos[0].get('nome', '')
            linhas.append(linha)
        relatorio['linhas'] = linhas

    if not bolebarras:
        return None
    baixas = defaultdict(lambda: defaultdict(list))
    for relatorio in relatorios:
        for r in relatorio['registros']:
            if r['boleto']:
                baixas[(r['boleto'], r['data_pagamento'])][relatorio['sistema']].append(r)
    if gai_lidas:
        for r in gai_lidas['parcelas']:
            if r['valor'] > ZERO:
                baixas[(r['boleto'].zfill(6), r['data_pagamento'])]['GAI'].append(r)
    linhas = []
    for chave, titulos in _indice(bolebarras['registros'], 'data').items():
        sistemas = baixas.get(chave, {})
        valor = sum((r['valor'] for r in titulos), ZERO)
        valor_sistemas = sum((r['valor'] for itens in sistemas.values() for r in itens), ZERO)
        if len(sistemas) > 1 or len({r['nosso_numero'] for r in titulos}) > 1:
            situacao = 'Identificação ambígua'
        elif not sistemas:
            situacao = 'Não encontrado nos sistemas enviados'
        else:
            situacao = 'Conferido' if valor == valor_sistemas else 'Valor divergente'
        linhas.append(dict(boleto=chave[0], data_pagamento=chave[1],
            nome=titulos[0].get('nome', ''), sistema=', '.join(sistemas),
            valores_por_sistema={s: sum((r['valor'] for r in itens), ZERO)
                                 for s, itens in sistemas.items()},
            valor=valor, valor_sistemas=valor_sistemas, situacao=situacao))
    # Títulos sem chave também precisam continuar visíveis como pendências.
    for r in bolebarras['registros']:
        if not _boleto(r.get('nosso_numero')):
            linhas.append(dict(boleto='', data_pagamento=r.get('data', ''),
                nome=r.get('nome', ''), sistema='', valor=r['valor'],
                valor_sistemas=ZERO, situacao='Sem nosso número válido'))
    ausentes = [r for r in linhas if not r['sistema']]
    divergentes = [r for r in linhas if r['situacao'] == 'Valor divergente']
    ambiguos = [r for r in linhas if r['situacao'] == 'Identificação ambígua']
    return dict(linhas=linhas, ausentes=ausentes, quantidade_ausentes=len(ausentes),
                sistemas_enviados=(["GAI"] if gai_lidas is not None else []) +
                                   [r['sistema'] for r in relatorios],
                total_francesinha=sum((r['valor'] for r in linhas), ZERO),
                quantidade_localizados=sum(bool(r['sistema']) for r in linhas),
                total_localizado=sum((r['valor'] for r in linhas if r['sistema']), ZERO),
                divergentes=divergentes, ambiguos=ambiguos,
                total_ausente=sum((r['valor'] for r in ausentes), ZERO),
                quantidade_pendentes=sum(r['situacao'] != 'Conferido' for r in linhas),
                total_pendente=sum((r['valor'] for r in linhas if r['situacao'] != 'Conferido'), ZERO))
