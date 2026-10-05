"""Conferência das parcelas pagas do GAI com a integração contábil BENNER."""
from collections import defaultdict
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import re

from gai import _sem_acento
from planilhas import ler_planilha

ZERO = Decimal('0.00')


def _data(valor):
    if isinstance(valor, (datetime, date)):
        return valor.strftime('%d/%m/%Y')
    texto = str(valor or '').strip().split(' ')[0]
    try:
        return datetime.strptime(texto, '%d/%m/%Y').strftime('%d/%m/%Y')
    except ValueError:
        return ''


def _numero(valor):
    texto = str(valor or '').strip()
    return str(int(texto)) if texto.isdigit() else texto


def _valor(valor):
    if isinstance(valor, str):
        valor = valor.replace('R$', '').strip()
        if ',' in valor:
            valor = valor.replace('.', '').replace(',', '.')
    try:
        numero = Decimal(str(valor)).quantize(Decimal('0.01'))
        if numero.is_finite():
            return numero
    except InvalidOperation:
        pass
    raise ValueError('Valor inválido na integração BENNER.')


def ler_benner(origem):
    registros, nao_identificados, datas = [], [], set()
    encontrou = False
    for aba, linhas in ler_planilha(origem).items():
        mapa = None
        for numero_linha, linha in enumerate(linhas, 1):
            cabecalho = {_sem_acento(v): i for i, v in enumerate(linha) if v is not None}
            if all(c in cabecalho for c in ('data', 'valor', 'd/c', 'historico')):
                mapa = cabecalho
                encontrou = True
                continue
            if mapa is None or not any(v is not None for v in linha):
                continue
            def celula(c):
                i = mapa[c]
                return linha[i] if i < len(linha) else None
            # Débitos e estornos não comprovam a integração de um recebimento.
            if _sem_acento(celula('d/c')) not in ('credito', 'c'):
                continue
            valor = _valor(celula('valor'))
            if valor <= ZERO:
                continue
            historico = str(celula('historico') or '')
            texto = _sem_acento(historico)
            alienacao = re.search(r'alienacao\s+(\d+)', texto)
            imovel = re.search(r'(?:imovel|imv)\s+(\d+)', texto)
            parcela = re.search(r'(?:pag\s+parc|parcela)\s+(\d+)(?:\s*/\s*\d+)?', texto)
            pagamento = re.search(r'dt\s+pag(?:amento)?\s+(\d{2}/\d{2}/\d{4})', texto)
            data = _data(pagamento[1] if pagamento else celula('data'))
            if data:
                datas.add(data)
            registro = dict(aba=aba, linha=numero_linha, historico=historico,
                alienacao=_numero(alienacao[1]) if alienacao else '',
                imovel=_numero(imovel[1]) if imovel else '',
                parcela=_numero(parcela[1]) if parcela else '',
                data_pagamento=data, valor=valor)
            if all(registro[c] for c in ('alienacao', 'imovel', 'parcela', 'data_pagamento')):
                registros.append(registro)
            else:
                nao_identificados.append(registro)
    if not encontrou:
        raise ValueError('BENNER: cabeçalhos Data, Valor, D/C e Histórico não encontrados.')
    avisos = []
    if nao_identificados:
        avisos.append(f'BENNER: {len(nao_identificados)} crédito(s) sem identificação completa de parcela; consulte os lançamentos para revisão.')
    return dict(registros=registros, nao_identificados=nao_identificados,
        quantidade=len(registros), quantidade_nao_identificados=len(nao_identificados),
        datas=sorted(datas), total=sum((r['valor'] for r in registros), ZERO), avisos=avisos)


def _chave(r):
    return tuple(_numero(r.get(c)) for c in ('alienacao', 'imovel', 'parcela')) + (_data(r.get('data_pagamento')),)


def conferir_benner(gai, benner):
    resultado = dict(status='aguardando_gai', ausentes=[], divergentes=[], encontrados=[],
        nao_comparados=[], quantidade_ausentes=0, total_ausente=ZERO,
        quantidade_divergentes=0, quantidade_encontrados=0, avisos=[])
    if gai is None:
        return resultado
    parcelas = [r for r in gai['parcelas'] if r['valor'] > ZERO]
    datas_gai = {_data(r['data_pagamento']) for r in parcelas} - {''}
    datas_benner = set(benner['datas'])
    if datas_gai and not datas_gai.intersection(datas_benner):
        resultado['status'] = 'datas_incompativeis'
        resultado['avisos'].append('GAI e BENNER não têm datas de pagamento em comum. Envie arquivos do mesmo período.')
        resultado['nao_comparados'] = parcelas
        return resultado
    resultado['status'] = 'comparado'
    indice = defaultdict(list)
    for r in benner['registros']:
        indice[_chave(r)].append(r)
    pendentes = []
    # Consumir primeiro as correspondências exatas evita que uma duplicata
    # com valor diferente tome o lugar de outra que está corretamente integrada.
    for r in parcelas:
        chave = _chave(r)
        if not all(chave) or chave[-1] not in datas_benner:
            resultado['nao_comparados'].append(r)
            continue
        candidatos = indice[chave]
        exato = next((i for i, b in enumerate(candidatos) if b['valor'] == r['valor']), None)
        if exato is not None:
            b = candidatos.pop(exato)
            resultado['encontrados'].append(dict(r, linha_benner=b['linha'], aba_benner=b['aba']))
        else:
            pendentes.append(r)
    for r in pendentes:
        candidatos = indice[_chave(r)]
        if candidatos:
            b = candidatos.pop(0)
            resultado['divergentes'].append(dict(r, valor_benner=b['valor'],
                diferenca=r['valor']-b['valor'], linha_benner=b['linha'], aba_benner=b['aba']))
        else:
            resultado['ausentes'].append(r.copy())
    if resultado['nao_comparados']:
        resultado['avisos'].append('Algumas parcelas do GAI não puderam ser comparadas: identificação incompleta ou data fora do arquivo BENNER.')
    resultado.update(quantidade_ausentes=len(resultado['ausentes']),
        total_ausente=sum((r['valor'] for r in resultado['ausentes']), ZERO),
        quantidade_divergentes=len(resultado['divergentes']),
        quantidade_encontrados=len(resultado['encontrados']))
    return resultado
