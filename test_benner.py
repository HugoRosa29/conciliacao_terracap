from datetime import datetime
from decimal import Decimal as D
from pathlib import Path
import os
import unittest
from unittest.mock import patch

from benner import ler_benner, conferir_benner


def parcela(valor="10", numero="1", data="28/09/2026"):
    return dict(alienacao="123", imovel="456", parcela=numero,
                data_pagamento=data, valor=D(valor), boleto="987654")


def integracao(*registros):
    return dict(registros=[dict(r, aba="Dados", linha=i+2) for i, r in enumerate(registros)],
                datas=["28/09/2026"])


class BennerTest(unittest.TestCase):
    def test_ausentes_valores_e_repeticoes(self):
        gai = {"parcelas": [parcela("11"), parcela("10"), parcela("10"), parcela(numero="2")]}
        benner = integracao(parcela("10"), parcela("12"))
        r = conferir_benner(gai, benner)
        self.assertEqual(r["quantidade_encontrados"], 1)
        self.assertEqual(r["quantidade_divergentes"], 1)
        self.assertEqual(r["divergentes"][0]["valor"], D("11"))
        self.assertEqual(r["quantidade_ausentes"], 2)
        self.assertEqual(r["total_ausente"], D("20"))

    def test_datas_diferentes_nao_viram_ausencias(self):
        r = conferir_benner({"parcelas": [parcela(data="30/09/2026")]}, integracao(parcela()))
        self.assertEqual(r["status"], "datas_incompativeis")
        self.assertEqual(r["ausentes"], [])
        r = conferir_benner({"parcelas": [parcela(), parcela(data="30/09/2026")]}, integracao(parcela()))
        self.assertEqual(len(r["nao_comparados"]), 1)
        self.assertEqual(r["quantidade_encontrados"], 1)

    def test_opcional_e_zeros(self):
        self.assertEqual(conferir_benner(None, integracao())["status"], "aguardando_gai")
        self.assertEqual(conferir_benner({"parcelas": [parcela("0")]}, integracao())["ausentes"], [])

    def test_leitura_historicos_e_creditos(self):
        linhas = [["Data", "Documento", "Valor", "D/C", "Histórico"],
            [datetime(2026, 9, 29), None, 10, "Crédito", "Alienação 123 Imv 456 Pag parc 01/2 Dt Pag 28/09/2026"],
            [datetime(2026, 9, 28), None, 20, "Crédito", "ALIENAÇÃO 123 IMÓVEL 456 (PARCELA 2)"],
            [datetime(2026, 9, 28), None, 30, "Crédito", "Alienação ?????? PIX"],
            [datetime(2026, 9, 28), None, 10, "Débito", "Alienação 123 Imv 456 Pag parc 1/2"],
            [datetime(2026, 9, 28), None, 0, "Crédito", "Alienação 123 Imv 456 Pag parc 3/2"]]
        with patch("benner.ler_planilha", return_value={"Dados": linhas}):
            r = ler_benner(b"teste")
        self.assertEqual(r["quantidade"], 2)
        self.assertEqual(r["total"], D("30"))
        self.assertEqual(r["registros"][0]["parcela"], "1")
        self.assertEqual(r["registros"][0]["data_pagamento"], "28/09/2026")
        self.assertEqual(r["quantidade_nao_identificados"], 1)
        with patch("benner.ler_planilha", return_value={"Errada": [["Outro formato"]]}):
            with self.assertRaises(ValueError):
                ler_benner(b"teste")

    @unittest.skipUnless(os.environ.get("CONCILIACAO_TESTE3"), "Defina CONCILIACAO_TESTE3")
    def test_arquivo_real(self):
        arquivo = next(Path(os.environ["CONCILIACAO_TESTE3"]).glob("*BENNER*"))
        r = ler_benner(arquivo)
        self.assertEqual(r["quantidade"], 499)
        self.assertEqual(r["quantidade_nao_identificados"], 7)
        self.assertEqual(r["datas"], ["28/09/2026"])
        self.assertEqual(r["registros"][0]["alienacao"], "112320")


if __name__ == "__main__":
    unittest.main()
