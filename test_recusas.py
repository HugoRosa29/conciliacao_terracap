"""Regressões da identificação de recusas e do contrato com a tela."""

from decimal import Decimal
import unittest

from conciliacao import conciliar, identificar_recusados
from ponte import para_tela


def recusa(boleto, alienacao, valor="0.00"):
    return dict(boleto=boleto, alienacao=alienacao, total_pago=Decimal(valor),
                parcela="1", data_pagamento="30/09/2026", motivo="Recusado")


def titulo(boleto, alienacao, nome):
    return dict(boleto=boleto, documento=alienacao, nome=nome,
                nosso_numero="6" + boleto + "07000")


class RecusasTest(unittest.TestCase):
    def test_identifica_boleto_e_segunda_via_sem_atribuir_nome_ambiguo(self):
        registros = [recusa("123456", "10"), recusa("999999", "10"),
                     recusa("999998", "20"), recusa("999997", "30")]
        relatorio = {"registros": registros}
        francesinha = {"registros": [
            titulo("123456", "10", "ANA SILVA"),
            titulo("123457", "10", "ANA SIL"),
            titulo("223456", "20", "MARIA COSTA"),
            titulo("223457", "20", "JOANA COSTA"),
        ]}
        self.assertEqual(identificar_recusados(relatorio, francesinha), 2)
        self.assertEqual(registros[0]["origem_do_nome"], "boleto")
        self.assertEqual(registros[1]["nome"], "ANA SILVA")
        self.assertEqual(registros[1]["origem_do_nome"], "alienacao")
        self.assertEqual(registros[1]["nosso_numero"], "")
        self.assertEqual(registros[2]["nome"], "")
        self.assertEqual(registros[3]["nome"], "")

    def test_tela_preserva_resumo_e_detalha_so_recusas_com_valor(self):
        registros = [recusa("123456", "10", "1921.07"),
                     recusa("123457", "20")]
        motivos = [{"motivo": "Recusado", "quantidade": 2,
                    "total": Decimal("1921.07")}]
        relatorio = dict(registros=registros, com_valor=registros[:1],
                         quantidade=2, quantidade_com_valor=1,
                         total=Decimal("1921.07"), motivos=motivos)
        identificar_recusados(relatorio, {"registros": []})
        resultado = conciliar()
        resultado["gai_nao_baixadas"] = relatorio
        tela = para_tela(resultado)["gai_nao_baixadas"]
        self.assertEqual(tela["motivos"], motivos)
        self.assertEqual(len(tela["registros"]), 1)
        self.assertEqual(tela["registros"][0]["total_pago"], Decimal("1921.07"))
        self.assertEqual(tela["registros"][0]["nome"], "")
        self.assertEqual(len(relatorio["registros"]), 2)
        self.assertEqual(tela["quantidade_identificados"], 0)


if __name__ == "__main__":
    unittest.main()
