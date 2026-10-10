from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from consultar import ErroConsulta, PRINCIPAL, SECUNDARIOS, extrair_cnaes, main, normalizar_cnpj

FIXTURES = Path(__file__).parent / "fixtures"


class TestCnpj(unittest.TestCase):
    def test_numerico_com_mascara_e_zero_inicial(self):
        self.assertEqual(normalizar_cnpj("00.000.000/0001-91"), "00000000000191")

    def test_alfanumerico_exemplo_da_receita(self):
        self.assertEqual(normalizar_cnpj("12.abc.345/01de-35"), "12ABC34501DE35")

    def test_mascara_sem_pontos_do_arquivo_de_entrada(self):
        self.assertEqual(normalizar_cnpj("00000000/0001-91"), "00000000000191")

    def test_primeiro_alfanumerico_publicado_pela_receita(self):
        self.assertEqual(normalizar_cnpj("00.000.000/E08G-12"), "00000000E08G12")

    def test_rejeita_entrada_invalida_sem_apagar_caracteres(self):
        for valor in ("00000000000000", "11111111111111", "00000000000190",
                      "texto 00.000.000/0001-91", "00000000000191!", "",
                      "12.ABC.345/01DE-AA", "00/000/000/0001/91", "０0000000000191"):
            with self.subTest(valor=valor), self.assertRaises(ErroConsulta):
                normalizar_cnpj(valor)


class TestExtracao(unittest.TestCase):
    def setUp(self):
        self.html = (FIXTURES / "comprovante.html").read_text(encoding="utf-8")

    def test_extrai_somente_dois_campos_e_todos_secundarios(self):
        resultado = extrair_cnaes(self.html, "00000000000191")
        self.assertEqual(set(resultado), {"cnae_principal", "cnaes_secundarios"})
        self.assertEqual(resultado["cnae_principal"], {
            "codigo": "62.01-5-01",
            "descricao": "Desenvolvimento de programas de computador sob encomenda",
        })
        self.assertEqual([item["codigo"] for item in resultado["cnaes_secundarios"]],
                         ["62.02-3-00", "62.04-0-00"])
        self.assertNotIn("Dispensada", str(resultado))
        self.assertNotIn("RUA FICTÍCIA", str(resultado))
        self.assertNotIn("0001-91", str(resultado))

    def test_legado_sem_secundarios(self):
        html = (FIXTURES / "legado.html").read_text(encoding="utf-8")
        self.assertEqual(extrair_cnaes(html)["cnaes_secundarios"], [])

    def test_espacos_entidades_acentos_e_descricao_quebrada(self):
        html = self.html.replace("CÓDIGO", "C&Oacute;DIGO").replace("DESCRIÇÃO", "DESCRICAO")
        html = html.replace("62.04-0-00 - Consultoria", "62.04-0-00 &ndash; Consultoria&nbsp;")
        html = html.replace("tecnologia da informação", "tecnologia<br>da informação")
        self.assertEqual(extrair_cnaes(html)["cnaes_secundarios"][-1]["descricao"],
                         "Consultoria em tecnologia da informação")

    def test_paginas_repetidas_nao_duplicam_cnaes(self):
        self.assertEqual(extrair_cnaes(self.html + self.html), extrair_cnaes(self.html))

    def test_secundarios_de_todas_paginas_sao_preservados(self):
        pagina2 = self.html.replace("62.04-0-00 - Consultoria em tecnologia da informação",
                                    "63.11-9-00 - Tratamento de dados")
        codigos = [item["codigo"] for item in extrair_cnaes(self.html + pagina2)["cnaes_secundarios"]]
        self.assertEqual(codigos, ["62.02-3-00", "62.04-0-00", "63.11-9-00"])

    def test_rejeita_outro_cnpj(self):
        with self.assertRaisesRegex(ErroConsulta, "outro CNPJ"):
            extrair_cnaes(self.html, "12ABC34501DE35")

    def test_campo_ausente_nao_captura_vizinhos(self):
        for campo in (PRINCIPAL, SECUNDARIOS):
            with self.subTest(campo=campo), self.assertRaisesRegex(ErroConsulta, "Campo não encontrado"):
                extrair_cnaes(self.html.replace(campo, "CAMPO REMOVIDO"))

    def test_rejeita_sem_descricao(self):
        html = self.html.replace("Desenvolvimento de programas de computador sob encomenda", "")
        with self.assertRaisesRegex(ErroConsulta, "sem descrição"):
            extrair_cnaes(html)

    def test_rejeita_mascarado_em_vez_de_inventar_vazio(self):
        html = (FIXTURES / "legado.html").read_text(encoding="utf-8")
        with self.assertRaisesRegex(ErroConsulta, "mascarado"):
            extrair_cnaes(html.replace("Não informada", "********"))

    def test_rejeita_principais_divergentes(self):
        pagina2 = self.html.replace("62.01-5-01", "62.03-1-00")
        with self.assertRaisesRegex(ErroConsulta, "exatamente um"):
            extrair_cnaes(self.html + pagina2)


class TestCli(unittest.TestCase):
    def test_json_em_stdout_sem_mensagens_misturadas(self):
        import json
        saida = StringIO()
        with redirect_stdout(saida):
            codigo = main(["--html", str(FIXTURES / "comprovante.html")])
        self.assertEqual(codigo, 0)
        self.assertEqual(len(json.loads(saida.getvalue())), 2)

    def test_cria_json_e_recusa_sobrescrita(self):
        with tempfile.TemporaryDirectory() as pasta, redirect_stderr(StringIO()):
            destino = Path(pasta) / "cnaes.json"
            args = ["--html", str(FIXTURES / "comprovante.html"), "--saida", str(destino)]
            self.assertEqual(main(args), 0)
            original = destino.read_bytes()
            self.assertEqual(main(args), 1)
            self.assertEqual(destino.read_bytes(), original)

    def test_html_cp1252(self):
        with tempfile.TemporaryDirectory() as pasta, redirect_stdout(StringIO()):
            caminho = Path(pasta) / "comprovante.html"
            caminho.write_bytes((FIXTURES / "legado.html").read_text(encoding="utf-8").encode("cp1252"))
            self.assertEqual(main(["--html", str(caminho)]), 0)

    def test_cnpj_invalido_nao_abre_navegador(self):
        with patch("consultar.consultar_receita") as navegador, redirect_stderr(StringIO()):
            self.assertEqual(main(["00000000000000"]), 1)
            navegador.assert_not_called()

    def test_arquivo_ausente_falha_sem_traceback(self):
        with redirect_stderr(StringIO()) as erros:
            self.assertEqual(main(["--html", "/arquivo-que-nao-existe.html"]), 1)
        self.assertIn("Erro:", erros.getvalue())


if __name__ == "__main__":
    unittest.main()
