"""Consulta assistida na Receita Federal e extração exclusiva dos CNAEs."""

import argparse
import json
from pathlib import Path
import re
import sys
import time
import unicodedata
from urllib.parse import urlparse

from bs4 import BeautifulSoup

URL_RECEITA = "https://solucoes.receita.fazenda.gov.br/Servicos/cnpjreva/"
PRINCIPAL = "CÓDIGO E DESCRIÇÃO DA ATIVIDADE ECONÔMICA PRINCIPAL"
SECUNDARIOS = "CÓDIGO E DESCRIÇÃO DAS ATIVIDADES ECONÔMICAS SECUNDÁRIAS"
INSCRICAO = "NÚMERO DE INSCRIÇÃO"
CODIGO_CNAE = re.compile(r"(?P<codigo>\d{2}\.\d{2}-\d-\d{2})\s*[-–—]\s*")


class ErroConsulta(ValueError):
    """Dados inválidos, indisponíveis ou layout não reconhecido."""


def normalizar_cnpj(valor: str) -> str:
    """Valida formato e DV numérico ou alfanumérico (ASCII menos 48)."""
    valor = valor.strip().upper()
    if not re.fullmatch(
        r"(?:[A-Z0-9]{12}[0-9]{2}|[A-Z0-9]{8}/[A-Z0-9]{4}-[0-9]{2}|[A-Z0-9]{2}\.[A-Z0-9]{3}\."
        r"[A-Z0-9]{3}/[A-Z0-9]{4}-[0-9]{2})", valor
    ):
        raise ErroConsulta("CNPJ deve ter 14 posições, com ou sem pontos, barra e hífen.")
    cnpj = re.sub(r"[./-]", "", valor)
    if len(set(cnpj)) == 1:
        raise ErroConsulta("CNPJ inválido: caracteres repetidos.")
    base = cnpj[:12]
    for pesos in ((5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2),
                  (6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)):
        resto = sum((ord(c) - 48) * p for c, p in zip(base, pesos)) % 11
        base += str(0 if resto < 2 else 11 - resto)
    if base != cnpj:
        raise ErroConsulta("CNPJ inválido: os dígitos verificadores não conferem.")
    return cnpj


def _texto(valor: str) -> str:
    return " ".join(unicodedata.normalize("NFC", valor).split())


def _comparavel(valor: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", _texto(valor))
                   if not unicodedata.combining(c)).upper()


def _valores(soup: BeautifulSoup, rotulo: str) -> list[str]:
    valores = []
    for celula in soup.find_all("td"):
        # As tabelas externas também contêm os rótulos; só a célula final é o campo.
        if celula.find("td"):
            continue
        texto = _texto(celula.get_text(" ", strip=True))
        if _comparavel(texto).startswith(_comparavel(rotulo)):
            valor = texto[len(_texto(rotulo)):].strip()
            valor = re.sub(r"^\s*\(\*\)\s*", "", valor)
            valores.append(valor)
    if not valores:
        raise ErroConsulta(f"Campo não encontrado: {rotulo}. O layout pode ter mudado.")
    return valores


def _atividades(valores: list[str], permitir_vazio: bool = False) -> list[dict]:
    atividades = []
    for valor in valores:
        valor = _texto(re.sub(r"\((?:Não )?[Dd]ispensada\s*\*\)", "", valor))
        if permitir_vazio and _comparavel(valor) in {"NAO INFORMADA", "NAO INFORMADO"}:
            continue
        codigos = list(CODIGO_CNAE.finditer(valor))
        if not codigos or codigos[0].start() != 0:
            raise ErroConsulta("CNAE vazio, mascarado ou em formato não reconhecido.")
        for i, codigo in enumerate(codigos):
            fim = codigos[i + 1].start() if i + 1 < len(codigos) else len(valor)
            descricao = valor[codigo.end():fim].strip()
            if not descricao:
                raise ErroConsulta("CNAE sem descrição no comprovante.")
            atividade = {"codigo": codigo.group("codigo"), "descricao": descricao}
            if atividade not in atividades:
                atividades.append(atividade)
    return atividades


def extrair_cnaes(html: str, cnpj_esperado: str | None = None) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    for elemento in soup(["script", "style"]):
        elemento.decompose()
    if cnpj_esperado:
        esperado = normalizar_cnpj(cnpj_esperado)
        for inscricao in _valores(soup, INSCRICAO):
            encontrado = inscricao.split()[0] if inscricao else ""
            if normalizar_cnpj(encontrado) != esperado:
                raise ErroConsulta("O comprovante pertence a outro CNPJ.")
    principal = _atividades(_valores(soup, PRINCIPAL))
    if len(principal) != 1:
        raise ErroConsulta("Esperado exatamente um CNAE principal.")
    secundarios = _atividades(_valores(soup, SECUNDARIOS), permitir_vazio=True)
    return {"cnae_principal": principal[0], "cnaes_secundarios": secundarios}
