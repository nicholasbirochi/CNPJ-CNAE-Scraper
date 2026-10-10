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


def _aguardar_comprovante(contexto, pagina, cnpj: str, espera: int) -> dict:
    from playwright.sync_api import Error as ErroNavegador

    limite = time.monotonic() + espera
    while time.monotonic() < limite:
        paginas = [p for p in contexto.pages if not p.is_closed()]
        if not paginas:
            raise ErroConsulta("O navegador foi fechado antes de concluir a consulta.")
        for atual in paginas:
            for frame in atual.frames:
                origem = urlparse(frame.url)
                if (origem.hostname != "solucoes.receita.fazenda.gov.br"
                        or not origem.path.lower().startswith("/servicos/cnpjreva")):
                    continue
                try:
                    html = frame.content()
                except ErroNavegador:
                    continue  # A navegação pode trocar o documento durante a leitura.
                if "ERRO AO VALIDAR CAPTCHA" in _comparavel(html):
                    raise ErroConsulta(
                        "A Receita recusou a validação do CAPTCHA. "
                        "Para consultar a lista sem CAPTCHA, use: python lote.py"
                    )
                if _comparavel(SECUNDARIOS) in _comparavel(html):
                    return extrair_cnaes(html, cnpj)
        paginas[0].wait_for_timeout(500)
    raise ErroConsulta(
        "Tempo esgotado. Conclua o CAPTCHA e clique em CONSULTAR no navegador. "
        "Se a Receita mostrar um erro, corrija-o antes de repetir a consulta."
    )


def consultar_receita(cnpj: str, espera: int = 180) -> dict:
    from playwright.sync_api import Error as ErroNavegador, sync_playwright

    cnpj = normalizar_cnpj(cnpj)
    try:
        with sync_playwright() as playwright:
            with playwright.chromium.launch(headless=False) as navegador:
                contexto = navegador.new_context(locale="pt-BR")
                pagina = contexto.new_page()
                pagina.goto(URL_RECEITA, wait_until="domcontentloaded", timeout=30000)
                pagina.locator('input[type="text"]:visible').fill(cnpj, timeout=15000)
                print(
                    "Na janela da Receita, conclua o CAPTCHA e clique em CONSULTAR. "
                    f"Aguardando por até {espera}s…", file=sys.stderr, flush=True
                )
                return _aguardar_comprovante(contexto, pagina, cnpj, espera)
    except ErroNavegador as erro:
        if "Executable doesn't exist" in str(erro):
            raise ErroConsulta(
                "Chromium não instalado. Execute: python -m playwright install chromium"
            ) from erro
        raise ErroConsulta(
            "Não foi possível concluir a consulta no navegador. "
            "Verifique a conexão, a disponibilidade da Receita e tente novamente."
        ) from erro


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("cnpj", nargs="?", help="CNPJ numérico ou alfanumérico")
    parser.add_argument("--html", type=Path, help="Extrair de comprovante HTML já salvo")
    parser.add_argument("--saida", type=Path, help="Criar arquivo JSON sem sobrescrever")
    parser.add_argument("--espera", type=int, default=180, help="Espera do CAPTCHA em segundos")
    args = parser.parse_args(argv)
    if not args.cnpj and not args.html:
        parser.error("informe o CNPJ ou use --html com um comprovante salvo")
    if args.espera <= 0:
        parser.error("--espera deve ser maior que zero")
    try:
        cnpj = normalizar_cnpj(args.cnpj) if args.cnpj else None
        if args.html:
            bruto = args.html.read_bytes()
            try:
                html = bruto.decode("utf-8-sig")
            except UnicodeDecodeError:
                html = bruto.decode("cp1252")
            resultado = extrair_cnaes(html, cnpj)
        else:
            resultado = consultar_receita(cnpj, args.espera)
        saida = json.dumps(resultado, ensure_ascii=False, indent=2) + "\n"
        if args.saida:
            with args.saida.open("x", encoding="utf-8") as arquivo:
                arquivo.write(saida)
            print(f"JSON salvo em {args.saida}", file=sys.stderr)
        else:
            print(saida, end="")
        return 0
    except (ErroConsulta, OSError, UnicodeError, ImportError) as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Consulta cancelada.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
