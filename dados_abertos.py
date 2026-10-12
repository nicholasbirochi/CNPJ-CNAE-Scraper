"""Preenche o output.xlsx com os Dados Abertos do CNPJ publicados pela Receita Federal."""

import argparse
import base64
import csv
from http.client import HTTPException
import io
import json
from pathlib import Path
import re
import sqlite3
import ssl
import sys
import tempfile
import time
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
import zipfile

import certifi

from consultar import ErroConsulta, normalizar_cnpj
from lote import RAIZ, _atividade, ler_cnpjs, processar

# Link público: https://arquivos.receitafederal.gov.br/index.php/s/YggdBLfdninEJX9
WEBDAV = "https://arquivos.receitafederal.gov.br/public.php/webdav/"
COMPARTILHAMENTO = "YggdBLfdninEJX9"
PASTA_PADRAO = Path(tempfile.gettempdir()) / "cnpj_dados_abertos"
ARQUIVOS = ["Cnaes.zip"] + [f"Estabelecimentos{i}.zip" for i in range(10)]
LINHA_ESTABELECIMENTO = re.compile(rb'^"[0-9A-Z]{8}";"[0-9A-Z]{4}";"[0-9]{2}";')


def _abrir(caminho: str, metodo: str = "GET", **cabecalhos):
    credencial = base64.b64encode(f"{COMPARTILHAMENTO}:".encode()).decode()
    pedido = Request(WEBDAV + caminho, method=metodo, headers={
        "Authorization": f"Basic {credencial}", "User-Agent": "CNPJ-CNAE-Scraper/1.0", **cabecalhos
    })
    return urlopen(pedido, timeout=60, context=ssl.create_default_context(cafile=certifi.where()))


def _listar(caminho: str) -> dict[str, int | None]:
    """Nomes e tamanhos de uma pasta do compartilhamento (WebDAV PROPFIND)."""
    with _abrir(caminho, "PROPFIND", Depth="1") as resposta:
        xml = resposta.read().decode("utf-8")
    itens = {}
    for bloco in re.findall(r"<d:response>(.*?)</d:response>", xml, re.S):
        nome = re.search(r"<d:href>[^<]*?([^/<]+)/?</d:href>", bloco).group(1)
        tamanho = re.search(r"<d:getcontentlength>(\d+)</d:getcontentlength>", bloco)
        itens[nome] = int(tamanho.group(1)) if tamanho else None
    return itens


def mes_mais_recente() -> str:
    meses = [nome for nome in _listar("") if re.fullmatch(r"\d{4}-\d{2}", nome)]
    if not meses:
        raise ErroConsulta("Nenhuma pasta mensal encontrada nos Dados Abertos do CNPJ.")
    return max(meses)


def baixar(mes: str, nome: str, tamanho: int, destino: Path) -> None:
    """Uma conexão por vez, retomando após quedas e abortando transferências travadas."""
    parcial = destino.with_suffix(".parcial")
    espera = 30
    while True:
        inicio = parcial.stat().st_size if parcial.exists() else 0
        if inicio == tamanho:
            parcial.replace(destino)
            return
        if inicio > tamanho:
            parcial.unlink()
            continue
        try:
            with _abrir(f"{mes}/{nome}", Range=f"bytes={inicio}-") as resposta:
                with parcial.open("ab" if resposta.status == 206 else "wb") as arquivo:
                    janela, recebidos = time.monotonic(), 0
                    while bloco := resposta.read(1 << 20):
                        arquivo.write(bloco)
                        recebidos += len(bloco)
                        espera = 30
                        if time.monotonic() - janela > 30:
                            if recebidos < 1 << 20:
                                raise TimeoutError("transferência travada")
                            janela, recebidos = time.monotonic(), 0
        except HTTPError as erro:
            if erro.code not in (408, 425, 429, 500, 502, 503, 504):
                raise ErroConsulta(f"Receita recusou {nome} (HTTP {erro.code}).") from erro
            motivo = f"HTTP {erro.code}"
        except (URLError, OSError, HTTPException) as erro:
            motivo = str(getattr(erro, "reason", erro)) or type(erro).__name__
        else:
            if (parcial.stat().st_size if parcial.exists() else 0) > inicio:
                continue
            motivo = "resposta vazia"
        atual = parcial.stat().st_size if parcial.exists() else 0
        print(f"  {nome}: {atual * 100 // tamanho}%, conexão interrompida ({motivo}); "
              f"nova tentativa em {espera}s.", file=sys.stderr, flush=True)
        time.sleep(espera)
        espera = min(espera * 2, 600)


def baixar_mes(pasta: Path, mes: str | None = None) -> Path:
    mes = mes or mes_mais_recente()
    destino = pasta / mes
    destino.mkdir(parents=True, exist_ok=True)
    tamanhos = _listar(f"{mes}/")
    for nome in ARQUIVOS:
        if nome not in tamanhos or not tamanhos[nome]:
            raise ErroConsulta(f"{nome} não encontrado em {mes} nos Dados Abertos do CNPJ.")
        if not (destino / nome).exists():
            print(f"Baixando {nome} de {mes} ({tamanhos[nome] / 2**20:.0f} MB)…", flush=True)
            baixar(mes, nome, tamanhos[nome], destino / nome)
    return destino


def ler_cnaes(arquivo: Path) -> dict[str, str]:
    with zipfile.ZipFile(arquivo) as pacote, pacote.open(pacote.namelist()[0]) as bruto:
        # Espaços repetidos na tabela oficial são normalizados como na BrasilAPI.
        return {codigo.zfill(7): " ".join(descricao.split()) for codigo, descricao
                in csv.reader(io.TextIOWrapper(bruto, encoding="latin-1"), delimiter=";")}


def buscar(arquivo: Path, alvos: set[bytes]) -> dict[str, list[str]]:
    """Devolve os campos dos estabelecimentos procurados em um Estabelecimentos*.zip."""
    encontrados = {}
    with zipfile.ZipFile(arquivo) as pacote:
        for nome in pacote.namelist():
            with pacote.open(nome) as bruto:
                resto, primeiro = b"", True
                while True:
                    bloco = bruto.read(1 << 24)
                    linhas = (resto + bloco).split(b"\n")
                    resto = linhas.pop() if bloco else b""
                    if primeiro and linhas and linhas[0]:
                        if not LINHA_ESTABELECIMENTO.match(linhas[0]):
                            raise ErroConsulta(f"Layout não reconhecido em {arquivo.name}.")
                        primeiro = False
                    for linha in linhas:
                        # Campos entre aspas: "BASE";"ORDEM";"DV";…
                        if linha[1:9] + linha[12:16] + linha[19:21] in alvos:
                            campos = next(csv.reader([linha.decode("latin-1").rstrip("\r")], delimiter=";"))
                            encontrados["".join(campos[:3])] = campos
                    if not bloco:
                        break
    return encontrados


def montar_resultado(campos: list[str], tabela: dict[str, str]) -> dict:
    def atividade(codigo: str) -> dict:
        codigo = codigo.strip().zfill(7)
        if codigo not in tabela:
            raise ErroConsulta(f"CNAE {codigo} ausente da tabela de CNAEs da Receita.")
        return _atividade(codigo, tabela[codigo])

    secundarios = []
    for codigo in campos[12].split(","):
        if codigo.strip().strip("0"):
            item = atividade(codigo)
            if item not in secundarios:
                secundarios.append(item)
    return {"cnae_principal": atividade(campos[11]), "cnaes_secundarios": secundarios}


def importar(pasta: Path, entrada: Path, saida: Path) -> int:
    validos = set()
    for original in ler_cnpjs(entrada, saida):
        try:
            validos.add(normalizar_cnpj(original))
        except ErroConsulta:
            pass
    arquivos = sorted(pasta.glob("Estabelecimentos*.zip"))
    if not (pasta / "Cnaes.zip").exists() or not arquivos:
        raise ErroConsulta(f"Cnaes.zip e Estabelecimentos*.zip não encontrados em {pasta}.")
    tabela = ler_cnaes(pasta / "Cnaes.zip")
    cache = sqlite3.connect(saida.with_suffix(".cache.sqlite3"))
    try:
        cache.execute("CREATE TABLE IF NOT EXISTS cnaes (cnpj TEXT PRIMARY KEY, consultado_em REAL, resultado TEXT)")
        recentes = {cnpj for (cnpj,) in cache.execute(
            "SELECT cnpj FROM cnaes WHERE consultado_em > ?", (time.time() - 30 * 86400,))}
        alvos = {cnpj.encode("ascii") for cnpj in validos - recentes}
        print(f"{len(validos)} CNPJs válidos, {len(alvos)} sem resultado no cache.", flush=True)
        preenchidos = 0
        for arquivo in arquivos:
            if not alvos:
                break
            inicio = time.monotonic()
            for cnpj, campos in buscar(arquivo, alvos).items():
                try:
                    resultado = montar_resultado(campos, tabela)
                except ErroConsulta as erro:
                    print(f"{cnpj}: {erro}", file=sys.stderr)
                    continue
                cache.execute("INSERT OR REPLACE INTO cnaes VALUES (?, ?, ?)",
                              (cnpj, time.time(), json.dumps(resultado, ensure_ascii=False)))
                alvos.discard(cnpj.encode("ascii"))
                preenchidos += 1
            cache.commit()
            print(f"{arquivo.name}: {preenchidos} preenchidos até agora "
                  f"({time.monotonic() - inicio:.0f}s).", flush=True)
    finally:
        cache.close()
    print(f"Dados Abertos: {preenchidos} CNPJs preenchidos; {len(alvos)} não encontrados.", flush=True)
    if alvos:
        print("Para consultar os restantes na BrasilAPI: python lote.py", flush=True)
    return processar(entrada, saida, somente_cache=True)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("entrada", nargs="?", type=Path, default=RAIZ / "data")
    parser.add_argument("--saida", type=Path, default=RAIZ / "output.xlsx")
    parser.add_argument("--pasta", type=Path, default=PASTA_PADRAO,
                        help="Onde guardar os arquivos da Receita (cerca de 5 GB por mês)")
    parser.add_argument("--mes", help="Mês publicado, como 2026-09 (padrão: o mais recente)")
    parser.add_argument("--zips", type=Path, help="Usar arquivos já baixados nesta pasta, sem acessar a internet")
    args = parser.parse_args(argv)
    try:
        pasta = args.zips or baixar_mes(args.pasta, args.mes)
        return importar(pasta, args.entrada, args.saida)
    except (ErroConsulta, OSError, ValueError, zipfile.BadZipFile) as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Interrompido. Downloads parciais são retomados na próxima execução.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
