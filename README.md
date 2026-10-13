# CNPJ CNAE Scraper

## Lote de CNPJs e output.xlsx

Coloque os arquivos **XLSX, CSV, TSV ou TXT** em `data/` e execute, após instalar as dependências:

```bash
python lote.py
```

O programa gera **`output.xlsx`** com três colunas: **CNPJ**, **CNAE principal** e **CNAEs secundários**. Código e descrição ficam juntos; todos os secundários aparecem na mesma célula, separados por quebras de linha. A fonte do lote é a [BrasilAPI gratuita](https://brasilapi.com.br/docs#tag/CNPJ), com dados oriundos da Receita Federal. Não exige CAPTCHA ou chave de API.

```bash
python lote.py data/entrada.xlsx --saida output.xlsx
python lote.py --exportar   # Atualizar o Excel a partir do cache, sem novas consultas
python lote.py --atualizar  # Ignorar o cache e refazer consultas
```

O leitor procura uma coluna cujo cabeçalho contenha `CNPJ`, como `CPF_CNPJ1`. Sem cabeçalho, procura identificadores em formato reconhecido. Lê todas as abas e arquivos, preservando a ordem e as linhas repetidas. Cada CNPJ é consultado uma única vez. Números armazenados pelo Excel sem zeros iniciais são convertidos para texto com 14 posições. Aceita a máscara sem pontos (`12345678/0001-90`) e sinaliza CNPJs com dígitos verificadores inválidos. Converta `.xls` antigo para `.xlsx`.

**Lotes grandes levam tempo.** São feitas consultas sequenciais com intervalo padrão de 2 segundos, além da resposta da API. Cada sucesso é salvo em `output.cache.sqlite3`. O Excel é atualizado a cada 300 consultas e ao encerrar. Ao repetir o comando, o programa reutiliza sucessos dos últimos 30 dias. Mantenha o computador ligado e conectado; execute somente um lote por arquivo de saída e feche o Excel antes de atualizar.

- `PENDENTE`: ainda não consultado.
- `ERRO: ...`: CNPJ inválido, indisponibilidade ou empresa ausente da base.
- `Não informada`: a fonte informou que não há CNAEs secundários.
- `Não consultado`: nenhum resultado válido foi obtido para a linha.

Ctrl+C salva o Excel antes de sair. Encerramentos abruptos preservam sucessos no cache. HTTP 403/429 encerra o lote sem contornar o bloqueio; tente mais tarde. A saída é substituída de forma atômica, e um arquivo existente com outro cabeçalho é preservado. Dados externos são gravados como texto, evitando fórmulas injetadas. HTTPS mantém validação de certificado via certifi.

`data/`, planilhas, cache e logs são ignorados no Git. A BrasilAPI pode ter defasagem e indisponibilidade; seus resultados não substituem um comprovante oficial atualizado. [Orientação dos mantenedores sobre lotes de CNAE](https://github.com/BrasilAPI/BrasilAPI/issues/402).

### Lotes grandes: Dados Abertos do CNPJ

Para dezenas de milhares de CNPJs, a BrasilAPI levaria dias. Use os [Dados Abertos do CNPJ](https://arquivos.receitafederal.gov.br/index.php/s/YggdBLfdninEJX9), publicados mensalmente pela Receita Federal: um download de cerca de 5 GB (`Cnaes.zip` e `Estabelecimentos0.zip` a `Estabelecimentos9.zip`) preenche a lista inteira em minutos, sem limite de consultas.

```bash
python dados_abertos.py                  # Baixa o mês mais recente e preenche output.xlsx
python dados_abertos.py --mes 2026-09    # Escolhe o mês publicado
python dados_abertos.py --zips PASTA     # Usa arquivos já baixados, sem internet
```

Os arquivos ficam na pasta temporária do sistema (`cnpj_dados_abertos`); use `--pasta` para outro local fora do OneDrive e apague-os depois de gerar o Excel. O download usa uma conexão por vez, retoma de onde parou e aguarda quando o servidor recusa conexões; não abra downloads paralelos, que levam a Receita a bloquear o acesso temporariamente. Os resultados entram no mesmo cache e no mesmo formato do `lote.py`, e resultados da BrasilAPI dos últimos 30 dias são preservados. CNPJs ausentes da base do mês, como os abertos após a publicação, ficam `PENDENTE`: complete-os com `python lote.py`. Os dados refletem a data de publicação do mês escolhido.

## Consulta direta na Receita

Consulta assistida no site oficial da Receita Federal para obter somente:

- **CNAE principal**, com código e descrição.
- **CNAEs secundários**, com código e descrição de todas as atividades informadas.

Python + Playwright + Beautiful Soup. A saída é JSON, sem razão social, endereço, sócios ou outros dados cadastrais. O CNPJ é usado para a consulta e para conferir o comprovante, mas não é incluído na saída.

## Instalação

Requer Python 3.11 ou superior e conexão com a internet para instalar as dependências e consultar a Receita.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
PLAYWRIGHT_SKIP_BROWSER_GC=1 python -m playwright install chromium
```

No Windows, ative com `.venv\Scripts\activate`.

## Consultar um CNPJ

```bash
python consultar.py 00.000.000/0001-91
```

1. O programa valida o CNPJ e abre a página oficial em uma janela do Chromium.
2. O CNPJ já estará preenchido. **Conclua o CAPTCHA manualmente e clique em CONSULTAR.**
3. Quando o comprovante aparecer, o programa confere o CNPJ, extrai os CNAEs, fecha a janela e imprime o JSON no terminal.

O exemplo de CNPJ acima é público; as atividades dos exemplos abaixo são fictícias. Substitua pelo CNPJ que deseja consultar.

Para salvar em arquivo ou aumentar o tempo de espera do CAPTCHA:

```bash
python consultar.py 00.000.000/0001-91 --saida cnaes.json --espera 300
```

O programa recusa sobrescrever arquivos existentes. Instruções e erros vão para stderr; stdout contém somente o JSON. Use Ctrl+C para cancelar.

## Formato da saída

```json
{
  "cnae_principal": {
    "codigo": "62.01-5-01",
    "descricao": "Desenvolvimento de programas de computador sob encomenda"
  },
  "cnaes_secundarios": [
    {
      "codigo": "62.04-0-00",
      "descricao": "Consultoria em tecnologia da informação"
    }
  ]
}
```

Se o comprovante disser “Não informada” no campo de atividades secundárias, `cnaes_secundarios` será `[]`. Conteúdo mascarado, campo ausente ou CNAE sem descrição produz erro, evitando um resultado incompleto silencioso. Anotações de dispensa de alvará são removidas, preservando a descrição da atividade.

## Usar um comprovante HTML salvo

```bash
python consultar.py --html comprovante.html
python consultar.py 00.000.000/0001-91 --html comprovante.html --saida cnaes.json
```

Salve o **HTML do comprovante já renderizado**, com os dois campos visíveis no conteúdo. O HTML inicial da página dinâmica de consulta e arquivos PDF não contêm os elementos esperados. Ao fornecer o CNPJ junto com `--html`, o programa confere se é o mesmo do comprovante.

Para experimentar sem acessar a Receita, use o HTML sintético incluído:

```bash
python consultar.py --html tests/fixtures/comprovante.html
```

## Testes

```bash
python -m unittest discover -s tests -v
```

Os testes usam dados sintéticos e não consultam a internet. Cobrem os dois layouts de tabela, múltiplas atividades, páginas repetidas, acentuação, CNPJ numérico/alfanumérico, campos incompletos, conferência do CNPJ e proteção de arquivos existentes. Para os Dados Abertos, usam ZIPs sintéticos e simulam a queda e a retomada do download.

## Limites e fontes

- A consulta é individual e depende de intervenção humana no CAPTCHA; o programa não tenta resolvê-lo nem usa serviços pagos.
- A Receita pode restringir navegadores automatizados ou ficar indisponível. Nesse caso, consulte pelo navegador habitual e use o HTML renderizado salvo.
- Mudanças na estrutura do comprovante podem exigir ajuste do parser. Comprovantes mascarados ou sem os campos necessários não são tratados como sucesso.
- O formato alfanumérico é validado pelo cálculo de DV da Receita: valor ASCII do caractere menos 48 e módulo 11.
- O projeto não garante permissão para coleta em massa. Use a consulta conforme os termos do serviço.
- HTMLs reais, resultados JSON e ambientes virtuais são ignorados pelo Git por padrão.

Fontes oficiais verificadas em 13/10/2026:

- [Consultar CNPJ — gov.br](https://www.gov.br/pt-br/servicos/consultar-cadastro-nacional-de-pessoas-juridicas)
- [Emissão do comprovante — Receita Federal](https://solucoes.receita.fazenda.gov.br/Servicos/cnpjreva/)
- [Documentação de cálculo do DV](https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/publicacoes/documentos-tecnicos/cnpj)
- [Playwright para Python](https://playwright.dev/python/docs/library)
