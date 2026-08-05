# -*- coding: utf-8 -*-
"""
Gera os endpoints JSON/CSV estaticos em docs/api/, publicados pelo GitHub Pages.

Nao ha servidor: sao arquivos servidos por HTTPS, com CORS liberado pelo proprio
GitHub Pages, o que permite consumir de qualquer sistema (inclusive do navegador).

Chamado automaticamente ao final de robo_ceagesp.py e de importar_excel.py.
Para regerar sozinho, sem consultar o site:

    python gerar_api.py
"""

import json
import os
import re
import unicodedata
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
DIR_API = os.path.join(BASE, "docs", "api")
URL_BASE = "https://tbrena.github.io/robo-ceagesp/api"

FONTE = "CEAGESP - Entreposto Terminal de Sao Paulo (ETSP)"
FONTE_URL = "https://ceagesp.gov.br/cotacoes/"


def apelido(texto):
    """'FILE DE TILAPIA' -> 'file-de-tilapia'"""
    t = unicodedata.normalize("NFKD", texto or "")
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", t.lower())).strip("-")


def iso(data_br):
    return datetime.strptime(data_br, "%d/%m/%Y").strftime("%Y-%m-%d")


def _registro(r):
    return {
        "data": iso(r["data"]),
        "produto": r["produto"],
        "produto_id": apelido(r["produto"]),
        "classificacao": r.get("classificacao") or "-",
        "unidade": r.get("unidade_peso") or "KG",
        "menor": r.get("menor"),
        "comum": r.get("comum"),
        "maior": r.get("maior"),
    }


def _escrever(nome, conteudo):
    caminho = os.path.join(DIR_API, nome)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as f:
        if nome.endswith(".json"):
            json.dump(conteudo, f, ensure_ascii=False, indent=1)
            f.write("\n")
        else:
            f.write(conteudo)
    return caminho


def gerar(historico):
    if not historico:
        return None

    registros = sorted((_registro(r) for r in historico),
                       key=lambda r: (r["data"], r["produto"], r["classificacao"]))
    datas = sorted({r["data"] for r in registros})
    produtos = sorted({r["produto"] for r in registros})
    agora = datetime.now().astimezone().replace(microsecond=0).isoformat()

    def envelope(**extra):
        base = {
            "fonte": FONTE,
            "fonte_url": FONTE_URL,
            "categoria": "PESCADOS",
            "moeda": "BRL",
            "observacao": ("Precos de venda no atacado, em R$ por unidade indicada. "
                           "'comum' e o valor mais praticado. Levantamento proprio a "
                           "partir dos boletins publicos da CEAGESP; nao e publicacao oficial."),
            "atualizado_em": agora,
        }
        base.update(extra)
        return base

    # /api/ -> indice autoexplicativo
    _escrever("index.json", envelope(
        descricao="API estatica de cotacoes de pescado no atacado da CEAGESP.",
        periodo={"inicio": datas[0], "fim": datas[-1]},
        totais={"boletins": len(datas), "registros": len(registros),
                "produtos": len(produtos)},
        endpoints=[
            {"url": URL_BASE + "/index.json", "descricao": "este indice"},
            {"url": URL_BASE + "/cotacoes.json", "descricao": "historico completo"},
            {"url": URL_BASE + "/cotacoes.csv", "descricao": "historico completo em CSV (UTF-8, virgula)"},
            {"url": URL_BASE + "/ultimo.json", "descricao": "ultimo boletim, com variacao sobre o anterior"},
            {"url": URL_BASE + "/produtos.json", "descricao": "produtos monitorados e seus identificadores"},
            {"url": URL_BASE + "/produto/{produto_id}.json", "descricao": "serie historica de um produto"},
        ],
        atualizacao="Duas vezes por dia util, as 09:00 e as 16:00 (horario de Brasilia). "
                    "A CEAGESP publica o boletim as segundas, quartas e sextas.",
    ))

    # /api/cotacoes.json
    _escrever("cotacoes.json", envelope(total=len(registros), dados=registros))

    # /api/cotacoes.csv
    colunas = ["data", "produto", "produto_id", "classificacao", "unidade",
               "menor", "comum", "maior"]
    linhas = [",".join(colunas)]
    for r in registros:
        linhas.append(",".join(
            '"%s"' % r[c] if isinstance(r[c], str) else ("" if r[c] is None else "%.2f" % r[c])
            for c in colunas))
    _escrever("cotacoes.csv", "\r\n".join(linhas) + "\r\n")

    # /api/ultimo.json -> com variacao percentual sobre o boletim anterior
    ultima, anterior = datas[-1], (datas[-2] if len(datas) > 1 else None)
    def do_dia(d):
        return [r for r in registros if r["data"] == d]
    antes = {(r["produto"], r["classificacao"]): r for r in do_dia(anterior)} if anterior else {}

    itens = []
    for r in do_dia(ultima):
        item = dict(r)
        base = antes.get((r["produto"], r["classificacao"]))
        if base and base["comum"] and r["comum"] is not None:
            item["variacao_percentual"] = round((r["comum"] - base["comum"]) / base["comum"] * 100, 2)
            item["comum_anterior"] = base["comum"]
        else:
            item["variacao_percentual"] = None
            item["comum_anterior"] = None
        itens.append(item)
    _escrever("ultimo.json", envelope(data=ultima, data_anterior=anterior,
                                      total=len(itens), dados=itens))

    # /api/produtos.json
    _escrever("produtos.json", envelope(total=len(produtos), dados=[
        {
            "produto": p,
            "produto_id": apelido(p),
            "url": "%s/produto/%s.json" % (URL_BASE, apelido(p)),
            "boletins": len({r["data"] for r in registros if r["produto"] == p}),
        }
        for p in produtos
    ]))

    # /api/produto/<id>.json
    for p in produtos:
        serie = [{k: v for k, v in r.items() if k not in ("produto", "produto_id")}
                 for r in registros if r["produto"] == p]
        _escrever(os.path.join("produto", apelido(p) + ".json"),
                  envelope(produto=p, produto_id=apelido(p),
                           total=len(serie), dados=serie))

    return DIR_API


if __name__ == "__main__":
    import sys
    sys.path.insert(0, BASE)
    from robo_ceagesp import ler_historico
    print("API gerada em:", gerar(ler_historico()))
