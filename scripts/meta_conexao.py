#!/usr/bin/env python3
"""
Converte as respostas da conexão do Meta Ads (ferramenta ads_get_ad_entities)
em CSVs no mesmo formato do export do Gerenciador, para o gerar_relatorio.py.

Uso:
  python3 meta_conexao.py --entrada pasta_com_json --saida pasta_csv

Arquivos esperados na pasta de entrada (cada um é a resposta da ferramenta,
salva como veio, com ou sem o envelope {"ad_entities": "..."}):
  conta.json        nível ad_account, sem detalhamento (totais exatos da conta)
  campanhas.json    nível campaign
  anuncios.json     nível ad
  dia.json          nível ad_account com time_increment "1"
  idade.json        nível ad_account, breakdowns ["age"]
  genero.json       nível ad_account, breakdowns ["gender"]
  regiao.json       nível ad_account, breakdowns ["region"]
  dispositivo.json  nível ad_account, breakdowns ["device_platform"]
  hora.json         nível ad_account, breakdowns ["hourly_stats_aggregated_by_audience_time_zone"]
Só conta.json + campanhas.json são obrigatórios; os demais viram seções opcionais.
"""
import argparse
import csv
import json
import os
import sys

# campo da conexão -> coluna do export (nomes em inglês do Gerenciador)
METRICAS = [
    ("amount_spent", "Amount spent (BRL)"),
    ("reach", "Reach"),
    ("impressions", "Impressions"),
    ("frequency", "Frequency"),
    ("clicks", "Clicks (all)"),
    ("link_click", "Link clicks"),
    ("page_engagement", "Page engagement"),
    ("post_reaction", "Post reactions"),
    ("post_save", "Post saves"),
    ("post_shares", "Post shares"),
    ("comment", "Post comments"),
    ("video_thruplay_watched_actions", "ThruPlays"),
    ("lead", "Leads"),
    ("landing_page_view", "Landing page views"),
    ("omni_purchase", "Purchases"),
]

DIMENSOES = {
    "dia": ("date_start", "Day"),
    "idade": ("age", "Age"),
    "genero": ("gender", "Gender"),
    "regiao": ("region", "Region"),
    "dispositivo": ("device_platform", "Device platform"),
    "hora": ("hourly_stats_aggregated_by_audience_time_zone", "Time of day (viewer's time zone)"),
}


def ler(path):
    txt = open(path, encoding="utf-8").read().strip()
    d = json.loads(txt)
    if isinstance(d, dict):
        d = d.get("ad_entities", d.get("data", d))
    if isinstance(d, str):
        d = json.loads(d)
    if isinstance(d, dict):
        d = [d]
    return d


def valor(v):
    if v is None:
        return ""
    if isinstance(v, dict):
        if "value" in v:
            return v["value"]
        if "values" in v:
            return sum(float(x.get("value") or 0) for x in v["values"])
    return v


def resultado(r):
    res = r.get("results")
    if not isinstance(res, dict):
        return "", ""
    total = sum(float(x.get("value") or 0) for x in res.get("values") or [])
    return total, res.get("indicator", "")


def linha_metricas(r):
    return {col: valor(r.get(campo)) for campo, col in METRICAS}


def escrever(path, linhas, colunas):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=colunas, extrasaction="ignore")
        w.writeheader()
        for ln in linhas:
            w.writerow(ln)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entrada", required=True)
    ap.add_argument("--saida", required=True)
    a = ap.parse_args()
    os.makedirs(a.saida, exist_ok=True)
    met_cols = [c for _, c in METRICAS]
    gerados = []

    def arq(nome):
        p = os.path.join(a.entrada, nome + ".json")
        return p if os.path.exists(p) else None

    # campanhas, com a linha de resumo da conta no topo (alcance exato, sem duplicidade)
    if arq("campanhas"):
        linhas = []
        if arq("conta"):
            conta = ler(arq("conta"))[0]
            ln = {"Campaign name": ""}
            ln.update(linha_metricas(conta))
            linhas.append(ln)
        else:
            print("AVISO: sem conta.json; o alcance total será a soma das campanhas.", file=sys.stderr)
        for r in ler(arq("campanhas")):
            ln = {"Campaign name": r.get("name", "")}
            ln.update(linha_metricas(r))
            ln["Results"], ln["Result indicator"] = resultado(r)
            linhas.append(ln)
        p = os.path.join(a.saida, "meta_campanhas.csv")
        escrever(p, linhas, ["Campaign name", "Results", "Result indicator"] + met_cols)
        gerados.append(p)

    if arq("anuncios"):
        linhas = []
        for r in ler(arq("anuncios")):
            ln = {"Ad name": r.get("name", ""), "Ad set name": r.get("adset_name", ""),
                  "Creative ID": r.get("creative_id", "")}
            ln.update(linha_metricas(r))
            ln["Results"], ln["Result indicator"] = resultado(r)
            linhas.append(ln)
        p = os.path.join(a.saida, "meta_anuncios.csv")
        escrever(p, linhas, ["Ad name", "Ad set name", "Creative ID", "Results", "Result indicator"] + met_cols)
        gerados.append(p)

    for nome, (campo, col) in DIMENSOES.items():
        if not arq(nome):
            continue
        linhas = []
        for r in ler(arq(nome)):
            ln = {col: r.get(campo, "")}
            ln.update(linha_metricas(r))
            linhas.append(ln)
        p = os.path.join(a.saida, f"meta_{nome}.csv")
        escrever(p, linhas, [col] + met_cols)
        gerados.append(p)

    print(json.dumps({"csvs": gerados}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
