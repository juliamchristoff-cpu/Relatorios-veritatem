#!/usr/bin/env python3
"""
Converte a planilha do script do Google Ads ("Veritatem Google Ads - <conta>", baixada
como .xlsx) nos CSVs que o gerar_relatorio.py lê, já filtrados pelo período pedido.

Uso:
  python3 google_planilha.py --planilha arquivo.xlsx --inicio 09/10/2026 --fim 15/10/2026 --saida pasta_csv

A planilha traz os dados dia a dia; aqui eles são somados no período. O % de impressões
na parte superior e no topo absoluto é a média ponderada pelas impressões.
"""
import argparse
import json
import os
import sys
from datetime import datetime

import pandas as pd

MATCH = {"EXACT": "Correspondência exata", "PHRASE": "Correspondência de frase", "BROAD": "Correspondência ampla"}
DISP = {"MOBILE": "Celulares", "DESKTOP": "Computadores", "TABLET": "Tablets",
        "CONNECTED_TV": "TVs", "OTHER": "Outros"}

CAMPOS = {  # nome no fim do campo do Google Ads -> chave interna
    "date": "dia", "campaign.name": "campanha", "ad_group.name": "grupo",
    "keyword.text": "palavra", "keyword.match_type": "correspondencia", "device": "dispositivo",
    "cost_micros": "custo_micros", "impressions": "impressoes", "clicks": "cliques",
    "interactions": "interacoes", "conversions": "conversoes", "all_conversions": "todas_conv",
    "top_impression_percentage": "topo", "absolute_top_impression_percentage": "topo_abs",
    "advertising_channel_type": "tipo",
}


def normalizar(df):
    ren = {}
    for c in df.columns:
        cs = str(c).strip().lower().replace(" ", "_")
        for fim, chave in CAMPOS.items():
            if cs == fim or cs.endswith("." + fim) or cs.endswith(fim) and fim in ("cost_micros", "all_conversions"):
                ren[c] = chave
                break
    alias = {"day": "dia", "date": "dia", "dia": "dia", "campaign": "campanha", "campanha": "campanha",
             "ad_group": "grupo", "grupo_de_anúncios": "grupo", "keyword": "palavra", "match_type": "correspondencia",
             "device": "dispositivo", "cost": "custo", "custo": "custo", "impr.": "impressoes", "impressions": "impressoes",
             "conversions": "conversoes", "all_conv.": "todas_conv", "interactions": "interacoes"}
    for c in df.columns:
        if c not in ren:
            k = str(c).strip().lower().replace(" ", "_")
            if k in alias:
                ren[c] = alias[k]
    df = df.rename(columns=ren)
    if "custo_micros" in df:
        df["custo"] = pd.to_numeric(df["custo_micros"], errors="coerce") / 1_000_000
    for c in ("impressoes", "cliques", "interacoes", "conversoes", "todas_conv", "topo", "topo_abs"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "dia" in df:
        df["dia"] = pd.to_datetime(df["dia"], errors="coerce").dt.normalize()
    return df


def agregar(df, chaves):
    soma = [c for c in ("custo", "impressoes", "cliques", "interacoes", "conversoes", "todas_conv") if c in df]
    g = df.groupby(chaves, dropna=False)
    out = g[soma].sum(min_count=1).reset_index()
    for c in ("topo", "topo_abs"):
        if c in df and "impressoes" in df:
            w = df.assign(_p=df[c] * df["impressoes"]).groupby(chaves, dropna=False)["_p"].sum()
            imp = g["impressoes"].sum()
            out[c] = (w / imp.replace(0, float("nan"))).values * 100
    if "conversoes" in out and "interacoes" in out:
        out["taxa_conv"] = out["conversoes"] / out["interacoes"].replace(0, float("nan")) * 100
    return out


SAIDA = {"campanha": "Campanha", "grupo": "Grupo de anúncios", "palavra": "Palavra-chave",
         "correspondencia": "Tipo de corresp.", "dia": "Dia", "dispositivo": "Dispositivo",
         "custo": "Custo", "impressoes": "Impr.", "cliques": "Cliques", "conversoes": "Conversões",
         "todas_conv": "Todas as conv.", "taxa_conv": "Taxa de conv.",
         "topo": "Taxa de impr. (parte sup.)", "topo_abs": "Taxa de impr. (topo abs.)"}


def salvar(df, path):
    cols = [c for c in SAIDA if c in df.columns]
    out = df[cols].rename(columns=SAIDA)
    if "Dia" in out:
        out["Dia"] = pd.to_datetime(out["Dia"]).dt.strftime("%Y-%m-%d")
    out.round(4).to_csv(path, index=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--planilha", required=True)
    ap.add_argument("--inicio", required=True)
    ap.add_argument("--fim", required=True)
    ap.add_argument("--saida", required=True)
    a = ap.parse_args()
    ini = datetime.strptime(a.inicio, "%d/%m/%Y")
    fim = datetime.strptime(a.fim, "%d/%m/%Y")
    abas = pd.read_excel(a.planilha, sheet_name=None)
    os.makedirs(a.saida, exist_ok=True)
    gerados, avisos = [], []

    info = abas.get("info")
    if info is not None and len(info.columns) >= 2:
        kv = {str(r.iloc[0]).strip(): str(r.iloc[1]).strip() for _, r in pd.concat(
            [pd.DataFrame([info.columns.tolist()], columns=info.columns), info]).iterrows()}
        de, ate = kv.get("dados de"), kv.get("dados até")
        try:
            if de and ini < pd.to_datetime(de):
                avisos.append(f"A planilha só tem dados a partir de {de}; o período pedido começa antes.")
            if ate and fim > pd.to_datetime(ate):
                avisos.append(f"A planilha tem dados até {ate} (atualizada em {kv.get('atualizado em')}); "
                              "o período pedido vai além. Rode o script de novo ou ajuste a data final.")
        except Exception:  # noqa: BLE001
            pass

    def periodo(nome):
        df = abas.get(nome)
        if df is None or df.empty:
            return None
        df = normalizar(df)
        if "dia" in df:
            df = df[(df["dia"] >= ini) & (df["dia"] <= fim)]
        return df if len(df) else None

    camp = periodo("campanhas")
    if camp is not None:
        p = os.path.join(a.saida, "google_campanhas.csv")
        salvar(agregar(camp, ["campanha"]), p)
        gerados.append(p)
        p = os.path.join(a.saida, "google_dia.csv")
        salvar(agregar(camp, ["dia"]), p)
        gerados.append(p)
    else:
        avisos.append("Sem dados de campanhas do Google no período.")
    grp = periodo("grupos")
    if grp is not None:
        p = os.path.join(a.saida, "google_grupos.csv")
        salvar(agregar(grp, ["grupo"]), p)
        gerados.append(p)
    kw = periodo("palavras")
    if kw is not None:
        kw["correspondencia"] = kw.get("correspondencia", pd.Series(dtype=str)).map(lambda x: MATCH.get(str(x), x))
        p = os.path.join(a.saida, "google_palavras.csv")
        salvar(agregar(kw, ["palavra", "correspondencia"]), p)
        gerados.append(p)
    dv = periodo("dispositivos")
    if dv is not None:
        dv["dispositivo"] = dv["dispositivo"].map(lambda x: DISP.get(str(x), x))
        p = os.path.join(a.saida, "google_dispositivo.csv")
        salvar(agregar(dv, ["dispositivo"]), p)
        gerados.append(p)
    print(json.dumps({"csvs": gerados, "avisos": avisos}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
