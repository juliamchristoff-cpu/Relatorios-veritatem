#!/usr/bin/env python3
"""
Gerador de relatórios de tráfego pago - Veritatem
=================================================

Lê os exports do Gerenciador de Anúncios do Meta e do Google Ads (CSV ou XLSX)
e gera um PDF no padrão Veritatem (estrutura herdada do Reportei).

Uso:
  python3 gerar_relatorio.py --cliente "PMD Eletro" --tipo semanal \
      --inicio 02/10/2026 --fim 08/10/2026 \
      --arquivos export1.csv export2.xlsx ... --saida /mnt/user-data/outputs

Cada arquivo é identificado automaticamente pelas colunas (campanhas, anúncios,
dia, idade, gênero, região, dispositivo, hora, grupos de anúncios, palavras-chave).
Seções sem arquivo correspondente simplesmente não aparecem no PDF.

O script nunca inventa números: tudo vem dos arquivos. Os avisos impressos no
final (AVISOS) devem ser repassados para quem pediu o relatório.
"""
import argparse
import base64
import csv
import io
import json
import math
import os
import re
import sys
import unicodedata
from datetime import datetime

import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")

# ---------------------------------------------------------------- identidade Veritatem
FONTES = os.path.join(ASSETS, "fontes")
TINTA = "#151B2B"         # texto principal
GRAFITE = "#1F2A44"       # azul-marinho: cards de destaque, barras e rodapé
LOGO_CINZA = "#3D5A8C"
PEDRA = "#6B7280"         # texto secundário
NEVOA = "#F7F7F5"         # fundo da página
PAPEL = "#FFFFFF"
LINHA = "#E4E4E1"
BRONZE = "#3D5A8C"        # azul de acento
BRONZE_CLARO = "#C9D4E8"  # azul claro (rótulos sobre fundo escuro)
TRILHO = "#ECECEA"
CAPA_DE, CAPA_ATE = "#2A3A5C", "#3E5784"   # degradê da capa
SERIE_1, SERIE_2, SERIE_3 = GRAFITE, BRONZE, "#93A6C8"
FUNIL_TONS = ["#1F2A44", "#26365A", "#2E4270", "#354D7E", "#3D5A8C", "#5A75A3"]


AVISOS = []


def aviso(msg):
    AVISOS.append(msg)


# ---------------------------------------------------------------- utilidades
def norm(s):
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"\s+", " ", s)
    return s


def norm_header(s):
    s = norm(s)
    s = re.sub(r"\((brl|usd|eur)\)$", "", s).strip()
    s = re.sub(r"\s*\((brl|usd|eur)\)\s*", " ", s).strip()
    return s


PT_DEC = re.compile(r"^-?(r\$\s*)?[\d.]*\d,\d+\s*%?$", re.I)


def detectar_locale(df):
    """True se os números do arquivo usam vírgula decimal (padrão BR)."""
    amostra = 0
    for col in df.columns:
        for v in df[col].head(60):
            if isinstance(v, str) and PT_DEC.match(v.strip()):
                return True
            amostra += 1
    return False


def num(v, br=False):
    if v is None:
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", "--", "-", " --", "nan", "NaN", "None"):
        return float("nan")
    s = s.replace("R$", "").replace("US$", "").replace("$", "").replace("%", "")
    s = s.replace("<", "").replace(">", "").replace(" ", "").replace(" ", "")
    if br:
        s = s.replace(".", "").replace(",", ".")
    else:
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def ok(x):
    return x is not None and not (isinstance(x, float) and math.isnan(x))


def div(a, b):
    if not ok(a) or not ok(b) or b == 0:
        return float("nan")
    return a / b


def br_num(x, casas=2):
    s = f"{x:,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def f_int(x):
    if not ok(x):
        return "-"
    return br_num(round(x), 0)


def f_money(x):
    if not ok(x):
        return "-"
    if 0 < abs(x) < 0.01:
        return "R$" + f_dec(x, 4)
    return "R$" + br_num(x, 2)


def f_dec(x, casas=2):
    if not ok(x):
        return "-"
    s = br_num(x, casas)
    if "," in s:
        s = s.rstrip("0").rstrip(",")
    return s


def f_pct(x):
    """x já em pontos percentuais (ex.: 4.92)."""
    if not ok(x):
        return "-"
    return f_dec(x, 2) + "%"


def f_smart(x):
    if not ok(x):
        return "-"
    if abs(x - round(x)) < 1e-9:
        return f_int(x)
    return f_dec(x, 2)


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def img_b64(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


# ---------------------------------------------------------------- leitura
def ler_bruto(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, header=None, dtype=object)
    raw = open(path, "rb").read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        txt = raw.decode("utf-16")
    else:
        for enc in ("utf-8-sig", "latin-1"):
            try:
                txt = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    linhas = txt.splitlines()
    amostra = "\n".join(linhas[:15])
    seps = {s: amostra.count(s) for s in (",", ";", "\t")}
    sep = max(seps, key=seps.get)
    rows = list(csv.reader(io.StringIO(txt), delimiter=sep))
    largura = max(len(r) for r in rows) if rows else 0
    rows = [r + [None] * (largura - len(r)) for r in rows]
    return pd.DataFrame(rows, dtype=object)


# aliases: chave canônica -> (exatos, prefixos)
META = {
    "campanha": (["nome da campanha", "campaign name"], []),
    "conjunto": (["nome do conjunto de anuncios", "ad set name"], []),
    "anuncio": (["nome do anuncio", "ad name"], []),
    "criativo": (["creative id", "id do criativo", "creative_id"], []),
    "gasto": (["valor usado", "amount spent", "valor gasto"], ["valor usado", "amount spent"]),
    "impressoes": (["impressoes", "impressions"], []),
    "alcance": (["alcance", "reach"], []),
    "frequencia": (["frequencia", "frequency"], []),
    "cliques_todos": (["cliques (todos)", "clicks (all)"], []),
    "cliques_link": (["cliques no link", "link clicks"], []),
    "resultados": (["resultados", "results"], []),
    "indicador": (["indicador de resultados", "indicador de resultado", "result indicator",
                   "tipo de resultado", "result type"], []),
    "conversas": (["conversas por mensagem iniciadas", "conversas iniciadas por mensagem",
                   "messaging conversations started"], ["conversas por mensagem iniciadas"]),
    "visitas_perfil": (["visitas ao perfil do instagram", "instagram profile visits"], []),
    "leads": (["leads", "cadastros"], []),
    "compras": (["compras", "purchases"], []),
    "lpv": (["visualizacoes da pagina de destino", "landing page views"], []),
    "eng_pagina": (["engajamento com a pagina", "page engagement"], []),
    "eng_publicacao": (["engajamento com a publicacao", "post engagement"], []),
    "reacoes": (["reacoes a publicacao", "reacoes ao post", "post reactions"], []),
    "salvos": (["publicacoes salvas", "post saves"], []),
    "compart": (["compartilhamentos da publicacao", "post shares"], []),
    "comentarios": (["comentarios na publicacao", "post comments"], []),
    "curtidas": (["curtidas na pagina", "page likes", "seguidores ou curtidas na pagina"], []),
    "thruplays": (["15-second thruplays", "thruplays de 15 segundos", "thruplays"], []),
    "inicio": (["inicio dos relatorios", "reporting starts"], []),
    "fim": (["termino dos relatorios", "reporting ends"], []),
    "dia": (["dia", "day"], []),
    "idade": (["idade", "age"], []),
    "genero": (["genero", "gender"], []),
    "regiao": (["regiao", "region"], []),
    "plataforma_disp": (["plataforma do dispositivo", "device platform"], []),
    "hora": ([], ["hora do dia", "time of day"]),
}

GOOGLE = {
    "campanha": (["campanha", "campaign"], []),
    "grupo": (["grupo de anuncios", "ad group"], []),
    "palavra": (["palavra-chave da rede de pesquisa", "palavra-chave", "keyword",
                 "search keyword", "palavra-chave de pesquisa"], []),
    "correspondencia": (["tipo de corresp.", "tipo de correspondencia",
                         "correspondencia da palavra-chave da rede de pesquisa",
                         "match type", "search keyword match type"], ["tipo de corresp"]),
    "dia": (["dia", "day"], []),
    "dispositivo": (["dispositivo", "device"], []),
    "custo": (["custo", "cost"], []),
    "impressoes": (["impr.", "impressoes", "impressions", "impr"], []),
    "cliques": (["cliques", "clicks"], []),
    "conversoes": (["conversoes", "conversions", "conv."], []),
    "todas_conv": (["todas as conv.", "todas as conversoes", "all conv.", "all conversions"], []),
    "taxa_conv": (["taxa de conv.", "taxa de conversao", "conv. rate", "conversion rate"], []),
}

GOOGLE_COLS_FORTES = {"custo", "cost", "impr.", "cpc med.", "cpc medio", "avg. cpc", "conversoes", "conversions"}
META_COLS_FORTES = {"valor usado", "amount spent", "alcance", "reach", "frequencia", "frequency"}


def mapear(colunas, tabela):
    mapa = {}
    normcols = [norm_header(c) for c in colunas]
    for chave, (exatos, prefixos) in tabela.items():
        for i, nc in enumerate(normcols):
            if nc in exatos:
                mapa[chave] = colunas[i]
                break
        if chave not in mapa:
            for i, nc in enumerate(normcols):
                if any(nc.startswith(p) for p in prefixos):
                    mapa[chave] = colunas[i]
                    break
    return mapa


def mapear_google_topo(colunas):
    top = absol = None
    for c in colunas:
        nc = norm_header(c)
        if not any(k in nc for k in ("impr", "%", "taxa", "parc", "rate")):
            continue
        if "abs" in nc and ("top" in nc or "topo" in nc):
            absol = absol or c
        elif "parte sup" in nc or "superior" in nc or "(top)" in nc or ("top" in nc and "abs" not in nc and "impr" in nc):
            top = top or c
    return top, absol


def achar_cabecalho(bruto):
    conhecidos = set()
    for t in (META, GOOGLE):
        for ex, pre in t.values():
            conhecidos.update(ex)
    melhor, melhor_i = 0, 0
    for i in range(min(15, len(bruto))):
        vals = [norm_header(v) for v in bruto.iloc[i].tolist() if v is not None and str(v).strip()]
        score = sum(1 for v in vals if v in conhecidos)
        if score > melhor:
            melhor, melhor_i = score, i
    return melhor_i


def carregar(path):
    bruto = ler_bruto(path)
    h = achar_cabecalho(bruto)
    cab, vistos = [], {}
    for i, c in enumerate(bruto.iloc[h].tolist()):
        nome = str(c).strip() if c is not None and str(c).strip() not in ("", "nan") else f"col{i}"
        if nome in vistos:  # coluna repetida no export (ex.: "Código da moeda" duas vezes)
            vistos[nome] += 1
            nome = f"{nome} ({vistos[nome]})"
        else:
            vistos[nome] = 1
        cab.append(nome)
    df = bruto.iloc[h + 1:].copy()
    df.columns = cab
    df = df.dropna(how="all")
    df = df.loc[:, [c for c in df.columns if not c.startswith("col")]]
    normcols = {norm_header(c) for c in df.columns}
    eh_google = bool(normcols & GOOGLE_COLS_FORTES) and not bool(normcols & {"valor usado", "amount spent"})
    eh_meta = bool(normcols & META_COLS_FORTES) and not eh_google
    if not eh_meta and not eh_google:
        if any(n.startswith("valor usado") or n.startswith("amount spent") for n in normcols):
            eh_meta = True
    return df, ("meta" if eh_meta else "google" if eh_google else None)


def preparar(df, tabela, numericas):
    mapa = mapear(list(df.columns), tabela)
    br = detectar_locale(df)
    out = pd.DataFrame(index=df.index)
    for chave, col in mapa.items():
        if chave in numericas:
            out[chave] = df[col].map(lambda v: num(v, br))
        else:
            out[chave] = df[col].map(lambda v: None if v is None or (isinstance(v, float) and math.isnan(v)) else str(v).strip())
    return out, mapa, br


META_NUM = {"gasto", "impressoes", "alcance", "frequencia", "cliques_todos", "cliques_link", "resultados",
            "conversas", "visitas_perfil", "leads", "compras", "lpv", "eng_pagina", "eng_publicacao",
            "reacoes", "salvos", "compart", "comentarios", "curtidas", "thruplays"}
GOOGLE_NUM = {"custo", "impressoes", "cliques", "conversoes", "todas_conv", "topo", "topo_abs", "taxa_conv"}


def classificar_meta(m):
    if "hora" in m:
        return "hora"
    if "idade" in m and "genero" in m:
        return "idade_genero"
    if "idade" in m:
        return "idade"
    if "genero" in m:
        return "genero"
    if "regiao" in m:
        return "regiao"
    if "plataforma_disp" in m:
        return "dispositivo"
    if "dia" in m:
        return "dia"
    if "anuncio" in m:
        return "anuncios"
    if "conjunto" in m:
        return "conjuntos"
    if "campanha" in m:
        return "campanhas"
    return "conta"


def classificar_google(m):
    if "palavra" in m:
        return "palavras"
    if "dia" in m:
        return "dia"
    if "dispositivo" in m:
        return "dispositivo"
    if "grupo" in m:
        return "grupos"
    if "campanha" in m:
        return "campanhas"
    return "conta"


def separar_resumo_meta(df, chave_nome):
    """Linha de resumo do export do Meta: nome vazio. Retorna (linhas, resumo|None)."""
    if chave_nome not in df.columns:
        return df, None
    vazio = df[chave_nome].isna() | (df[chave_nome].astype(str).str.strip().isin(["", "None", "nan"]))
    resumo = df[vazio]
    linhas = df[~vazio]
    if len(resumo) >= 1 and len(linhas) >= 1:
        return linhas, resumo.iloc[0]
    return linhas, None


def remover_totais_google(df):
    if df.empty:
        return df
    mask = pd.Series(False, index=df.index)
    for c in df.columns:
        mask |= df[c].astype(str).map(norm).str.startswith("total")
    return df[~mask]


def carregar_todos(arquivos):
    dados = {"meta": {}, "google": {}}
    for p in arquivos:
        try:
            df, plataforma = carregar(p)
        except Exception as e:  # noqa: BLE001
            aviso(f"Não consegui ler o arquivo {os.path.basename(p)} ({e}).")
            continue
        if plataforma == "meta":
            d, mapa, _ = preparar(df, META, META_NUM)
            tipo = classificar_meta(d.columns)
        elif plataforma == "google":
            df = remover_totais_google(df)
            d, mapa, br = preparar(df, GOOGLE, GOOGLE_NUM)
            top, absol = mapear_google_topo(list(df.columns))
            if top:
                d["topo"] = df[top].map(lambda v: num(v, br))
            if absol:
                d["topo_abs"] = df[absol].map(lambda v: num(v, br))
            d = remover_totais_google(d)
            tipo = classificar_google(d.columns)
        else:
            aviso(f"Não reconheci o arquivo {os.path.basename(p)} como export do Meta ou do Google Ads.")
            continue
        if tipo in dados[plataforma]:
            aviso(f"Dois arquivos do mesmo tipo ({plataforma} / {tipo}); usei o último: {os.path.basename(p)}.")
        dados[plataforma][tipo] = d
        print(f"  {os.path.basename(p)} -> {plataforma.upper()} / {tipo}", file=sys.stderr)
    return dados


# ---------------------------------------------------------------- rótulos
def rotulo_resultado(ind):
    if not ind:
        return ""
    s = norm(ind)
    regras = [
        ("messaging_conversation_started", "Conversas iniciadas por mensagem"),
        ("conversas", "Conversas iniciadas por mensagem"),
        ("profile_visit", "Visitas ao perfil"), ("visitas ao perfil", "Visitas ao perfil"),
        ("landing_page_view", "Visualizações da página"),
        ("lead", "Leads"), ("purchase", "Compras"), ("compra", "Compras"),
        ("link_click", "Cliques no link"), ("cliques no link", "Cliques no link"),
        ("thruplay", "ThruPlays"), ("video_view", "Visualizações do vídeo"),
        ("post_engagement", "Engajamentos"), ("engajamento", "Engajamentos"),
        ("page_like", "Curtidas na página"), ("like", "Curtidas na página"),
        ("reach", "Alcance"), ("alcance", "Alcance"),
    ]
    for chave, rot in regras:
        if chave in s:
            return rot
    return str(ind).split(":")[-1].split(".")[-1].replace("_", " ").capitalize()


def limpar_regiao(r):
    r = re.sub(r"\s*\((state|estado|province)\)\s*", "", str(r), flags=re.I).strip()
    mapa = {"federal district": "Distrito Federal", "unknown": "Desconhecida"}
    r = mapa.get(r.lower(), r)
    return r[:1].upper() + r[1:]


def limpar_genero(g):
    m = {"female": "Feminino", "male": "Masculino", "unknown": "Desconhecido",
         "feminino": "Feminino", "masculino": "Masculino", "desconhecido": "Desconhecido",
         "mulheres": "Feminino", "homens": "Masculino"}
    return m.get(norm(g), str(g))


def limpar_plataforma(p):
    m = {"mobile_app": "Celular (app)", "mobile_web": "Celular (navegador)", "desktop": "Computador",
         "aplicativo para celular": "Celular (app)", "web para celular": "Celular (navegador)",
         "computador": "Computador"}
    return m.get(norm(p), str(p))


def hora_de(v):
    m = re.match(r"\s*(\d{1,2})", str(v))
    return int(m.group(1)) if m else None


MESES = {"jan": 1, "fev": 2, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "mai": 5, "may": 5, "jun": 6,
         "jul": 7, "ago": 8, "aug": 8, "set": 9, "sep": 9, "out": 10, "oct": 10, "nov": 11, "dez": 12, "dec": 12}


def data_de(v):
    if v is None:
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return pd.Timestamp(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{1,2}) de ([a-zç]+)\.? de (\d{4})", norm(s))
    if m and m.group(2)[:3] in MESES:
        return pd.Timestamp(int(m.group(3)), MESES[m.group(2)[:3]], int(m.group(1)))
    try:
        return pd.to_datetime(s, dayfirst=True).normalize()
    except Exception:  # noqa: BLE001
        return None


# ---------------------------------------------------------------- gráficos
from matplotlib import font_manager as _fm  # noqa: E402

for _f in ("Manrope-Regular.ttf", "Manrope-SemiBold.ttf"):
    try:
        _fm.fontManager.addfont(os.path.join(FONTES, _f))
    except Exception:  # noqa: BLE001
        pass

plt.rcParams.update({
    "font.family": ["Manrope", "Inter", "Liberation Sans", "DejaVu Sans"],
    "font.size": 9, "axes.edgecolor": LINHA, "axes.labelcolor": PEDRA,
    "xtick.color": PEDRA, "ytick.color": PEDRA, "svg.fonttype": "path", "legend.fontsize": 8.5,
})


def fig_svg(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="svg", bbox_inches="tight", transparent=True)
    plt.close(fig)
    return "data:image/svg+xml;base64," + base64.b64encode(buf.getvalue()).decode()


def _eixo(ax):
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(LINHA)
    ax.grid(axis="y", color="#E6E6E3", linewidth=0.8, linestyle=(0, (2, 3)))
    ax.set_axisbelow(True)
    ax.tick_params(length=0, pad=6)


def _fmt_y(dinheiro=False, pct=False):
    from matplotlib.ticker import FuncFormatter
    if dinheiro:
        return FuncFormatter(lambda v, _: "R$" + br_num(v, 0))
    if pct:
        return FuncFormatter(lambda v, _: f_dec(v, 2) + "%")
    return FuncFormatter(lambda v, _: br_num(v, 0))


def _legenda(ax, handles):
    ax.legend(handles=handles, loc="lower left", bbox_to_anchor=(0, 1.02), ncol=len(handles),
              frameon=False, handlelength=1.0, handletextpad=0.5, columnspacing=1.4, borderaxespad=0)


def grafico_linha(datas, series, figsize=(5.6, 2.6)):
    """series: lista de (nome, valores, cor, eixo('y1'|'y2'), formato('int'|'money'|'pct'))."""
    from matplotlib.ticker import MaxNLocator
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    _eixo(ax)
    ax2 = None
    x = list(range(len(datas)))
    handles = []
    for i, (nome, vals, cor, eixo, fmt) in enumerate(series):
        alvo = ax
        if eixo == "y2":
            if ax2 is None:
                ax2 = ax.twinx()
                for s in ("top", "right", "left", "bottom"):
                    ax2.spines[s].set_visible(False)
                ax2.tick_params(length=0, pad=6)
            alvo = ax2
        alvo.yaxis.set_major_formatter(_fmt_y(pct=fmt == "pct", dinheiro=fmt == "money"))
        if fmt == "int":
            alvo.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=5))
        else:
            alvo.yaxis.set_major_locator(MaxNLocator(nbins=5))
        h, = alvo.plot(x, vals, color=cor, linewidth=2.2, label=nome, solid_capstyle="round")
        if i == 0:
            alvo.fill_between(x, vals, color=cor, alpha=0.05, linewidth=0)
        if x:
            alvo.plot([x[-1]], [vals[-1]], "o", color=cor, markersize=4.5)
        handles.append(h)
        alvo.set_ylim(bottom=0)
    passo = max(1, len(x) // 8)
    ax.set_xticks(x[::passo])
    ax.set_xticklabels([d.strftime("%d/%m") for d in datas][::passo])
    ax.set_xlim(-0.2, max(0.2, len(x) - 0.8))
    _legenda(ax, handles)
    return fig_svg(fig)


def grafico_barras(categorias, series, figsize=(5.6, 2.6), rot=0):
    """series: lista de (nome, valores, cor)."""
    import numpy as np
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    _eixo(ax)
    n = len(series)
    larg = 0.78 / n
    x = np.arange(len(categorias))
    handles = []
    for i, (nome, vals, cor) in enumerate(series):
        handles.append(ax.bar(x + (i - (n - 1) / 2) * larg, vals, width=larg * 0.9, color=cor, label=nome))
    ax.set_xticks(x)
    ax.set_xticklabels(categorias, rotation=rot)
    ax.yaxis.set_major_formatter(_fmt_y())
    _legenda(ax, handles)
    return fig_svg(fig)


# ---------------------------------------------------------------- componentes HTML
MESES_NOME = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
              "setembro", "outubro", "novembro", "dezembro"]


def periodo_extenso(ini, fim):
    try:
        d1, m1, a1 = (int(x) for x in ini.split("/"))
        d2, m2, a2 = (int(x) for x in fim.split("/"))
    except Exception:  # noqa: BLE001
        return f"{ini} a {fim}"
    if (m1, a1) == (m2, a2):
        return f"{d1:02d} a {d2:02d} de {MESES_NOME[m2 - 1]} de {a2}"
    if a1 == a2:
        return f"{d1:02d} de {MESES_NOME[m1 - 1]} a {d2:02d} de {MESES_NOME[m2 - 1]} de {a2}"
    return f"{d1:02d} de {MESES_NOME[m1 - 1]} de {a1} a {d2:02d} de {MESES_NOME[m2 - 1]} de {a2}"


def fontes_css():
    css = ""
    for fam, arq, peso in (("Outfit", "Outfit-Medium.ttf", 500), ("Outfit", "Outfit-SemiBold.ttf", 600),
                           ("Manrope", "Manrope-Regular.ttf", 400), ("Manrope", "Manrope-SemiBold.ttf", 600),
                           ("Manrope", "Manrope-Bold.ttf", 700)):
        p = os.path.join(FONTES, arq)
        if os.path.exists(p):
            b64 = base64.b64encode(open(p, "rb").read()).decode()
            css += (f"@font-face{{font-family:'{fam}';src:url(data:font/ttf;base64,{b64}) format('truetype');"
                    f"font-weight:{peso};font-style:normal}}")
    return css


CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Manrope,Inter,Arial,sans-serif;color:%(tinta)s;background:%(nevoa)s;width:1080px;padding:36px 36px 0;-webkit-font-smoothing:antialiased}
h1,h2,h3,.tile strong,.kpi strong,.funil-linha .val,.bloco-cab b{font-family:Outfit,Manrope,sans-serif}
.capa{position:relative;overflow:hidden;background:linear-gradient(115deg,%(capa_de)s 0%%,%(capa_ate)s 100%%);color:#fff;border-radius:22px;padding:40px 48px 50px}
.capa .marca{height:30px;position:relative}
.capa .simbolo{display:none}
.capa .ondas{position:absolute;left:0;right:0;bottom:0;width:100%%;height:100%%}
.capa .eyebrow{position:relative;margin-top:78px;font-size:12px;letter-spacing:.22em;text-transform:uppercase;color:%(bronze_claro)s;font-weight:600}
.capa h1{position:relative;font-size:52px;font-weight:600;letter-spacing:-.01em;margin-top:12px;line-height:1.05}
.capa .periodo{position:relative;margin-top:14px;font-size:17px;color:rgba(255,255,255,.78)}
.capa .chips{position:relative;margin-top:26px;display:flex;gap:8px}
.capa .chips span{border:1px solid rgba(255,255,255,.22);border-radius:999px;padding:6px 14px;font-size:12px;color:rgba(255,255,255,.85)}
.capa .fio{position:absolute;left:48px;bottom:0;width:120px;height:4px;background:%(bronze_claro)s;opacity:.55}
.plat{margin-top:46px}
.plat-cab{display:flex;align-items:flex-end;justify-content:space-between;border-bottom:1px solid %(linha)s;padding-bottom:14px;margin-bottom:22px}
.plat-cab h2{font-size:34px;font-weight:600;letter-spacing:-.01em;line-height:1}
.plat-cab span{font-size:11.5px;letter-spacing:.18em;text-transform:uppercase;color:%(pedra)s;font-weight:600}
.tiles{display:flex;gap:14px}
.tile{flex:1;background:%(grafite)s;color:#fff;border-radius:18px;padding:22px 24px}
.tile span{font-size:11.5px;letter-spacing:.1em;text-transform:uppercase;color:%(bronze_claro)s;font-weight:600}
.tile strong{display:block;font-size:40px;font-weight:600;margin-top:14px;letter-spacing:-.01em;line-height:1}
.tile small{display:block;margin-top:10px;font-size:12.5px;color:rgba(255,255,255,.72)}
.tile.acento{background:%(bronze)s}
.tile.acento span{color:#fff;opacity:.9}
.tile.acento small{color:rgba(255,255,255,.85)}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:14px}
.kpi{background:%(papel)s;border:1px solid %(linha)s;border-radius:14px;padding:16px 18px}
.kpi span{font-size:12px;color:%(pedra)s;font-weight:600}
.kpi strong{display:block;font-size:25px;font-weight:600;margin-top:8px;line-height:1.1}
.bloco{margin-top:36px}
.bloco-cab{display:flex;align-items:center;gap:12px;margin-bottom:14px}
.bloco-cab b{font-size:14px;color:%(bronze)s;font-weight:600}
.bloco-cab h3{font-size:19px;font-weight:600}
.bloco-cab i{flex:1;height:1px;background:%(linha)s}
.cartao{background:%(papel)s;border:1px solid %(linha)s;border-radius:16px;padding:20px 22px}
.cartao.tab{padding:4px 10px}
.cartao h4{font-size:13px;font-weight:700;margin-bottom:10px}
.cartao h4 small{font-weight:400;color:%(pedra)s;margin-left:6px}
table{width:100%%;border-collapse:collapse;font-size:11.5px}
th{color:%(pedra)s;font-weight:600;font-size:9.5px;letter-spacing:.06em;text-transform:uppercase;padding:13px 9px 11px;text-align:right;border-bottom:1px solid %(linha)s;vertical-align:bottom}
td{padding:11px 9px;text-align:right;border-bottom:1px solid #EFEFEC;vertical-align:middle;font-variant-numeric:tabular-nums}
tr:last-child td{border-bottom:none}
th:first-child,td:first-child{text-align:left}
td.nome{max-width:250px;font-weight:600;font-size:11px;word-break:break-word}
td.sub{min-width:110px;max-width:150px;font-size:10.5px;color:%(pedra)s;text-align:left;word-break:break-word}
td.res small{display:block;color:%(pedra)s;font-size:9.5px;margin-top:2px}
.dois{display:flex;gap:16px;align-items:stretch}
.dois>div{flex:1;min-width:0}
.tres{display:flex;gap:14px}
.tres>div{flex:1;min-width:0}
.graf img{width:100%%;display:block}
.graf+.graf{margin-top:22px}
.funil-linha{display:flex;align-items:center;gap:14px}
.funil-linha .rot{width:178px;font-size:12px;font-weight:600}
.funil-linha .trilho{flex:1;height:30px;background:%(trilho)s;border-radius:8px;overflow:hidden}
.funil-linha .barra{height:100%%;border-radius:8px}
.funil-linha .val{width:92px;text-align:right;font-size:19px;font-weight:600}
.funil-taxa{margin:5px 0 5px 192px;font-size:10.5px;color:%(pedra)s}
.share{display:flex;align-items:center;gap:12px;padding:6px 0}
.share .rot{width:150px;font-size:12px}
.share .trilho{flex:1;height:10px;background:%(trilho)s;border-radius:6px;overflow:hidden}
.share .barra{height:100%%;background:%(grafite)s;border-radius:6px}
.share:first-of-type .barra{background:%(bronze)s}
.share .val{width:120px;text-align:right;font-size:12px;font-weight:700;font-variant-numeric:tabular-nums}
.share .val small{color:%(pedra)s;font-weight:400;margin-left:6px}
.compacto .share{flex-wrap:wrap;gap:6px 10px;padding:7px 0}
.compacto .share .rot{flex:1;width:auto}
.compacto .share .val{width:auto}
.compacto .share .trilho{order:3;flex-basis:100%%;height:8px}
.kpis.k3{grid-template-columns:repeat(3,1fr)}
.ad{display:flex;align-items:center;gap:9px;min-width:185px}
.ad img{width:44px;height:44px;object-fit:cover;border-radius:8px;flex:none;background:%(trilho)s}
.galeria{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:16px}
.criativo{background:%(papel)s;border:1px solid %(linha)s;border-radius:16px;overflow:hidden}
.criativo img{width:100%%;aspect-ratio:4/5;object-fit:cover;display:block;background:%(trilho)s}
.criativo div{padding:12px 14px 14px}
.criativo b{display:block;font-size:11px;line-height:1.35;min-height:30px}
.criativo strong{display:block;font-family:Outfit,Manrope,sans-serif;font-size:22px;font-weight:600;margin-top:8px}
.criativo small{display:block;font-size:10.5px;color:%(pedra)s;margin-top:2px}
.glossario{display:grid;grid-template-columns:1fr 1fr;gap:12px 30px;font-size:11.5px;line-height:1.55;color:#4B5263}
.glossario b{color:%(tinta)s}
.rodape{margin-top:48px;background:%(grafite)s;border-radius:22px 22px 0 0;padding:26px 40px;display:flex;align-items:center;justify-content:space-between;color:rgba(255,255,255,.66);font-size:11.5px}
.rodape img{height:24px}
""" % dict(tinta=TINTA, nevoa=NEVOA, grafite=GRAFITE, bronze=BRONZE, bronze_claro=BRONZE_CLARO,
           linha=LINHA, pedra=PEDRA, papel=PAPEL, trilho=TRILHO, capa_de=CAPA_DE, capa_ate=CAPA_ATE)


def tiles(itens):
    """itens: (rótulo, valor, subtítulo, acento)."""
    h = '<div class="tiles">'
    for rot, val, sub, acento in itens:
        h += (f'<div class="tile{" acento" if acento else ""}"><span>{esc(rot)}</span>'
              f'<strong>{esc(val)}</strong>{f"<small>{esc(sub)}</small>" if sub else ""}</div>')
    return h + "</div>"


def kpis(itens):
    k3 = " k3" if len(itens) % 4 and len(itens) % 3 == 0 else ""
    return f'<div class="kpis{k3}">' + "".join(
        f'<div class="kpi"><span>{esc(r)}</span><strong>{esc(v)}</strong></div>' for r, v in itens) + "</div>"


class Blocos:
    def __init__(self):
        self.n = 0

    def __call__(self, titulo, conteudo):
        self.n += 1
        return (f'<div class="bloco"><div class="bloco-cab"><b>{self.n:02d}</b><h3>{esc(titulo)}</h3><i></i></div>'
                f"{conteudo}</div>")


def cartao(conteudo, tab=False, titulo=None):
    t = f"<h4>{titulo}</h4>" if titulo else ""
    return f'<div class="cartao{" tab" if tab else ""}">{t}{conteudo}</div>'


def tabela(cabecalho, linhas, classes=None):
    classes = classes or {}
    h = "<table><thead><tr>" + "".join(f"<th>{esc(c)}</th>" for c in cabecalho) + "</tr></thead><tbody>"
    for ln in linhas:
        h += "<tr>"
        for i, cel in enumerate(ln):
            cls = classes.get(i, "")
            if isinstance(cel, dict):  # miniatura + texto
                h += (f'<td class="{cls}"><div class="ad"><img src="{cel["img"]}" alt="">'
                      f'<span>{esc(cel["texto"])}</span></div></td>')
                continue
            if isinstance(cel, tuple):  # (valor, legenda)
                h += f'<td class="res {cls}">{esc(cel[0])}<small>{esc(cel[1])}</small></td>'
            else:
                h += f'<td class="{cls}">{esc(cel)}</td>'
        h += "</tr>"
    return h + "</tbody></table>"


def funil(etapas):
    etapas = [(r, v) for r, v in etapas if ok(v) and v > 0]
    if not etapas:
        return ""
    maxv = max(v for _, v in etapas)
    h = ""
    for i, (r, v) in enumerate(etapas):
        if i > 0:
            taxa = div(v, etapas[i - 1][1])
            h += f'<div class="funil-taxa">{f_pct(taxa * 100)} da etapa anterior</div>'
        larg = 14 + 86 * (math.log10(v + 1) / math.log10(maxv + 1)) if maxv > 0 else 100
        cor = FUNIL_TONS[min(i, len(FUNIL_TONS) - 1)] if len(etapas) <= len(FUNIL_TONS) else GRAFITE
        if i == len(etapas) - 1:
            cor = BRONZE
        h += (f'<div class="funil-linha"><div class="rot">{esc(r)}</div><div class="trilho">'
              f'<div class="barra" style="width:{larg:.1f}%;background:{cor}"></div></div>'
              f'<div class="val">{f_smart(v)}</div></div>')
    return h


def participacao(itens, formato=f_smart):
    itens = [(r, v) for r, v in itens if ok(v) and v > 0]
    if not itens:
        return '<div style="color:%s;font-size:11px">Sem dados no período.</div>' % PEDRA
    total = sum(v for _, v in itens)
    maxv = max(v for _, v in itens)
    h = ""
    for r, v in itens:
        h += (f'<div class="share"><div class="rot">{esc(r)}</div><div class="trilho">'
              f'<div class="barra" style="width:{v / maxv * 100:.1f}%"></div></div>'
              f'<div class="val">{esc(formato(v))}<small>{f_dec(v / total * 100, 1)}%</small></div></div>')
    return h


def cabecalho_plataforma(nome):
    return (f'<section class="plat"><div class="plat-cab"><h2>{esc(nome)}</h2>'
            f"<span>Desempenho no período</span></div>")


GLOSSARIO_META = [
    ("Alcance", "número de pessoas diferentes que viram os anúncios."),
    ("Impressões", "total de vezes que os anúncios apareceram, contando repetições para a mesma pessoa."),
    ("Frequência", "média de vezes que cada pessoa viu os anúncios (impressões divididas pelo alcance)."),
    ("CPC", "custo médio de cada clique no link do anúncio."),
    ("CPM", "custo para os anúncios aparecerem mil vezes."),
    ("CTR", "percentual de impressões que viraram clique no link."),
]



# ---------------------------------------------------------------- META
def somar(df, cols):
    return {c: (df[c].sum(min_count=1) if c in df.columns else float("nan")) for c in cols}


TIPOS_ACAO = [("eng_pagina", "Engajamento com a página"), ("eng_publicacao", "Engajamento com a publicação"),
              ("cliques_link", "Clique nos links"), ("reacoes", "Reações às postagens"),
              ("comentarios", "Comentários nas postagens"), ("salvos", "Postagens salvas"),
              ("compart", "Compartilhamento das postagens"), ("curtidas", "Curtidas na página"),
              ("thruplays", "ThruPlays (vídeo assistido por 15s)")]

DESTAQUES = [("conversas", "Conversas iniciadas por mensagem", "Conversas iniciadas por mensagem"),
             ("leads", "Leads", "Leads"), ("compras", "Compras", "Compras"),
             ("visitas_perfil", "Visitas ao Perfil do Instagram", "Visitas ao perfil"),
             ("lpv", "Visualizações da página de destino", "Visualizações da página")]


ALCANCE_TOTAL = None  # definido por --alcance-total (alcance da conta, sem duplicidade)
PREVIAS = {}  # id do criativo -> {"uri", "w", "h"}; definido por --previas


def id_criativo(v):
    if v is None:
        return ""
    s = str(v).strip()
    try:
        if "e" in s.lower() or s.endswith(".0"):
            s = str(int(float(s)))
    except ValueError:
        pass
    return s


def carregar_links_previas(caminho):
    """JSON {id_criativo: {"url": ..., "grande": bool}}: imagens carregadas pelo navegador de quem abre o HTML."""
    out = {}
    if not caminho or not os.path.exists(caminho):
        return out
    for cid, info in json.load(open(caminho, encoding="utf-8")).items():
        if isinstance(info, str):
            info = {"url": info}
        if info.get("url"):
            out[id_criativo(cid)] = {"uri": esc(info["url"]), "w": 1000 if info.get("grande") else 64, "h": 0}
    return out


def carregar_previas(pasta):
    from PIL import Image
    out = {}
    if not pasta or not os.path.isdir(pasta):
        return out
    for nome in os.listdir(pasta):
        base, ext = os.path.splitext(nome)
        if ext.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        try:
            im = Image.open(os.path.join(pasta, nome)).convert("RGB")
        except Exception:  # noqa: BLE001
            aviso(f"Prévia ilegível: {nome}")
            continue
        w, h = im.size
        im.thumbnail((640, 800))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=82)
        out[id_criativo(base)] = {"uri": "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(),
                                  "w": w, "h": h}
    return out


def totais_meta(d):
    cols = list(META_NUM)
    camp = d.get("campanhas")
    fonte = None
    if camp is not None:
        linhas, resumo = separar_resumo_meta(camp, "campanha")
        d["campanhas"] = linhas
        if resumo is not None:
            fonte = {c: resumo.get(c, float("nan")) for c in cols}
    if fonte is None and "conta" in d:
        fonte = somar(d["conta"], cols)
        fonte["frequencia"] = float("nan")
    if fonte is None and "anuncios" in d:
        linhas, resumo = separar_resumo_meta(d["anuncios"], "anuncio")
        d["anuncios"] = linhas
        if resumo is not None:
            fonte = {c: resumo.get(c, float("nan")) for c in cols}
    if "anuncios" in d:
        d["anuncios"], _ = separar_resumo_meta(d["anuncios"], "anuncio")
    if fonte is None:
        base = d.get("campanhas") if d.get("campanhas") is not None else d.get("anuncios")
        if base is None:
            return None
        fonte = somar(base, cols)
        fonte["frequencia"] = float("nan")  # recalculada abaixo (impressões / alcance)
        if ALCANCE_TOTAL is None:
          aviso("Meta: o export de campanhas veio sem a linha de resumo. Alcance total e frequência foram "
              "calculados somando as campanhas, o que pode contar a mesma pessoa mais de uma vez. "
              "Para o número exato, exporte marcando 'Incluir linha de resumo'.")
    # resultados por tipo (para destaques sem coluna própria)
    base = d.get("campanhas") if d.get("campanhas") is not None else d.get("anuncios")
    por_tipo = {}
    if base is not None and "resultados" in base.columns and "indicador" in base.columns:
        for _, r in base.iterrows():
            rot = rotulo_resultado(r.get("indicador"))
            if ok(r.get("resultados")):
                por_tipo[rot] = por_tipo.get(rot, 0) + r["resultados"]
    fonte["_por_tipo"] = por_tipo
    if ALCANCE_TOTAL is not None:
        fonte["alcance"] = ALCANCE_TOTAL
        fonte["frequencia"] = float("nan")
    if not ok(fonte.get("frequencia")):
        fonte["frequencia"] = div(fonte.get("impressoes"), fonte.get("alcance"))
    return fonte


ROTULO_CURTO = {"Conversas iniciadas por mensagem": "conversas", "Visitas ao perfil": "visitas ao perfil",
                "Visualizações da página": "visualizações", "Cliques no link": "cliques no link"}


def linhas_entidade(df, nome_col, incluir_conjunto=False):
    linhas = []
    df = df.copy()
    df["_res"] = df["resultados"] if "resultados" in df.columns else float("nan")
    ordem = df.sort_values(["_res", "gasto"] if "gasto" in df.columns else ["_res"], ascending=False, na_position="last")
    for _, r in ordem.iterrows():
        gasto = r.get("gasto", float("nan"))
        res = r.get("resultados", float("nan"))
        rot = rotulo_resultado(r.get("indicador"))
        rot = ROTULO_CURTO.get(rot, rot.lower() if rot else rot)
        clq = r.get("cliques_todos", float("nan"))
        if not ok(clq):
            clq = r.get("cliques_link", float("nan"))
        imp = r.get("impressoes", float("nan"))
        alc = r.get("alcance", float("nan"))
        freq = r.get("frequencia", float("nan"))
        if not ok(freq):
            freq = div(imp, alc)
        ln = [r.get(nome_col) or "-"]
        prev = PREVIAS.get(id_criativo(r.get("criativo"))) if nome_col == "anuncio" else None
        if prev:
            ln = [{"img": prev["uri"], "texto": ln[0]}]
        if incluir_conjunto:
            ln.append(r.get("conjunto") or "-")
        ln += [f_money(gasto), (f_smart(res), rot) if ok(res) else "-",
               (f_money(div(gasto, res)), rot) if ok(res) and res > 0 else "-",
               f_int(alc), f_int(imp), f_dec(freq, 2), f_money(div(gasto, clq)),
               f_money(div(gasto, imp) * 1000 if ok(div(gasto, imp)) else float("nan")),
               f_pct(div(clq, imp) * 100 if ok(div(clq, imp)) else float("nan"))]
        linhas.append(ln)
    return linhas


def agrupar(df, chave, transform=None):
    g = df.copy()
    if transform:
        g[chave] = g[chave].map(transform)
    cols = [c for c in ("gasto", "alcance", "impressoes", "cliques_todos", "cliques_link") if c in g.columns]
    return g.groupby(chave, dropna=True)[cols].sum(min_count=1)


def galeria(df):
    """Até 4 criativos com mais resultados, quando a prévia tem qualidade para aparecer grande."""
    if not PREVIAS or "criativo" not in df.columns:
        return ""
    g = df.copy()
    g["_res"] = g["resultados"] if "resultados" in g.columns else 0
    g = g.sort_values(["_res", "gasto"] if "gasto" in g.columns else ["_res"], ascending=False)
    cards, vistos = [], set()
    for _, r in g.iterrows():
        cid = id_criativo(r.get("criativo"))
        prev = PREVIAS.get(cid)
        if not prev or prev["w"] < 300 or cid in vistos or prev["uri"] in vistos:
            continue
        vistos.update({cid, prev["uri"]})
        res = r.get("resultados")
        rot = rotulo_resultado(r.get("indicador"))
        cards.append(f'<div class="criativo"><img src="{prev["uri"]}" alt=""><div><b>{esc(r.get("anuncio") or "-")}</b>'
                     f'<strong>{f_smart(res) if ok(res) else "-"}</strong><small>{esc(rot)}</small>'
                     f'<small>{f_money(r.get("gasto"))} investidos</small></div></div>')
        if len(cards) == 4:
            break
    return f'<div class="galeria">{"".join(cards)}</div>' if cards else ""


UNIDADE = {"conversas": "por conversa", "leads": "por lead", "compras": "por compra",
           "visitas_perfil": "por visita ao perfil", "lpv": "por visualização da página"}


def secao_meta(d, cliente, destaques_forcados):
    t = totais_meta(d)
    if t is None:
        aviso("Meta: nenhum export de campanhas ou anúncios encontrado; seção do Meta não foi gerada.")
        return ""
    bloco = Blocos()
    base = d.get("campanhas") if d.get("campanhas") is not None else d.get("anuncios")
    gasto, imp, alc = t.get("gasto"), t.get("impressoes"), t.get("alcance")
    clink, ctodos = t.get("cliques_link"), t.get("cliques_todos")

    # resultados principais
    escolhidos = []
    for chave, rot, rot_ind in DESTAQUES:
        if destaques_forcados and chave not in destaques_forcados:
            continue
        v = t.get(chave)
        if not ok(v) or v == 0:
            v = t["_por_tipo"].get(rot_ind, float("nan"))
        if ok(v) and v > 0:
            escolhidos.append((chave, rot, rot_ind, v))
    if not escolhidos and t["_por_tipo"]:
        rot, v = max(t["_por_tipo"].items(), key=lambda kv: kv[1])
        if rot not in ("Alcance", "Cliques no link"):
            escolhidos.append((None, rot, rot, v))
    MSG["meta"] = {"gasto": gasto, "impressoes": imp, "alcance": alc, "cliques": ctodos,
                   "resultados": [(c, v) for c, _, _, v in escolhidos if c in ("conversas", "leads", "compras")]}

    def custo_de(chave, rot_ind, v):
        if base is None or "gasto" not in base.columns:
            return float("nan")
        if "indicador" in base.columns:
            m = base["indicador"].map(rotulo_resultado) == rot_ind
        elif chave and chave in base.columns:
            m = base[chave].fillna(0) > 0
        else:
            return float("nan")
        if not m.any():
            return float("nan")
        return div(base.loc[m, "gasto"].sum(), v)

    n_camp = len(d["campanhas"]) if d.get("campanhas") is not None else 0
    lista_tiles = [("Investimento", f_money(gasto),
                    f"em {n_camp} campanha{'s' if n_camp != 1 else ''}" if n_camp else "no período", False)]
    for i, (chave, rot, rot_ind, v) in enumerate(escolhidos[:2]):
        c = custo_de(chave, rot_ind, v)
        sub = f"{f_money(c)} {UNIDADE.get(chave, 'por resultado')}" if ok(c) else ""
        lista_tiles.append((rot, f_smart(v), sub, i == 0))

    lista_kpis = [("Impressões", f_int(imp)), ("Alcance", f_int(alc)), ("Frequência", f_dec(t.get("frequencia"), 2))]
    if ok(ctodos):
        lista_kpis.append(("Cliques (todos)", f_int(ctodos)))
    ctr = div(clink, imp)
    lista_kpis += [("Cliques no link", f_int(clink)), ("CPC médio (link)", f_money(div(gasto, clink))),
                   ("CPM médio", f_money(div(gasto, imp) * 1000 if ok(div(gasto, imp)) else float("nan"))),
                   ("CTR (link)", f_pct(ctr * 100 if ok(ctr) else float("nan")))]
    for chave, rot, rot_ind, v in escolhidos[2:]:
        lista_kpis.append((rot, f_smart(v)))

    h = cabecalho_plataforma("Meta Ads") + tiles(lista_tiles) + kpis(lista_kpis)

    # funil + evolução diária
    etapas = [("Impressões", imp), ("Alcance", alc), ("Cliques (todos)", ctodos), ("Cliques no link", clink)]
    etapas += [(rot, v) for _, rot, _, v in escolhidos[:2]]
    html_funil = cartao(funil(etapas))
    dia = d.get("dia")
    if dia is not None and "dia" in dia.columns:
        dd = dia.copy()
        dd["data"] = dd["dia"].map(data_de)
        dd = dd.dropna(subset=["data"])
        cols = [c for c in ("gasto", "impressoes", "cliques_todos", "cliques_link") if c in dd.columns]
        g = dd.groupby("data")[cols].sum(min_count=1).sort_index()
        cl = g["cliques_todos"] if "cliques_todos" in g else g.get("cliques_link")
        nome_cl = "Cliques (todos)" if "cliques_todos" in g else "Cliques no link"
        graf = f'<div class="graf"><h4>Investimento por dia</h4><img src="{grafico_linha(list(g.index), [("Investimento", g["gasto"].fillna(0).tolist(), SERIE_1, "y1", "money")], figsize=(5.0, 2.2))}"></div>'
        if cl is not None and "impressoes" in g:
            ctr_d = (cl / g["impressoes"].replace(0, float("nan")) * 100).fillna(0).tolist()
            graf += (f'<div class="graf"><h4>Cliques e CTR por dia</h4><img src="'
                     + grafico_linha(list(g.index), [(nome_cl, cl.fillna(0).tolist(), SERIE_1, "y1", "int"),
                                                      ("CTR", ctr_d, SERIE_2, "y2", "pct")], figsize=(5.0, 2.2))
                     + '"></div>')
        h += bloco("Funil e evolução diária", f'<div class="dois"><div>{html_funil}</div><div>{cartao(graf)}</div></div>')
    else:
        h += bloco("Funil", html_funil)

    acoes = [(rot, t.get(c)) for c, rot in TIPOS_ACAO if ok(t.get(c)) and t.get(c) > 0]
    if acoes:
        acoes.sort(key=lambda x: -x[1])
        h += bloco("Engajamento e ações", cartao(participacao(acoes)))

    ref = base
    so_link = ref is not None and "cliques_todos" not in ref.columns
    cab_base = ["Investimento", "Resultados", "Custo por resultado", "Alcance", "Impressões",
                "Frequência", "CPC (link)" if so_link else "CPC", "CPM", "CTR (link)" if so_link else "CTR (todos)"]
    if d.get("campanhas") is not None and len(d["campanhas"]):
        h += bloco("Campanhas", cartao(tabela(["Campanha"] + cab_base,
                                              linhas_entidade(d["campanhas"], "campanha"), {0: "nome"}), tab=True))
    if d.get("anuncios") is not None and len(d["anuncios"]):
        com_conj = "conjunto" in d["anuncios"].columns
        ads = linhas_entidade(d["anuncios"], "anuncio", incluir_conjunto=com_conj)
        cab = ["Anúncio"] + (["Conjunto"] if com_conj else []) + cab_base
        h += bloco("Anúncios em destaque", galeria(d["anuncios"]) +
                   cartao(tabela(cab, ads[:15], {0: "nome", 1: "sub"} if com_conj else {0: "nome"}), tab=True))

    # público
    publico = []
    lado_a = ""
    if d.get("dispositivo") is not None:
        g = agrupar(d["dispositivo"], "plataforma_disp", limpar_plataforma)
        if "alcance" in g:
            g = g.sort_values("alcance", ascending=False)
            lado_a += cartao(participacao(list(zip(g.index, g["alcance"]))), titulo="Alcance por dispositivo")
    fonte_gen = d.get("genero") if d.get("genero") is not None else d.get("idade_genero")
    if fonte_gen is not None:
        g = agrupar(fonte_gen, "genero", limpar_genero)
        if "alcance" in g:
            g = g.sort_values("alcance", ascending=False)
            lado_a += ('<div style="height:16px"></div>' if lado_a else "") + \
                cartao(participacao(list(zip(g.index, g["alcance"]))), titulo="Alcance por gênero")
    lado_b = ""
    fonte_idade = d.get("idade") if d.get("idade") is not None else d.get("idade_genero")
    if fonte_idade is not None:
        g = agrupar(fonte_idade, "idade").sort_index()
        series = [(n, g[c].tolist(), cor) for n, c, cor in (("Impressões", "impressoes", SERIE_1), ("Alcance", "alcance", SERIE_2)) if c in g]
        lado_b = cartao(f'<div class="graf"><img src="{grafico_barras(list(g.index), series, figsize=(5.0, 2.9))}"></div>',
                        titulo="Impressões e alcance por idade")
    if lado_a and lado_b:
        publico.append(f'<div class="dois"><div>{lado_a}</div><div>{lado_b}</div></div>')
    elif lado_a or lado_b:
        publico.append(lado_a or lado_b)
    if d.get("hora") is not None:
        g = agrupar(d["hora"], "hora", hora_de).reindex(range(24)).fillna(0)
        series = [(n, g[c].tolist(), cor) for n, c, cor in (("Impressões", "impressoes", SERIE_1), ("Alcance", "alcance", SERIE_2)) if c in g]
        publico.append(('<div style="height:16px"></div>' if publico else "") + cartao(
            f'<div class="graf"><img src="{grafico_barras([f"{i:02d}h" for i in g.index], series, figsize=(13.0, 2.9))}"></div>',
            titulo="Impressões e alcance por horário"))
    if publico:
        h += bloco("Público", "".join(publico))

    if d.get("regiao") is not None:
        g = agrupar(d["regiao"], "regiao", limpar_regiao)
        if "alcance" in g:
            g = g.sort_values("alcance", ascending=False)
        linhas = []
        for reg, r in g.head(15).iterrows():
            clq = r.get("cliques_todos", float("nan"))
            if not ok(clq):
                clq = r.get("cliques_link", float("nan"))
            linhas.append([reg, f_money(r.get("gasto")), f_int(r.get("alcance")), f_int(r.get("impressoes")),
                           f_dec(div(r.get("impressoes"), r.get("alcance")), 2), f_int(clq),
                           f_money(div(r.get("gasto"), clq)) if ok(clq) and clq > 0 else "R$0,00",
                           f_money(div(r.get("gasto"), r.get("impressoes")) * 1000),
                           f_pct(div(clq, r.get("impressoes")) * 100 if ok(div(clq, r.get("impressoes"))) else 0)])
        h += bloco("Regiões com maior alcance", cartao(tabela(
            ["Região", "Investimento", "Alcance", "Impressões", "Frequência", "Cliques", "CPC", "CPM", "CTR"],
            linhas, {0: "nome"}), tab=True))

    h += bloco("Como ler os números", cartao('<div class="glossario">' + "".join(
        f"<div><b>{esc(a)}:</b> {esc(b)}</div>" for a, b in GLOSSARIO_META) + "</div>"))
    return h + "</section>"



# ---------------------------------------------------------------- GOOGLE
def palavra_fmt(p, tipo):
    t = norm(tipo or "")
    p = str(p or "-")
    if p.startswith(("[", '"')):
        return p
    if "exat" in t or "exact" in t:
        return f"[{p}]"
    if "frase" in t or "phrase" in t:
        return f'"{p}"'
    return p


def media_pond(df, col):
    if col not in df.columns or "impressoes" not in df.columns:
        return float("nan")
    m = df[col].notna() & df["impressoes"].notna()
    if not m.any() or df.loc[m, "impressoes"].sum() == 0:
        return float("nan")
    return (df.loc[m, col] * df.loc[m, "impressoes"]).sum() / df.loc[m, "impressoes"].sum()


def linha_google(nome, r, tem_todas=True, tem_abs=True):
    custo, imp, clq = r.get("custo"), r.get("impressoes"), r.get("cliques")
    conv, todas = r.get("conversoes"), r.get("todas_conv")
    ln = [nome, f_money(custo), f_int(imp), f_int(clq),
          f_pct(div(clq, imp) * 100 if ok(div(clq, imp)) else float("nan")), f_money(div(custo, clq)),
          f_smart(conv), f_money(div(custo, conv)) if ok(conv) and conv > 0 else "R$0,00"]
    if tem_todas:
        ln += [f_smart(todas), f_money(div(custo, todas)) if ok(todas) and todas > 0 else "R$0,00"]
    if tem_abs:
        ln.append(f_pct(r.get("topo_abs")))
    return ln


CAB_GOOGLE = ["Custo", "Impressões", "Cliques", "CTR (Taxa de cliques)", "CPC médio", "Conversões",
              "Custo por conversão"]


def cab_google(tem_todas, tem_abs):
    c = list(CAB_GOOGLE)
    if tem_todas:
        c += ["Todas as conversões", "Custo por todas as conversões"]
    if tem_abs:
        c.append("% de anúncios na 1ª posição")
    return c


def tabela_google(df, chave, titulo_col, limite=None, ordenar="cliques"):
    cols = [c for c in ("custo", "impressoes", "cliques", "conversoes", "todas_conv") if c in df.columns]
    g = df.copy()
    if chave == "palavra" and "correspondencia" in g.columns:
        g["palavra"] = [palavra_fmt(p, t) for p, t in zip(g["palavra"], g["correspondencia"])]
    agg = g.groupby(chave)[cols].sum(min_count=1)
    for c in ("topo_abs",):
        if c in g.columns:
            agg[c] = g.groupby(chave).apply(lambda x: media_pond(x, c))
    entrega = agg[[c for c in ("impressoes", "custo", "cliques") if c in agg.columns]].fillna(0).sum(axis=1)
    if (entrega > 0).any():
        agg = agg[entrega > 0]
    agg = agg.sort_values(ordenar if ordenar in agg else cols[0], ascending=False)
    if limite:
        agg = agg.head(limite)
    tem_todas = "todas_conv" in g.columns and g["todas_conv"].notna().any()
    tem_abs = "topo_abs" in g.columns and g["topo_abs"].notna().any()
    linhas = [linha_google(nome, r, tem_todas, tem_abs) for nome, r in agg.iterrows()]
    return tabela([titulo_col] + cab_google(tem_todas, tem_abs), linhas, {0: "nome"})


def secao_google(d, cliente):
    camp, grp = d.get("campanhas"), d.get("grupos")
    if camp is not None and grp is not None and "impressoes" not in camp.columns and "impressoes" in grp.columns:
        if len(camp) == 1:
            camp = camp.copy()
            camp["impressoes"] = grp["impressoes"].sum(min_count=1)
            d["campanhas"] = camp
            aviso("Google: o export de campanhas veio sem impressões; usei o total do relatório de grupos (mesma campanha).")
        else:
            aviso("Google: o export de campanhas veio sem impressões; inclua a coluna 'Impr.' para ter CTR e impressões por campanha.")
    base = d.get("campanhas")
    if base is None:
        base = d.get("grupos") if d.get("grupos") is not None else d.get("conta")
    if base is None:
        if d:
            aviso("Google: faltou o export de campanhas; seção do Google não foi gerada.")
        return ""
    bloco = Blocos()
    soma = lambda c: base[c].sum(min_count=1) if c in base.columns else float("nan")  # noqa: E731
    custo, imp, clq, conv, todas = soma("custo"), soma("impressoes"), soma("cliques"), soma("conversoes"), soma("todas_conv")
    topo = media_pond(base, "topo")
    ctr = div(clq, imp)
    tx = div(conv, clq)
    if "taxa_conv" in base.columns and "conversoes" in base.columns:
        # mesma conta do Google Ads: conversões / interações (interações = conversões / taxa)
        m = base["taxa_conv"].notna() & (base["taxa_conv"] > 0) & base["conversoes"].notna()
        inter = (base.loc[m, "conversoes"] / (base.loc[m, "taxa_conv"] / 100)).sum()
        if inter > 0:
            tx = base.loc[m, "conversoes"].sum() / inter

    MSG["google"] = {"custo": custo, "impressoes": imp, "cliques": clq, "conversoes": conv}
    n_camp = len(d["campanhas"]) if d.get("campanhas") is not None else 0
    cpconv = div(custo, conv)
    lista_tiles = [("Investimento", f_money(custo),
                    f"em {n_camp} campanha{'s' if n_camp != 1 else ''}" if n_camp else "no período", False),
                   ("Conversões", f_smart(conv),
                    f"{f_money(cpconv)} por conversão" if ok(conv) and conv > 0 else "", True)]
    lista_kpis = [("Impressões", f_int(imp)), ("Cliques", f_int(clq)),
                  ("CTR", f_pct(ctr * 100 if ok(ctr) else float("nan"))), ("CPC médio", f_money(div(custo, clq))),
                  ("Taxa de conversão", f_pct(tx * 100 if ok(tx) else float("nan")))]
    if ok(topo):
        lista_kpis.append(("Impressões na parte superior", f_pct(topo)))
    if ok(todas):
        lista_kpis.append(("Todas as conversões", f_smart(todas)))
    h = cabecalho_plataforma("Google Ads") + tiles(lista_tiles) + kpis(lista_kpis)

    dia = d.get("dia")
    if dia is not None:
        dd = dia.copy()
        dd["data"] = dd["dia"].map(data_de)
        dd = dd.dropna(subset=["data"])
        cols = [c for c in ("custo", "impressoes", "cliques", "conversoes") if c in dd.columns]
        g = dd.groupby("data")[cols].sum(min_count=1).sort_index().fillna(0)
        datas = list(g.index)
        tam = (3.5, 2.5)
        ctr_d = (g["cliques"] / g["impressoes"].replace(0, float("nan")) * 100).fillna(0).tolist()
        cpc_d = (g["custo"] / g["cliques"].replace(0, float("nan"))).fillna(0).tolist()
        graficos = [("Cliques e CTR", grafico_linha(datas, [("Cliques", g["cliques"].tolist(), SERIE_1, "y1", "int"),
                                                            ("CTR", ctr_d, SERIE_2, "y2", "pct")], figsize=tam))]
        if "conversoes" in g:
            tx_d = (g["conversoes"] / g["cliques"].replace(0, float("nan")) * 100).fillna(0).tolist()
            graficos.append(("Conversões e taxa", grafico_linha(datas, [("Conversões", g["conversoes"].tolist(), SERIE_1, "y1", "int"),
                                                                        ("Taxa", tx_d, SERIE_2, "y2", "pct")], figsize=tam)))
        graficos.append(("Custo e CPC médio", grafico_linha(datas, [("Custo", g["custo"].tolist(), SERIE_1, "y1", "money"),
                                                                    ("CPC médio", cpc_d, SERIE_2, "y2", "money")], figsize=tam)))
        h += bloco("Evolução diária", '<div class="tres">' + "".join(
            "<div>" + cartao('<div class="graf"><img src="' + s + '"></div>', titulo=esc(t_)) + "</div>" for t_, s in graficos) + "</div>")

    if d.get("campanhas") is not None:
        h += bloco("Campanhas", cartao(tabela_google(d["campanhas"], "campanha", "Campanha", ordenar="custo"), tab=True))
    if d.get("grupos") is not None:
        h += bloco("Grupos de anúncios", cartao(tabela_google(d["grupos"], "grupo", "Grupo de anúncios", ordenar="custo"), tab=True))
    if d.get("palavras") is not None:
        h += bloco("Palavras-chave com mais cliques",
                   cartao(tabela_google(d["palavras"], "palavra", "Palavra-chave", limite=10), tab=True))
    if d.get("dispositivo") is not None:
        cols = [c for c in ("custo", "cliques", "conversoes") if c in d["dispositivo"]]
        g = d["dispositivo"].groupby("dispositivo")[cols].sum(min_count=1)
        partes = []
        for col, tit, fmt in (("custo", "Custo", f_money), ("cliques", "Cliques", f_smart), ("conversoes", "Conversões", f_smart)):
            if col in g:
                gg = g.sort_values(col, ascending=False)
                partes.append(f'<div class="compacto">{cartao(participacao(list(zip(gg.index, gg[col])), fmt), titulo=tit)}</div>')
        h += bloco("Dispositivos", '<div class="tres">' + "".join(partes) + "</div>")
    return h + "</section>"



# ---------------------------------------------------------------- mensagem para o cliente
MSG = {}  # totais usados na mensagem de WhatsApp (preenchidos pelas seções)
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro",
         "outubro", "novembro", "dezembro"]


def _n(x):
    return f"{int(round(x)):,}".replace(",", ".")


def _rs(x):
    return "R$" + f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _linha(v, sing, plur):
    if not ok(v):
        return None
    return f"> {_n(v)} {sing if round(v) == 1 else plur}"


def montar_mensagem(cliente, tipo, inicio, fim):
    """Mensagem padrão da Veritatem para WhatsApp (negrito com *, citação com >)."""
    from datetime import datetime, timedelta, timezone
    hora = datetime.now(timezone(timedelta(hours=-3))).hour
    saud = "Bom dia" if hora < 12 else ("Boa tarde" if hora < 18 else "Boa noite")
    ini, fi = inicio[:5], fim[:5]
    if tipo == "mensal":
        mes = MESES[int(inicio[3:5]) - 1]
        abertura = f"Aqui estão os principais destaques do mês de {mes}. 👇"
    else:
        dias = (datetime.strptime(fim, "%d/%m/%Y") - datetime.strptime(inicio, "%d/%m/%Y")).days + 1
        abertura = f"Aqui estão os principais destaques dos últimos {dias} dias. 👇"
    partes = [f"{saud}, pessoal! Tudo bem?", abertura,
              f"*VISÃO GERAL | {cliente.upper()}*\n*Período:* {ini} até {fi}."]
    g = MSG.get("google")
    if g:
        linhas = [_linha(g["impressoes"], "impressão dos anúncios.", "impressões dos anúncios."),
                  _linha(g["cliques"], "clique nos anúncios.", "cliques nos anúncios."),
                  _linha(g["conversoes"], "conversão", "conversões")]
        partes.append("*GOOGLE ADS:*\n" + "\n".join(x for x in linhas if x))
        if ok(g["custo"]):
            partes.append(f"*VALOR INVESTIDO:*\n{_rs(g['custo'])}")
    m = MSG.get("meta")
    if m:
        linhas = [_linha(m["impressoes"], "impressão dos anúncios.", "impressões dos anúncios."),
                  _linha(m["alcance"], "pessoa alcançada.", "pessoas alcançadas."),
                  _linha(m["cliques"], "clique nos anúncios.", "cliques nos anúncios.")]
        nomes = {"conversas": ("conversa iniciada.", "conversas iniciadas."), "leads": ("lead.", "leads."),
                 "compras": ("compra.", "compras.")}
        linhas += [_linha(v, *nomes[c]) for c, v in m["resultados"]]
        partes.append("*META ADS:*\n" + "\n".join(x for x in linhas if x))
        if ok(m["gasto"]):
            partes.append(f"*VALOR INVESTIDO:*\n{_rs(m['gasto'])}")
    return "\n\n".join(partes)


def ondas_svg(w=1008, h=330, linhas=30):
    """Linhas onduladas claras (fundo da capa), no estilo topográfico da Veritatem."""
    import math
    caminhos = []
    for k in range(linhas):
        t = k / (linhas - 1)
        base = h * (0.62 + 0.45 * t)
        pts = []
        for i in range(0, 101):
            x = w * i / 100
            u = i / 100
            morro = math.exp(-((u - 0.58) / 0.13) ** 2) * (0.62 - 0.35 * t) * h * 0.55
            morro += math.exp(-((u - 0.78) / 0.10) ** 2) * (0.45 - 0.25 * t) * h * 0.4
            morro += math.exp(-((u - 0.36) / 0.16) ** 2) * (0.28 - 0.15 * t) * h * 0.4
            y = base - morro + 6 * math.sin(u * 9 + k * 0.35)
            pts.append(f"{x:.1f},{y:.1f}")
        caminhos.append(f'<polyline points="{" ".join(pts)}"/>')
    return (f'<svg class="ondas" viewBox="0 0 {w} {h}" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">'
            f'<g fill="none" stroke="#C9D4E8" stroke-opacity=".22" stroke-width="1">{"".join(caminhos)}</g></svg>')


# ---------------------------------------------------------------- montagem
def montar_html(cliente, tipo, inicio, fim, dados, destaques):
    logo = img_b64(os.path.join(ASSETS, "logo_horizontal_branca.png"))
    simbolo = img_b64(os.path.join(ASSETS, "simbolo_branco.png"))
    titulo = f"Relatório {'Mensal' if tipo == 'mensal' else 'Semanal'} - {cliente}"
    secoes, plataformas = [], []
    if dados["google"]:
        s = secao_google(dados["google"], cliente)
        if s:
            secoes.append(s)
            plataformas.append("Google Ads")
    if dados["meta"]:
        s = secao_meta(dados["meta"], cliente, destaques)
        if s:
            secoes.append(s)
            plataformas.append("Meta Ads")
    if not secoes:
        raise SystemExit("Nenhum dado de Meta ou Google reconhecido nos arquivos enviados.")
    capa = (f'<header class="capa">{ondas_svg()}<img class="marca" src="{logo}">'
            f'<div class="eyebrow">Relatório {"mensal" if tipo == "mensal" else "semanal"} de desempenho</div>'
            f"<h1>{esc(cliente)}</h1><div class=\"periodo\">{esc(periodo_extenso(inicio, fim))}</div>"
            f'<div class="chips">{"".join(f"<span>{p}</span>" for p in plataformas)}</div><div class="fio"></div></header>')
    rodape = (f'<footer class="rodape"><img src="{logo}"><div>Dados de {esc(inicio)} a {esc(fim)}, '
              f'extraídos de {" e ".join(plataformas)}.</div></footer>')
    return (f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>{esc(titulo)}</title>'
            f"<style>{fontes_css()}{CSS}</style></head><body>{capa}{''.join(secoes)}{rodape}</body></html>"), titulo


def html_para_pdf(html, destino):
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            nav = p.chromium.launch()
            pag = nav.new_page(viewport={"width": 1080, "height": 1200})
            pag.set_content(html, wait_until="load")
            altura = pag.evaluate("document.documentElement.scrollHeight")
            pag.pdf(path=destino, width="1080px", height=f"{altura + 2}px", print_background=True,
                    margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
            nav.close()
        return "chromium"
    except Exception as e:  # noqa: BLE001
        print(f"  Chromium indisponível ({e}); tentando WeasyPrint.", file=sys.stderr)
    from weasyprint import HTML  # pip install weasyprint
    html = html.replace("<style>", "<style>@page{size:1136px 1600px;margin:0}")
    HTML(string=html).write_pdf(destino)
    aviso("PDF gerado em páginas (sem Chromium); o layout pode diferir levemente do padrão de página única.")
    return "weasyprint"


def medir_altura(html):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        nav = p.chromium.launch()
        pag = nav.new_page(viewport={"width": 1080, "height": 1200})
        pag.route("**/*", lambda r: r.abort() if r.request.url.startswith("http") else r.continue_())
        pag.set_content(html, wait_until="load")
        h = pag.evaluate("document.documentElement.scrollHeight")
        nav.close()
    return h


def ocultar_cliente(cliente, extra=""):
    """Seções escondidas por cliente (references/clientes.json) + --ocultar."""
    out = {}
    cfg_path = os.path.join(os.path.dirname(ASSETS), "references", "clientes.json")
    alvo = norm(cliente)
    try:
        cfg = json.load(open(cfg_path, encoding="utf-8"))
    except Exception:  # noqa: BLE001
        cfg = {}
    for nome, regras in cfg.items():
        if nome.startswith("_"):
            continue
        if norm(nome) in alvo or alvo in norm(nome):
            for plat, secs in regras.get("ocultar", {}).items():
                out.setdefault(plat, set()).update(secs)
    for item in [x.strip() for x in extra.split(",") if x.strip()]:
        plat, _, sec = item.partition(":")
        out.setdefault(plat, set()).add(sec)
    return out


def nome_arquivo(cliente, tipo, inicio, fim):
    limpo = re.sub(r"[^A-Za-z0-9]+", "_", unicodedata.normalize("NFKD", cliente).encode("ascii", "ignore").decode()).strip("_")
    return f"Veritatem_{limpo}_{'Mensal' if tipo == 'mensal' else 'Semanal'}_{inicio.replace('/', '-')}_a_{fim.replace('/', '-')}.pdf"


def main():
    ap = argparse.ArgumentParser(description="Relatório Veritatem")
    ap.add_argument("--cliente", required=True)
    ap.add_argument("--tipo", choices=["semanal", "mensal"], required=True)
    ap.add_argument("--inicio", required=True, help="dd/mm/aaaa")
    ap.add_argument("--fim", required=True, help="dd/mm/aaaa")
    ap.add_argument("--arquivos", nargs="+", required=True)
    ap.add_argument("--saida", default=".")
    ap.add_argument("--destaques", default="",
                    help="Opcional. Resultados em destaque no Meta: conversas,leads,compras,visitas_perfil,lpv")
    ap.add_argument("--alcance-total", type=float, default=None,
                    help="Alcance total da conta (sem duplicidade), quando o export veio sem linha de resumo")
    ap.add_argument("--previas", default=None,
                    help="Pasta com as imagens dos criativos, nomeadas pelo ID do criativo (ex.: 2296495411107944.jpg)")
    ap.add_argument("--previas-links", default=None,
                    help="JSON {id_criativo: {url, grande}}. Gera um HTML para abrir no navegador e salvar como PDF")
    ap.add_argument("--ocultar", default="",
                    help="Seções a esconder, ex.: meta:dispositivo,meta:hora,meta:regiao (soma-se ao clientes.json)")
    ap.add_argument("--html", action="store_true", help="Também salva o HTML (para conferência)")
    a = ap.parse_args()

    global ALCANCE_TOTAL, PREVIAS
    ALCANCE_TOTAL = a.alcance_total
    PREVIAS = carregar_previas(a.previas)
    if a.previas_links:
        PREVIAS.update(carregar_links_previas(a.previas_links))
    print("Arquivos identificados:", file=sys.stderr)
    dados = carregar_todos(a.arquivos)
    for plat, secs in ocultar_cliente(a.cliente, a.ocultar).items():
        for sec in secs:
            dados.get(plat, {}).pop(sec, None)
    destaques = {x.strip() for x in a.destaques.split(",") if x.strip()}
    html, titulo = montar_html(a.cliente, a.tipo, a.inicio, a.fim, dados, destaques)
    os.makedirs(a.saida, exist_ok=True)
    destino = os.path.join(a.saida, nome_arquivo(a.cliente, a.tipo, a.inicio, a.fim))
    if a.previas_links:
        destino = destino[:-4] + ".html"
        altura = medir_altura(html)
        extra = ("@page{size:1080px %dpx;margin:0}html,body{-webkit-print-color-adjust:exact;print-color-adjust:exact}"
                 "@media print{body{margin:0}}" % (altura + 4))
        open(destino, "w", encoding="utf-8").write(html.replace("<style>", "<style>" + extra, 1))
        motor = "html-para-navegador"
    else:
        if a.html:
            open(destino[:-4] + ".html", "w", encoding="utf-8").write(html)
        motor = html_para_pdf(html, destino)
    mensagem = montar_mensagem(a.cliente, a.tipo, a.inicio, a.fim)
    open(os.path.splitext(destino)[0] + "_mensagem.txt", "w", encoding="utf-8").write(mensagem)
    print(json.dumps({"pdf": destino, "titulo": titulo, "motor": motor, "mensagem": mensagem,
                      "secoes": {k: sorted(v.keys()) for k, v in dados.items()},
                      "avisos": AVISOS}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
