# -*- coding: utf-8 -*-
"""
TERCEIRA ETAPA - Corpo HTML do email diario. Universo principal: clientes C.I.
+ "Depósito Trânsito a Base BR" (tipo B), que COMPOE o total em todo lugar.
Conteudo: histórico por tipo (com Tr. p. CI e Tr. p.Base); movimentações com
reconciliação; destaque do Base BR; quadro dos depósitos danificados/perdidos.
Persiste relatorios/equip/<DATA>.json. NAO escreve historico.json.

Uso: python gerar_email.py [caminho_csv]
"""
import os, re, sys, glob, json
import datetime as dt
from collections import Counter
import pandas as pd

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
PASTA_RELAT = os.path.join(SCRIPT_DIR, "relatorios")
PASTA_EMAIL = os.path.join(SCRIPT_DIR, "email")
PASTA_EQUIP = os.path.join(PASTA_RELAT, "equip")
HIST        = os.path.join(PASTA_RELAT, "historico.json")
OUT_HTML    = os.path.join(PASTA_EMAIL, "corpo_email.html")
OUT_RESUMO  = os.path.join(PASTA_EMAIL, "resumo.json")

PREF = "C.I"
COL_DESTINO = "ÚLTIMO MOVIMIENTO EQUIPAMENTO (DESTINO)"
CAP = 40
BASE_BR = "Depósito Trânsito a Base BR"
DEPS_DANIF = ["Depósito Equipamentos Danificados BBC JOAO MONLEVADE (SOTREQ)",
              "Depósito Equip. Danificados BR", "Depósito Equip. Perdidos BR"]
ESPECIAIS = set([BASE_BR] + DEPS_DANIF)
RETORNO_SITRACK = "Depósito Retorno Sitrack"   # excluído da contagem total
AZUL = "#1F4E78"; OLIVA = "#7f7a52"; VERMELHO = "#963634"; DOURADO = "#b8860b"
CAT_LABEL = {"STD": "STD", "CAM": "Câmera", "O": "Outros"}


def categoria_std(m): m = (m or "").strip(); return m.startswith("STD") and "RSTMini" in m and m != "STD145_RSTMini"
def categoria_cameras(m): return (m or "").strip().startswith("Streamax")
def cat_de(m): return "STD" if categoria_std(m) else ("CAM" if categoria_cameras(m) else "O")


def variacao(atual, anterior):
    anterior = anterior or 0
    d = atual - anterior
    pct = (d / anterior * 100) if anterior else 0
    if d > 0:   return "#c0392b", "&#9650;", d, pct
    if d < 0:   return "#1e8449", "&#9660;", d, pct
    return "#7f8c8d", "&#9644;", d, pct


def cor_net(net):
    if net > 0:  return "#c0392b", "&#9650;"
    if net < 0:  return "#1e8449", "&#9660;"
    return "#7f8c8d", "&#9644;"


def std_cam_dict(v):
    if isinstance(v, dict):
        return {k: v.get(k, "" if k != "total" else 0) for k in ("total", "A", "R", "T", "B", "O")}
    return {"total": v or 0, "A": "", "R": "", "T": "", "B": "", "O": ""}


def carregar_equip_prev(hoje):
    if not os.path.isdir(PASTA_EQUIP):
        return None, None
    arqs = []
    for p in glob.glob(os.path.join(PASTA_EQUIP, "*.json")):
        m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(p))
        if m and m.group(1) < hoje:
            arqs.append((m.group(1), p))
    if not arqs:
        return None, None
    data_prev, caminho = max(arqs)
    try:
        return data_prev, json.load(open(caminho, encoding="utf-8"))
    except Exception:
        return None, None


def filtra(equip, pred):
    return {k: v for k, v in (equip or {}).items() if pred(v)}


def movimentos(eh, ep, cat):
    h = {k: v for k, v in eh.items() if v.get("cat") == cat}
    p = {k: v for k, v in (ep or {}).items() if v.get("cat") == cat}
    sh, sp = set(h), set(p)
    ent = sorted([(k, h[k]["m"], h[k]["d"]) for k in sh - sp], key=lambda x: x[2])
    sai = sorted([(k, p[k]["m"], p[k]["d"]) for k in sp - sh], key=lambda x: x[2])
    mov = sorted([(k, h[k]["m"], p[k]["d"], h[k]["d"]) for k in sh & sp if p[k]["d"] != h[k]["d"]], key=lambda x: x[2])
    return ent, sai, mov


def diff_equip(hoje, prev):
    sh, sp = set(hoje), set(prev or {})
    ent = sorted([(k, hoje[k]) for k in sh - sp], key=lambda x: (x[1].get("d", ""), x[1].get("m", "")))
    sai = sorted([(k, (prev or {})[k]) for k in sp - sh], key=lambda x: (x[1].get("d", ""), x[1].get("m", "")))
    return ent, sai


def tabela_historico(hist):
    cols = ["Total", "Ativação", "Retorno", "Tr. p. CI", "Tr. p.Base", "Outros"]
    cab2 = "".join(f'<th style="padding:6px 7px;border:1px solid #ddd;font-size:11px">{c}</th>' for c in cols)
    linhas = ""
    for h in hist[-14:]:
        s = std_cam_dict(h.get("std")); c = std_cam_dict(h.get("cameras"))
        try:
            dia = dt.datetime.strptime(h["data"], "%Y-%m-%d").strftime("%d/%m")
        except Exception:
            dia = h.get("data", "")
        cels = "".join(f'<td style="padding:6px 7px;border:1px solid #eee;text-align:right">{v}</td>'
                       for v in [s["total"], s["A"], s["R"], s["T"], s["B"], s["O"],
                                 c["total"], c["A"], c["R"], c["T"], c["B"], c["O"]])
        linhas += f'<tr><td style="padding:6px 7px;border:1px solid #eee;font-weight:600">{dia}</td>{cels}</tr>'
    return f"""
    <div style="font-size:15px;font-weight:700;margin:22px 0 8px">Histórico por tipo de depósito</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:11px">
      <tr><th style="border:1px solid #ddd;background:#fff"></th>
        <th colspan="6" style="padding:6px;border:1px solid {OLIVA};background:{OLIVA};color:#fff">STD</th>
        <th colspan="6" style="padding:6px;border:1px solid {VERMELHO};background:{VERMELHO};color:#fff">CAMERAS</th></tr>
      <tr style="background:#eef3fa"><th style="padding:6px 7px;border:1px solid #ddd;font-size:11px">Data</th>{cab2}{cab2}</tr>
      {linhas}
    </table>
    <div style="font-size:11px;color:#888;margin-top:4px">Tr. p. CI = trânsito (C.I.) &middot; Tr. p.Base = Depósito Trânsito a Base BR (tipo B)</div>"""


def quadro_reconciliacao(prev_data, cats):
    linhas = ""; tot_e = tot_s = 0
    for rot, e, s in cats:
        tot_e += e; tot_s += s; net = e - s
        cor, seta = cor_net(net)
        linhas += (f'<tr><td style="padding:6px 10px;border:1px solid #eee">{rot}</td>'
                   f'<td style="padding:6px 10px;border:1px solid #eee;text-align:right">{e}</td>'
                   f'<td style="padding:6px 10px;border:1px solid #eee;text-align:right">{s}</td>'
                   f'<td style="padding:6px 10px;border:1px solid #eee;text-align:right;color:{cor};font-weight:700">{seta} {net:+d}</td></tr>')
    net_t = tot_e - tot_s; cor, seta = cor_net(net_t)
    lt = (f'<tr style="background:#eef3fa;font-weight:700"><td style="padding:6px 10px;border:1px solid #ddd">TOTAL</td>'
          f'<td style="padding:6px 10px;border:1px solid #ddd;text-align:right">{tot_e}</td>'
          f'<td style="padding:6px 10px;border:1px solid #ddd;text-align:right">{tot_s}</td>'
          f'<td style="padding:6px 10px;border:1px solid #ddd;text-align:right;color:{cor}">{seta} {net_t:+d}</td></tr>')
    return f"""
    <div style="font-size:13px;color:#5b6b7b;margin:6px 0 6px">Movimentações vs {prev_data} &middot; <b>líquido = entraram − saíram</b> (o TOTAL reflete a variação do dia)</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:12px;margin-bottom:6px">
      <tr style="background:{AZUL};color:#fff"><td style="padding:6px 10px;border:1px solid {AZUL}">Categoria</td>
        <td style="padding:6px 10px;border:1px solid {AZUL};text-align:right">Entraram</td>
        <td style="padding:6px 10px;border:1px solid {AZUL};text-align:right">Saíram</td>
        <td style="padding:6px 10px;border:1px solid {AZUL};text-align:right">Líquido</td></tr>
      {linhas}{lt}
    </table>"""


def bloco_movimento(titulo, cor, ent, sai, mov):
    def ls(itens):
        return "".join(f'<tr><td style="padding:4px 8px;border:1px solid #eee">{i}</td>'
                       f'<td style="padding:4px 8px;border:1px solid #eee">{m}</td>'
                       f'<td style="padding:4px 8px;border:1px solid #eee">{d}</td></tr>' for i, m, d in itens[:CAP])
    def lm(itens):
        return "".join(f'<tr><td style="padding:4px 8px;border:1px solid #eee">{i}</td>'
                       f'<td style="padding:4px 8px;border:1px solid #eee">{m}</td>'
                       f'<td style="padding:4px 8px;border:1px solid #eee">{de} &rarr; {pa}</td></tr>' for i, m, de, pa in itens[:CAP])
    def mais(n): return f'<div style="font-size:11px;color:#888;margin:2px 0 10px">… e mais {n}</div>' if n > CAP else ''
    sec = f'<div style="font-size:15px;font-weight:700;margin:20px 0 6px;color:{cor}">{titulo}</div>'
    sec += (f'<div style="font-size:13px;margin-bottom:8px">Entraram: <b>{len(ent)}</b> &middot; '
            f'Saíram: <b>{len(sai)}</b> &middot; Movimentaram entre depósitos: <b>{len(mov)}</b></div>')
    if ent:
        sec += ('<div style="font-size:12px;font-weight:600;margin:8px 0 2px">Entraram (novos)</div>'
                '<table role="presentation" width="100%" style="border-collapse:collapse;font-size:11px">'
                '<tr style="background:#eef3fa"><td style="padding:4px 8px;border:1px solid #ddd">Equip.</td>'
                '<td style="padding:4px 8px;border:1px solid #ddd">Modelo</td><td style="padding:4px 8px;border:1px solid #ddd">Depósito</td></tr>'
                f'{ls(ent)}</table>{mais(len(ent))}')
    if sai:
        sec += ('<div style="font-size:12px;font-weight:600;margin:8px 0 2px">Saíram (não aparecem mais)</div>'
                '<table role="presentation" width="100%" style="border-collapse:collapse;font-size:11px">'
                '<tr style="background:#eef3fa"><td style="padding:4px 8px;border:1px solid #ddd">Equip.</td>'
                '<td style="padding:4px 8px;border:1px solid #ddd">Modelo</td><td style="padding:4px 8px;border:1px solid #ddd">Estava em</td></tr>'
                f'{ls(sai)}</table>{mais(len(sai))}')
    if mov:
        sec += ('<div style="font-size:12px;font-weight:600;margin:8px 0 2px">Movimentaram entre depósitos</div>'
                '<table role="presentation" width="100%" style="border-collapse:collapse;font-size:11px">'
                '<tr style="background:#eef3fa"><td style="padding:4px 8px;border:1px solid #ddd">Equip.</td>'
                '<td style="padding:4px 8px;border:1px solid #ddd">Modelo</td><td style="padding:4px 8px;border:1px solid #ddd">De &rarr; Para</td></tr>'
                f'{lm(mov)}</table>{mais(len(mov))}')
    return sec


def _tab_cat(itens, col3):
    linhas = "".join(
        f'<tr><td style="padding:4px 8px;border:1px solid #eee">{i}</td>'
        f'<td style="padding:4px 8px;border:1px solid #eee">{v.get("m","")}</td>'
        f'<td style="padding:4px 8px;border:1px solid #eee;font-weight:600">{CAT_LABEL.get(v.get("cat"),"?")}</td>'
        f'<td style="padding:4px 8px;border:1px solid #eee">{v.get("d","")}</td></tr>' for i, v in itens[:CAP])
    return (f'<table role="presentation" width="100%" style="border-collapse:collapse;font-size:11px">'
            f'<tr style="background:#eef3fa"><td style="padding:4px 8px;border:1px solid #ddd">Equip.</td>'
            f'<td style="padding:4px 8px;border:1px solid #ddd">Modelo</td><td style="padding:4px 8px;border:1px solid #ddd">Categoria</td>'
            f'<td style="padding:4px 8px;border:1px solid #ddd">{col3}</td></tr>{linhas}</table>'
            + (f'<div style="font-size:11px;color:#888;margin:2px 0 8px">… e mais {len(itens)-CAP}</div>' if len(itens) > CAP else ''))


def bloco_base_br(prev_data, bb_hoje, bb_prev):
    total = len(bb_hoje)
    ent, sai = diff_equip(bb_hoje, bb_prev)
    net = len(ent) - len(sai); cor, seta = cor_net(net)
    cnt_h = Counter(v["cat"] for v in bb_hoje.values())
    cnt_p = Counter(v.get("cat") for v in (bb_prev or {}).values())
    def lc(lbl, cat):
        nh, npv = cnt_h.get(cat, 0), cnt_p.get(cat, 0); c, s = cor_net(nh - npv)
        return (f'<td style="padding:8px 12px;text-align:center"><div style="font-size:20px;font-weight:700">{nh}</div>'
                f'<div style="font-size:11px;color:#5b6b7b">{lbl} &middot; <span style="color:{c}">{s} {nh-npv:+d}</span></div></td>')
    cab = (f'<div style="margin:26px 0 6px"><div style="display:inline-block;background:{DOURADO};color:#fff;'
           f'font-size:16px;font-weight:700;padding:8px 14px;border-radius:8px 8px 0 0">★ Depósito Trânsito a Base BR (TIPO B)</div></div>')
    box = (f'<div style="border:2px solid {DOURADO};border-radius:0 8px 8px 8px;padding:14px 16px;background:#fffdf5">'
           f'<div style="font-size:11px;color:#9a8a4a;margin-bottom:6px">Já incluído nos totais e nas movimentações acima — destaque do depósito.</div>'
           f'<div style="font-size:14px;margin-bottom:8px">Total no depósito: <b style="font-size:20px">{total}</b> '
           f'&middot; variação vs {prev_data}: <span style="color:{cor};font-weight:700;font-size:16px">{seta} {net:+d}</span></div>'
           f'<table role="presentation" width="100%" style="margin-bottom:6px"><tr>{lc("STD","STD")}{lc("Câmera","CAM")}{lc("Outros","O")}</tr></table>'
           f'<div style="font-size:13px;margin-top:6px">Movimentações: Entraram <b>{len(ent)}</b> &middot; Saíram <b>{len(sai)}</b></div>')
    if ent:
        box += '<div style="font-size:12px;font-weight:700;margin:10px 0 2px;color:#1e8449">Entraram</div>' + _tab_cat(ent, "Depósito")
    if sai:
        box += '<div style="font-size:12px;font-weight:700;margin:10px 0 2px;color:#c0392b">Saíram</div>' + _tab_cat(sai, "Estava em")
    if not bb_prev:
        box += '<div style="font-size:12px;color:#7f8c8d;margin-top:6px">Sem dia anterior para comparar.</div>'
    return cab + box + '</div>'


def bloco_3deps(prev_data, dep_hoje, dep_prev):
    ent, sai = diff_equip(dep_hoje, dep_prev)
    cab = (f'<div style="margin:26px 0 6px"><div style="display:inline-block;background:#6b2d2d;color:#fff;'
           f'font-size:16px;font-weight:700;padding:8px 14px;border-radius:8px 8px 0 0">⚠ Depósitos de danificados / perdidos</div></div>')
    box = (f'<div style="border:2px solid #6b2d2d;border-radius:0 8px 8px 8px;padding:14px 16px;background:#fcf6f6">'
           f'<div style="font-size:12px;color:#5b6b7b;margin-bottom:8px">Análise <b>independente</b> do total acima. '
           f'Depósitos: BBC JOAO MONLEVADE (SOTREQ), Equip. Danificados BR, Equip. Perdidos BR.</div>'
           f'<div style="font-size:14px;margin-bottom:8px">Entraram recentemente: <b style="color:#1e8449;font-size:18px">{len(ent)}</b> '
           f'&middot; Saíram: <b style="color:#c0392b;font-size:18px">{len(sai)}</b></div>')
    if ent:
        box += '<div style="font-size:13px;font-weight:700;margin:10px 0 2px;color:#1e8449">Entraram recentemente (STD / Câmera / Outros)</div>' + _tab_cat(ent, "Depósito")
    if sai:
        box += '<div style="font-size:13px;font-weight:700;margin:12px 0 2px;color:#c0392b">Saíram (não aparecem mais)</div>' + _tab_cat(sai, "Estava em")
    if not dep_prev:
        box += '<div style="font-size:12px;color:#7f8c8d;margin-top:6px">Sem dia anterior para comparar.</div>'
    if not ent and not sai and dep_prev:
        box += '<div style="font-size:12px;color:#7f8c8d;margin-top:6px">Sem entradas/saídas em relação ao dia anterior.</div>'
    return cab + box + '</div>'


def main():
    csv = sys.argv[1] if len(sys.argv) > 1 else max(glob.glob(os.path.join(PASTA_RELAT, "*.csv")), key=os.path.getmtime)
    os.makedirs(PASTA_EMAIL, exist_ok=True)
    os.makedirs(PASTA_EQUIP, exist_ok=True)

    df = pd.read_csv(csv, sep=";", encoding="latin-1", dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    df["NOME"] = df[COL_DESTINO].fillna("").str.split(" - ", n=1, expand=True)[1].fillna("").str.strip()
    df["ci"] = df["CLIENTE (NOME)"].astype(str).str.strip().str.upper().str.startswith(PREF.upper())

    m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(csv))
    hoje = m.group(1) if m else dt.date.today().isoformat()

    sub = df[df["ci"] | df["NOME"].isin(ESPECIAIS)]
    equip_hoje = {}
    for _, r in sub.iterrows():
        eid = str(r["EQUIPAMENTO"]).strip()
        if not eid:
            continue
        mod = r["MODELO"]
        equip_hoje[eid] = {"m": mod, "d": (r["NOME"] or "Na"), "cat": cat_de(mod), "ci": bool(r["ci"])}
    json.dump(equip_hoje, open(os.path.join(PASTA_EQUIP, f"{hoje}.json"), "w", encoding="utf-8"), ensure_ascii=False)

    data_prev, equip_prev = carregar_equip_prev(hoje)

    pred_main = lambda v: (v.get("ci", True) or v.get("d") == BASE_BR) and v.get("d") != RETORNO_SITRACK
    main_hoje = filtra(equip_hoje, pred_main)
    main_prev = filtra(equip_prev, pred_main) if equip_prev else None
    bb_hoje = filtra(equip_hoje, lambda v: v.get("d") == BASE_BR)
    bb_prev = filtra(equip_prev, lambda v: v.get("d") == BASE_BR) if equip_prev else None
    dep_hoje = filtra(equip_hoje, lambda v: v.get("d") in DEPS_DANIF)
    dep_prev = filtra(equip_prev, lambda v: v.get("d") in DEPS_DANIF) if equip_prev else None

    hist = []
    if os.path.exists(HIST):
        try:
            hist = json.load(open(HIST, encoding="utf-8"))
        except Exception:
            hist = []
    cur = next((h for h in hist if h.get("data") == hoje), None)
    prev = next((h for h in reversed(hist) if h.get("data", "") < hoje), None)
    total = (cur or {}).get("total", len(main_hoje))
    std = std_cam_dict((cur or {}).get("std"))["total"]
    cam = std_cam_dict((cur or {}).get("cameras"))["total"]

    if prev:
        pt = prev.get("total"); ps = std_cam_dict(prev.get("std"))["total"]; pc = std_cam_dict(prev.get("cameras"))["total"]
        cor, seta, d, pct = variacao(total, pt)
        comp = (f'<span style="color:{cor};font-weight:700">{seta} {abs(d)} ({pct:+.2f}%)</span> em relação a {prev.get("data")}')
        cs, ss, ds, _ = variacao(std, ps); cc, sc, dc, _ = variacao(cam, pc)
        badge = f'<span style="color:{cor}">{seta} {d:+d}</span>'
        badge_std = f'<span style="color:{cs}">{ss} {ds:+d}</span>'
        badge_cam = f'<span style="color:{cc}">{sc} {dc:+d}</span>'
    else:
        comp = '<span style="color:#7f8c8d">Primeiro relatório — sem dia anterior.</span>'
        badge = badge_std = badge_cam = '<span style="color:#7f8c8d">—</span>'

    if main_prev:
        e_s, s_s, m_s = movimentos(main_hoje, main_prev, "STD")
        e_c, s_c, m_c = movimentos(main_hoje, main_prev, "CAM")
        e_o, s_o, m_o = movimentos(main_hoje, main_prev, "O")
        recon = quadro_reconciliacao(data_prev, [("STD", len(e_s), len(s_s)),
                                                 ("CAMERAS", len(e_c), len(s_c)),
                                                 ("Outros (demais modelos)", len(e_o), len(s_o))])
        mov_html = recon + bloco_movimento("STD", OLIVA, e_s, s_s, m_s) + bloco_movimento("CAMERAS", VERMELHO, e_c, s_c, m_c)
    else:
        mov_html = '<div style="font-size:13px;color:#7f8c8d;margin-top:10px">Sem dia anterior para comparar movimentações.</div>'

    base_br_html = bloco_base_br(data_prev or "—", bb_hoje, bb_prev)
    deps_html = bloco_3deps(data_prev or "—", dep_hoje, dep_prev)

    try:
        data_fmt = dt.datetime.strptime(hoje, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        data_fmt = hoje

    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="margin:0;background:#f4f6f8;font-family:Segoe UI,Arial,sans-serif;color:#2c3e50">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:24px 0"><tr><td align="center">
<table role="presentation" width="720" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.06)">
  <tr><td style="background:{AZUL};padding:26px 30px">
    <div style="color:#fff;font-size:22px;font-weight:700">Relatório Diário &middot; Equipamentos em Depósito</div>
    <div style="color:#cdd9e8;font-size:14px;margin-top:4px">Sillion &middot; {data_fmt} &middot; clientes C.I. + Base BR</div>
  </td></tr>
  <tr><td style="padding:26px 30px">
    <div style="font-size:15px;margin-bottom:18px">Total em depósito: <b style="font-size:18px">{total}</b> &nbsp; {comp}</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:6px"><tr>
      <td width="33%" style="padding:6px"><div style="background:#eef3fa;border-radius:10px;padding:14px;text-align:center"><div style="font-size:24px;font-weight:700;color:{AZUL}">{total}</div><div style="font-size:12px;color:#5b6b7b">Total &middot; {badge}</div></div></td>
      <td width="33%" style="padding:6px"><div style="background:#f3f1e7;border-radius:10px;padding:14px;text-align:center"><div style="font-size:24px;font-weight:700;color:{OLIVA}">{std}</div><div style="font-size:12px;color:#5b6b7b">STD &middot; {badge_std}</div></div></td>
      <td width="33%" style="padding:6px"><div style="background:#f7ecec;border-radius:10px;padding:14px;text-align:center"><div style="font-size:24px;font-weight:700;color:{VERMELHO}">{cam}</div><div style="font-size:12px;color:#5b6b7b">CAMERAS &middot; {badge_cam}</div></div></td>
    </tr></table>
    {tabela_historico(hist)}
    <div style="font-size:15px;font-weight:700;margin:24px 0 4px">Análise de movimentações (C.I. + Base BR)</div>
    {mov_html}
    {base_br_html}
    {deps_html}
  </td></tr>
  <tr><td style="padding:16px 30px;background:#fafbfc;color:#9aa7b4;font-size:12px;border-top:1px solid #eee">
    Relatório automático (Lista de Equipamentos do Sitrack). Planilha completa em anexo.
  </td></tr>
</table>
</td></tr></table></body></html>"""
    open(OUT_HTML, "w", encoding="utf-8").write(html)

    resumo = {"data": hoje, "assunto": f"Relatório Diário - Equipamentos em Depósito ({data_fmt})",
              "total": total, "std": std, "cameras": cam,
              "anterior": (prev or {}).get("total"), "anterior_data": (prev or {}).get("data"),
              "arquivo_xlsx": f"{hoje}_listadoequipos_despositos.xlsx"}
    json.dump(resumo, open(OUT_RESUMO, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("OK ->", OUT_HTML)
    print("total:", total, "std:", std, "cam:", cam, "| anterior:", (prev or {}).get("data"))
    print("universo:", len(main_hoje), "| base br:", len(bb_hoje), "| 3 deps:", len(dep_hoje))


if __name__ == "__main__":
    main()
