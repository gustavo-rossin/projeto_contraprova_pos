#!/usr/bin/env python3
"""Monta a apresentação HTML a partir do resumo.json (tratar_nfs.py) e do mapa.json (mapa_svg.py).

Uso: python3 montar_apresentacao.py <template.html> <resumo.json> <mapa.json> <saida.html>
"""
import json
import re
import sys

tpl, res_p, map_p, out_p = sys.argv[1:5]
R = json.load(open(res_p, encoding="utf-8"))
M = json.load(open(map_p, encoding="utf-8"))
html = open(tpl, encoding="utf-8").read()


def n0(v):
    return f"{v:,.0f}".replace(",", ".")


def dec(v, d):
    return f"{v:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl(v):
    a = abs(v)
    if a >= 1e9:
        return f"R$ {dec(v / 1e9, 2)} bi"
    if a >= 1e6:
        return f"R$ {dec(v / 1e6, 1)} mi"
    if a >= 1e3:
        return f"R$ {dec(v / 1e3, 1)} mil"
    return f"R$ {dec(v, 2)}"


def curto(v):
    if v >= 1e6:
        return dec(v / 1e6, 2) + " mi"
    if v >= 1e3:
        return dec(v / 1e3, 0 if v >= 1e5 else 1) + " mil"
    return n0(v)


def pct(v, d=1):
    return dec(v * 100, d) + "%"


MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro",
         "novembro", "dezembro"]
T = R["totais"]
uf = {u["uf"]: u for u in R["por_uf"]}
ini, fim = R["periodo"]["inicio"], R["periodo"]["fim"]
periodo = f"{MESES[int(ini[4:]) - 1][:3]}–{MESES[int(fim[4:]) - 1][:3]}/{fim[:4]}"
mensal = R["mensal"]
pico = max(mensal, key=lambda m: m["valor"])
taxas = [m["divergentes"] / m["notas"] for m in mensal]
regra = {r["id"]: r for r in R["regras"]}
top_abs = max(R["por_uf"], key=lambda u: u["divergentes"])
tripe = next((r for r in R["top_r9"] if r["produto"].strip().upper().startswith("TRIPE")), R["top_r9"][0])

V = {
    "periodo": periodo,
    "notas": n0(T["notas"]), "itens": n0(T["itens"]),
    "notas_curto": curto(T["notas"]), "itens_curto": curto(T["itens"]), "fornecedores_curto": curto(T["fornecedores"]),
    "orgaos": n0(T["orgaos"]), "ticket": brl(T["ticket_mediano"]), "n_meses": str(len(mensal)),
    "valor_total": brl(T["valor"]),
    "dup_notas": n0(R["qualidade"]["linhas_dup_notas"]), "dup_itens": n0(R["qualidade"]["linhas_dup_itens"]),
    "arquivos": f"{ini} a {fim}",
    "pct_sp_rj": pct((uf["SP"]["valor_emit"] + uf["RJ"]["valor_emit"]) / T["valor"], 0),
    "es_dest": brl(uf["ES"]["valor_dest"]), "es_emit": brl(uf["ES"]["valor_emit"]),
    "mes_pico": MESES[int(pico["competencia"][4:]) - 1].capitalize() + "/" + pico["competencia"][:4],
    "valor_pico": brl(pico["valor"]),
    "div_min": pct(min(taxas)), "div_max": pct(max(taxas)),
    "div_notas": n0(T["notas_divergentes"]), "div_pct": pct(T["notas_divergentes"] / T["notas"]),
    "div_valor": brl(T["valor_divergente"]), "multi": n0(T["multi_regra"]),
    "r1x_notas": n0(R["r1_explicadas"]["notas"]), "r1x_valor": brl(R["r1_explicadas"]["valor"]),
    "r1_notas": n0(regra["R1_cabecalho_x_itens"]["notas"]), "r1_valor": brl(regra["R1_cabecalho_x_itens"]["valor"]),
    "r3_notas": n0(regra["R3_chave_inconsistente"]["notas"]),
    "avulsas": n0(R["qualidade"]["avulsas_890_899"]),
    "se_notas": n0(uf["SE"]["notas_emit"]), "sp_notas": n0(uf["SP"]["notas_emit"]),
    "top_abs_uf": top_abs["uf"], "top_abs_n": n0(top_abs["divergentes"]),
    "sem_info": n0(next((o["notas"] for o in R["orgaos"] if o["orgao_sup"] == "Sem informação"), 0)),
    "itens_nota": dec(T["itens"] / T["notas"], 1),
    "tripe": brl(tripe["vunit"]), "r9_itens": n0(R["r9_itens"]), "r9_excesso": brl(R["r9_excesso"]),
    "pct_top_forn": pct(sum(f["valor"] for f in R["fornecedores"]) / T["valor_divergente"], 0),
    "mad": dec(R["benford"]["mad"], 4), "benford_n": n0(R["benford"]["n"]),
}

html = re.sub(r"\{\{(\w+)\}\}", lambda m: V[m.group(1)], html)
html = html.replace("/*__DATA__*/null", json.dumps(R, ensure_ascii=False, separators=(",", ":")))
html = html.replace("/*__MAPA__*/null", json.dumps(M, separators=(",", ":")))
open(out_p, "w", encoding="utf-8").write(html)
# versão autônoma (para abrir direto no navegador)
alone = ('<!doctype html>\n<html lang="pt-BR">\n<head>\n<meta charset="utf-8">\n'
         '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
         '<style>body{margin:0}</style>\n</head>\n<body>\n' + html + '\n</body>\n</html>\n')
open(out_p.replace(".html", "_standalone.html"), "w", encoding="utf-8").write(alone)
print("ok", len(html), "bytes")
