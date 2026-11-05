#!/usr/bin/env python3
"""
Tratamento das NF-e destinadas a órgãos federais (Portal da Transparência)
e aplicação de regras de detecção de divergências.

Uso:
    python3 tratar_nfs.py <pasta_csvs> <pasta_saida>

Entrada: arquivos AAAAMM_NFe_NotaFiscal.csv e AAAAMM_NFe_NotaFiscalItem.csv
         (separador ';', encoding latin-1, valores no formato 1.234,56)
Saída:   resumo.json (agregados para a apresentação) + CSVs de notas sinalizadas
"""
import glob
import json
import os
import sys

import numpy as np
import pandas as pd

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)

CUF = {11: "RO", 12: "AC", 13: "AM", 14: "RR", 15: "PA", 16: "AP", 17: "TO", 21: "MA", 22: "PI",
       23: "CE", 24: "RN", 25: "PB", 26: "PE", 27: "AL", 28: "SE", 29: "BA", 31: "MG", 32: "ES",
       33: "RJ", 35: "SP", 41: "PR", 42: "SC", 43: "RS", 50: "MS", 51: "MT", 52: "GO", 53: "DF"}
UFS = sorted(CUF.values())

H_COLS = {
    "CHAVE DE ACESSO": "chave", "SÉRIE": "serie", "NÚMERO": "numero",
    "NATUREZA DA OPERAÇÃO": "natureza", "DATA EMISSÃO": "dt_emissao",
    "EVENTO MAIS RECENTE": "evento", "DATA/HORA EVENTO MAIS RECENTE": "dt_evento",
    "CPF/CNPJ Emitente": "doc_emit", "RAZÃO SOCIAL EMITENTE": "emitente",
    "UF EMITENTE": "uf_emit", "MUNICÍPIO EMITENTE": "mun_emit",
    "ÓRGÃO SUPERIOR DESTINATÁRIO": "orgao_sup", "ÓRGÃO DESTINATÁRIO": "orgao",
    "CNPJ DESTINATÁRIO": "cnpj_dest", "NOME DESTINATÁRIO": "destinatario",
    "UF DESTINATÁRIO": "uf_dest", "DESTINO DA OPERAÇÃO": "destino",
    "PRESENÇA DO COMPRADOR": "presenca",
    "VALOR NOTA FISCAL": "valor",
}
I_COLS = {
    "CHAVE DE ACESSO": "chave", "NÚMERO PRODUTO": "n_item",
    "DESCRIÇÃO DO PRODUTO/SERVIÇO": "produto", "CÓDIGO NCM/SH": "ncm",
    "NCM/SH (TIPO DE PRODUTO)": "ncm_desc", "CFOP": "cfop", "QUANTIDADE": "qtd",
    "UNIDADE": "unid", "VALOR UNITÁRIO": "vunit", "VALOR TOTAL": "vtotal",
}


def num_br(s: pd.Series) -> pd.Series:
    """'1.234,56' -> 1234.56"""
    return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                         errors="coerce")


def dv_chave(ch: pd.Series) -> np.ndarray:
    """Dígito verificador (módulo 11) dos 43 primeiros dígitos da chave de acesso."""
    ok = ch.str.fullmatch(r"\d{44}").fillna(False).to_numpy()
    arr = np.zeros((len(ch), 43), dtype=np.int64)
    if ok.any():
        b = np.frombuffer("".join(ch[ok].str[:43]).encode(), dtype=np.uint8).reshape(-1, 43) - 48
        arr[ok] = b
    pesos = np.array([(i % 8) + 2 for i in range(43)][::-1])  # 2..9 da direita p/ esquerda
    r = (arr * pesos).sum(axis=1) % 11
    dv = 11 - r
    dv[dv >= 10] = 0
    return np.where(ok, dv, -1)


# ---------------------------------------------------------------- carga
notas, itens_agg, itens_red, dup_itens = [], [], [], []
for f in sorted(glob.glob(os.path.join(SRC, "*_NFe_NotaFiscal.csv"))):
    comp = os.path.basename(f)[:6]
    print("lendo", comp, flush=True)
    h = pd.read_csv(f, sep=";", encoding="latin1", dtype=str, usecols=list(H_COLS)).rename(columns=H_COLS)
    h["competencia"] = comp
    h["valor"] = num_br(h["valor"])
    notas.append(h)

    fi = f.replace("_NotaFiscal.csv", "_NotaFiscalItem.csv")
    it = pd.read_csv(fi, sep=";", encoding="latin1", dtype=str, usecols=list(I_COLS)).rename(columns=I_COLS)
    n0 = len(it)
    it = it.drop_duplicates(["chave", "n_item"])  # o portal publica linhas repetidas
    dup_itens.append(n0 - len(it))
    for c in ("qtd", "vunit", "vtotal"):
        it[c] = num_br(it[c])
    it["c5"] = it["cfop"].str[:1] == "5"
    it["c6"] = it["cfop"].str[:1] == "6"
    # R2 – quantidade x valor unitário diferente do valor total do item
    calc = it["qtd"] * it["vunit"]
    d = (calc - it["vtotal"]).abs()
    it["div_item"] = (d > 0.05) & (d > 0.01 * it["vtotal"].abs().clip(lower=1))
    it["dif_item"] = np.where(it["div_item"], calc - it["vtotal"], 0.0)
    g = it.groupby("chave", sort=False).agg(
        soma_itens=("vtotal", "sum"), n_itens=("vtotal", "size"),
        itens_div=("div_item", "sum"), dif_itens=("dif_item", "sum"),
        cfop5=("c5", "max"), cfop6=("c6", "max"),
        cfop_min=("cfop", "min"))
    itens_agg.append(g.reset_index())
    red = it[["chave", "n_item", "produto", "ncm", "ncm_desc", "unid", "qtd", "vunit", "vtotal"]].copy()
    red["competencia"] = comp
    itens_red.append(red)
    del it, h

nf = pd.concat(notas, ignore_index=True)
ag = pd.concat(itens_agg, ignore_index=True)
it = pd.concat(itens_red, ignore_index=True)
del notas, itens_agg, itens_red
for c in ("produto", "ncm", "ncm_desc", "unid", "competencia"):
    it[c] = it[c].astype("category")
print("notas", len(nf), "itens", len(it), flush=True)

# duplicidade de chave entre arquivos mensais
linhas_publicadas = len(nf)
linhas_dup_notas = int(nf.duplicated().sum())
ag = ag.groupby("chave", as_index=False).agg(
    soma_itens=("soma_itens", "sum"), n_itens=("n_itens", "sum"), itens_div=("itens_div", "sum"),
    dif_itens=("dif_itens", "sum"), cfop5=("cfop5", "max"), cfop6=("cfop6", "max"), cfop_min=("cfop_min", "min"))
nf = nf.drop_duplicates("chave").merge(ag, on="chave", how="left")
nf["n_itens"] = nf["n_itens"].fillna(0).astype(int)
nf["soma_itens"] = nf["soma_itens"].fillna(0.0)

nf["dt_emissao"] = pd.to_datetime(nf["dt_emissao"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
nf["dt_evento"] = pd.to_datetime(nf["dt_evento"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
nf["mes"] = nf["dt_emissao"].dt.strftime("%Y-%m")

# ---------------------------------------------------------------- regras
flags = {}

# R1 – valor da nota ≠ soma dos itens (frete, IPI, ST e desconto explicam diferenças pequenas)
nf["dif_cab_itens"] = nf["valor"] - nf["soma_itens"]
pct = nf["dif_cab_itens"].abs() / nf[["valor", "soma_itens"]].max(axis=1).clip(lower=1)
nf["pct_dif"] = pct
# razão nota/itens entre 0,75 e 1,35 é explicável por desoneração de ICMS (Conv. ICMS 87/02, vendas
# de medicamentos ao poder público), descontos, IPI, ST e frete; fora dessa faixa a diferença é sinalizada
razao = nf["valor"] / nf["soma_itens"].where(nf["soma_itens"] > 0)
dif_rel = (nf["n_itens"] > 0) & (nf["dif_cab_itens"].abs() > 100) & (pct > 0.10)
explicavel = razao.between(0.75, 1.35)
r1_explicadas = int((dif_rel & explicavel).sum())
r1_explicadas_valor = float(nf.loc[dif_rel & explicavel, "valor"].sum())
flags["R1_cabecalho_x_itens"] = dif_rel & ~explicavel

# R2 – itens com qtd × unitário ≠ total
flags["R2_qtd_x_unitario"] = nf["itens_div"].fillna(0) > 0

# R3 – chave de acesso inconsistente com os campos da nota
ch = nf["chave"].fillna("")
cuf = pd.to_numeric(ch.str[:2], errors="coerce").map(CUF)
aamm = ch.str[2:6]
cnpj_ch = ch.str[6:20]
serie_ch = pd.to_numeric(ch.str[22:25], errors="coerce")
num_ch = pd.to_numeric(ch.str[25:34], errors="coerce")
dv = pd.to_numeric(ch.str[43:44], errors="coerce")
doc = nf["doc_emit"].fillna("").str.replace(r"\D", "", regex=True).str.zfill(14)
r3 = pd.DataFrame({
    "uf": cuf != nf["uf_emit"],
    "aamm": aamm != nf["dt_emissao"].dt.strftime("%y%m"),
    # séries 890-899 = NF-e avulsa/emitida no portal do Fisco: a chave traz o CNPJ do órgão emissor
    # CPF de emitente vem mascarado no portal: só dá para conferir quando é CNPJ
    "cnpj": (cnpj_ch != doc) & ~serie_ch.between(890, 899) & nf["doc_emit"].str.fullmatch(r"\d{14}").fillna(False),
    "serie": serie_ch != pd.to_numeric(nf["serie"], errors="coerce"),
    "numero": num_ch != pd.to_numeric(nf["numero"], errors="coerce"),
    "dv": dv != dv_chave(ch),
})
r3 = r3.fillna(False).astype(bool)
mot = pd.Series("", index=nf.index)
for c in r3.columns:
    mot = mot.where(~r3[c], mot + c + ",")
nf["r3_motivo"] = mot.str.rstrip(",")
flags["R3_chave_inconsistente"] = r3.any(axis=1)

# R4 – mesma série/número do mesmo emitente com chaves diferentes, ou chave repetida
k = nf["doc_emit"].astype(str) + "|" + nf["serie"].astype(str) + "|" + nf["numero"].astype(str)
flags["R4_duplicidade"] = k.duplicated(keep=False)

# R5 – CFOP incompatível com o destino declarado / UFs
interna = nf["destino"].str.startswith("1")
inter = nf["destino"].str.startswith("2")
mesma_uf = nf["uf_emit"] == nf["uf_dest"]
flags["R5_cfop_destino"] = (interna & (nf["cfop6"] == True)) | (inter & (nf["cfop5"] == True)) \
    | (interna & ~mesma_uf & nf["uf_dest"].isin(UFS) & ~nf["presenca"].fillna("").str.startswith("1")) \
    | (inter & mesma_uf)

# R6 – nota zerada ou sem itens publicados
flags["R6_zerada_sem_itens"] = (nf["valor"].fillna(0) <= 0) | (nf["n_itens"] == 0)

# R7 – sem evento de autorização registrado
flags["R7_sem_autorizacao"] = ~nf["evento"].fillna("").str.contains("Autoriza", case=False)

# R8 – evento de autorização registrado antes da emissão
flags["R8_evento_antes_emissao"] = (nf["dt_evento"] < nf["dt_emissao"] - pd.Timedelta(hours=1))

F = pd.DataFrame(flags).fillna(False).astype(bool)
nf = pd.concat([nf, F], axis=1)
nf["n_regras"] = F.sum(axis=1)
nf["divergente"] = nf["n_regras"] > 0

# ---------------------------------------------------------------- R9 preço unitário atípico
it["unid_n"] = it["unid"].astype(str).str.upper().str.strip().str.replace(r"[^A-Z]", "", regex=True).str[:2]
it = it[(it["vunit"] > 0) & (it["qtd"] > 0)]
# grupo de comparação: NCM + unidade + 1ª palavra da descrição (separa p.ex. medicamentos de alto custo)
it["palavra"] = it["produto"].astype(str).str.upper().str.extract(r"([A-ZÀ-Ú]{3,})", expand=False).fillna("")
grp = it.groupby(["ncm", "unid_n", "palavra"], observed=True)["vunit"]
it["mediana_grp"] = grp.transform("median")
it["n_grp"] = grp.transform("size")
it["razao"] = it["vunit"] / it["mediana_grp"]
it["excesso"] = (it["vunit"] - it["mediana_grp"]) * it["qtd"]
out = it[(it["n_grp"] >= 30) & (it["razao"] >= 10) & (it["excesso"] >= 10000)].copy()
out = out.merge(nf[["chave", "emitente", "uf_emit", "orgao", "orgao_sup", "uf_dest", "dt_emissao"]], on="chave", how="left")
out = out.sort_values("excesso", ascending=False)
out.to_csv(os.path.join(OUT, "R9_precos_atipicos.csv"), index=False, sep=";", encoding="utf-8-sig", decimal=",")
r9_chaves = set(out["chave"])
nf["R9_preco_atipico"] = nf["chave"].isin(r9_chaves)
nf["n_regras"] += nf["R9_preco_atipico"].astype(int)
nf["divergente"] = nf["n_regras"] > 0

# ---------------------------------------------------------------- Benford (1º dígito do valor)
v = nf.loc[nf["valor"] >= 10, "valor"]
d1 = v.astype(str).str.lstrip("0.").str[0].astype(int)
obs = d1.value_counts(normalize=True).reindex(range(1, 10)).fillna(0)
esp = pd.Series({d: np.log10(1 + 1 / d) for d in range(1, 10)})
mad = float((obs - esp).abs().mean())

# ---------------------------------------------------------------- saídas
REGRAS = {
    "R1_cabecalho_x_itens": "Valor da nota fora da faixa 0,75–1,35× a soma dos itens",
    "R2_qtd_x_unitario": "Item com quantidade × valor unitário ≠ valor total",
    "R3_chave_inconsistente": "Chave de acesso não confere com UF, CNPJ, série, número, data ou DV",
    "R4_duplicidade": "Mesma série/número do emitente em chaves diferentes",
    "R5_cfop_destino": "CFOP ou UFs incompatíveis com o destino declarado (exceto venda presencial)",
    "R6_zerada_sem_itens": "Nota com valor zero ou sem itens publicados",
    "R7_sem_autorizacao": "Sem evento de Autorização de Uso registrado",
    "R8_evento_antes_emissao": "Evento registrado antes da data de emissão",
    "R9_preco_atipico": "Item ≥10× a mediana de itens similares e excesso ≥ R$10 mil",
}
cols_regra = list(REGRAS)
total_valor = float(nf["valor"].sum())


def resumo_regra(c):
    s = nf[nf[c]]
    return {"id": c, "descricao": REGRAS[c], "notas": int(len(s)),
            "valor": float(s["valor"].sum()), "pct_notas": len(s) / len(nf)}


por_uf = []
for uf in UFS:
    e = nf[nf["uf_emit"] == uf]
    d = nf[nf["uf_dest"] == uf]
    por_uf.append({
        "uf": uf, "notas_emit": int(len(e)), "valor_emit": float(e["valor"].sum()),
        "notas_dest": int(len(d)), "valor_dest": float(d["valor"].sum()),
        "divergentes": int(e["divergente"].sum()),
        "valor_divergente": float(e.loc[e["divergente"], "valor"].sum()),
        "taxa_div": float(e["divergente"].mean()) if len(e) else 0.0,
        "fornecedores": int(e["doc_emit"].nunique()),
        "interestadual_pct": float((e["uf_dest"] != uf).mean()) if len(e) else 0.0,
    })

mensal = nf.groupby("competencia").agg(notas=("chave", "size"), valor=("valor", "sum"),
                                       divergentes=("divergente", "sum")).reset_index()
orgaos = nf.groupby("orgao_sup").agg(notas=("chave", "size"), valor=("valor", "sum"),
                                     divergentes=("divergente", "sum")).sort_values("valor", ascending=False).head(10).reset_index()

# fluxo interestadual (origem -> destino) top 12
fluxo = nf[nf["uf_emit"] != nf["uf_dest"]].groupby(["uf_emit", "uf_dest"])["valor"].sum().sort_values(ascending=False).head(12).reset_index()


def mascara(doc):
    doc = str(doc)
    return "***.%s.%s-**" % (doc[3:6], doc[6:9]) if len(doc) == 11 else doc


# maiores divergências R1 (valor absoluto da diferença)
top_r1 = nf[nf["R1_cabecalho_x_itens"]].assign(abs_dif=lambda x: x["dif_cab_itens"].abs()) \
    .sort_values("abs_dif", ascending=False).drop_duplicates(["doc_emit", "valor", "soma_itens"]).head(8)
top_r1_list = [{"emitente": r.emitente, "uf": r.uf_emit, "orgao": r.orgao, "numero": r.numero,
                "valor": r.valor, "soma_itens": r.soma_itens, "dif": r.dif_cab_itens,
                "data": r.dt_emissao.strftime("%d/%m/%Y") if pd.notna(r.dt_emissao) else ""}
               for r in top_r1.itertuples()]

out["produto"] = out["produto"].astype(str)
top_r9 = out.groupby(["produto", "emitente", "vunit"], as_index=False).agg(
    ncm_desc=("ncm_desc", "first"), unid=("unid", "first"), qtd=("qtd", "sum"), mediana=("mediana_grp", "first"),
    razao=("razao", "first"), excesso=("excesso", "sum"), uf=("uf_emit", "first"), orgao=("orgao", "first"),
    notas=("chave", "nunique")).sort_values("excesso", ascending=False).head(8)
top_r9_list = [{"produto": r.produto[:70], "ncm_desc": str(r.ncm_desc)[:60], "unid": str(r.unid),
                "qtd": r.qtd, "vunit": r.vunit, "mediana": r.mediana, "razao": r.razao, "notas": int(r.notas),
                "excesso": r.excesso, "emitente": r.emitente, "uf": r.uf, "orgao": r.orgao}
               for r in top_r9.itertuples()]

# ranking de fornecedores por volume de notas divergentes (exceto R6/R7, que são de publicação)
sub = nf[nf[["R1_cabecalho_x_itens", "R2_qtd_x_unitario", "R3_chave_inconsistente", "R4_duplicidade",
             "R5_cfop_destino", "R8_evento_antes_emissao", "R9_preco_atipico"]].any(axis=1)]
forn = sub.groupby(["doc_emit", "emitente", "uf_emit"]).agg(notas=("chave", "size"), valor=("valor", "sum"),
                                                            regras=("n_regras", "max")).reset_index()
tot_forn = nf.groupby("doc_emit")["chave"].size().rename("notas_total")
forn = forn.merge(tot_forn, on="doc_emit").sort_values("valor", ascending=False).head(10)
forn_list = [{"emitente": r.emitente, "doc": mascara(r.doc_emit), "uf": r.uf_emit, "notas": int(r.notas),
              "notas_total": int(r.notas_total), "valor": float(r.valor)} for r in forn.itertuples()]

r3_motivos = nf.loc[nf["R3_chave_inconsistente"], "r3_motivo"].str.split(",").explode().value_counts().to_dict()

resumo = {
    "periodo": {"inicio": nf["competencia"].min(), "fim": nf["competencia"].max()},
    "totais": {"notas": int(len(nf)), "itens": int(nf["n_itens"].sum()), "valor": total_valor,
               "fornecedores": int(nf["doc_emit"].nunique()), "orgaos": int(nf["orgao"].nunique()),
               "notas_divergentes": int(nf["divergente"].sum()),
               "valor_divergente": float(nf.loc[nf["divergente"], "valor"].sum()),
               "ticket_mediano": float(nf["valor"].median()),
               "multi_regra": int((nf["n_regras"] >= 2).sum())},
    "regras": [resumo_regra(c) for c in cols_regra],
    "r3_motivos": r3_motivos,
    "por_uf": por_uf,
    "mensal": mensal.to_dict("records"),
    "orgaos": orgaos.to_dict("records"),
    "fluxo": fluxo.to_dict("records"),
    "top_r1": top_r1_list,
    "top_r9": top_r9_list,
    "fornecedores": forn_list,
    "benford": {"obs": obs.round(5).tolist(), "esp": esp.round(5).tolist(), "mad": mad, "n": int(len(v))},
    "r1_explicadas": {"notas": r1_explicadas, "valor": r1_explicadas_valor},
    "r1_hist": [{"faixa": str(k), "notas": int(v)} for k, v in
                pd.cut(razao[dif_rel], [0, .25, .5, .75, 1.0, 1.35, 2, 5, np.inf], right=False)
                .value_counts().sort_index().items()],
    "qualidade": {"linhas_publicadas": linhas_publicadas, "linhas_dup_notas": linhas_dup_notas,
                  "linhas_dup_itens": int(sum(dup_itens)),
                  "avulsas_890_899": int(serie_ch.between(890, 899).sum())},
    "r9_itens": int(len(out)), "r9_excesso": float(out["excesso"].sum()),
}
with open(os.path.join(OUT, "resumo.json"), "w", encoding="utf-8") as fh:
    json.dump(resumo, fh, ensure_ascii=False, indent=1, default=float)

# base de notas sinalizadas para auditoria
cols = ["competencia", "chave", "serie", "numero", "dt_emissao", "doc_emit", "emitente", "uf_emit", "mun_emit",
        "orgao_sup", "orgao", "cnpj_dest", "uf_dest", "destino", "cfop_min", "valor", "soma_itens",
        "dif_cab_itens", "n_itens", "itens_div", "r3_motivo", "n_regras"] + cols_regra
nf.loc[nf["divergente"], cols].sort_values(["n_regras", "valor"], ascending=False) \
    .to_csv(os.path.join(OUT, "notas_divergentes.csv"), index=False, sep=";", encoding="utf-8-sig", decimal=",")
pd.DataFrame(por_uf).to_csv(os.path.join(OUT, "producao_por_uf.csv"), index=False, sep=";",
                            encoding="utf-8-sig", decimal=",")
print(json.dumps(resumo["totais"], indent=1, ensure_ascii=False))
print(pd.DataFrame(resumo["regras"]).to_string())
print("R3 motivos", r3_motivos)
print("Benford MAD", mad)
