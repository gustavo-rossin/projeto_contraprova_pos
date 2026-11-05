// Gera a versão PowerPoint da apresentação Contraprova a partir de resumo.json e mapa.json.
// Uso: node apresentacao/montar_pptx.js dados/resumo.json dados/mapa.json apresentacao/contraprova_nfe.pptx
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { aplicarTema } = require("./aplicar_tema.js");

const [, , RES, MAP, OUT] = process.argv;
const D = JSON.parse(fs.readFileSync(RES, "utf8"));
const MAPA = JSON.parse(fs.readFileSync(MAP, "utf8"));

const HEX = {
  ink: "13212A", ink2: "44545F", muted: "6E7D88", rule: "C9D3D9", paper: "EDF1F3", white: "FFFFFF",
  petrol: "0B6A6F", mark: "F5C518", alert: "B23F2B", soft: "D6E9E8",
  q: ["DCEDEC", "A9D3D1", "6FB4B3", "3B9294", "16707A", "0A4A55"],
};
const THEME = {
  name: "Contraprova", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.ink2, lt2: HEX.paper, accent1: HEX.petrol, accent2: HEX.alert,
    accent3: HEX.mark, accent4: HEX.q[2], accent5: HEX.q[3], accent6: HEX.muted, hlink: HEX.petrol, folHlink: HEX.ink2 },
};
const MONO = "Courier New";

// ---------- formatação pt-BR ----------
const fmt = (v, d) => v.toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
function brl(v) {
  const a = Math.abs(v);
  if (a >= 1e9) return "R$ " + fmt(v / 1e9, 2) + " bi";
  if (a >= 1e6) return "R$ " + fmt(v / 1e6, 1) + " mi";
  if (a >= 1e3) return "R$ " + fmt(v / 1e3, 1) + " mil";
  return "R$ " + fmt(v, 2);
}
const n0 = v => fmt(v, 0);
const pct = (v, d = 1) => fmt(v * 100, d) + "%";
const NOMES = { AC: "Acre", AL: "Alagoas", AP: "Amapá", AM: "Amazonas", BA: "Bahia", CE: "Ceará", DF: "Distrito Federal", ES: "Espírito Santo", GO: "Goiás", MA: "Maranhão", MT: "Mato Grosso", MS: "Mato Grosso do Sul", MG: "Minas Gerais", PA: "Pará", PB: "Paraíba", PR: "Paraná", PE: "Pernambuco", PI: "Piauí", RJ: "Rio de Janeiro", RN: "Rio Grande do Norte", RS: "Rio Grande do Sul", RO: "Rondônia", RR: "Roraima", SC: "Santa Catarina", SP: "São Paulo", SE: "Sergipe", TO: "Tocantins" };
const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
const T = D.totais;
const UF = Object.fromEntries(D.por_uf.map(u => [u.uf, u]));
const periodo = `${MESES[+D.periodo.inicio.slice(4) - 1]}–${MESES[+D.periodo.fim.slice(4) - 1]}/${D.periodo.fim.slice(0, 4)}`;

// ---------- apresentação ----------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Contraprova · NF-e destinadas à União";
pres.author = "Contraprova (organização fictícia)";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333, MX = 0.6, CW = W - 2 * MX;

pres.defineSlideMaster({
  title: "CAPA", background: { color: HEX.white },
  objects: [
    { rect: { x: MX, y: 0.42, w: 0.13, h: 0.13, fill: { color: C.accent3 } } },
    { text: { text: "CONTRAPROVA", options: { x: MX + 0.2, y: 0.33, w: 3, h: 0.3, margin: 0, fontFace: MONO, fontSize: 11, bold: true, color: C.text1, charSpacing: 2, isTextBox: true } } },
  ],
});
pres.defineSlideMaster({
  title: "CONTEUDO", background: { color: HEX.white },
  margin: [0.5, 0.6, 0.5, 0.6],
  objects: [
    { rect: { x: MX, y: 0.42, w: 0.13, h: 0.13, fill: { color: C.accent3 } } },
    { text: { text: "CONTRAPROVA", options: { x: MX + 0.2, y: 0.33, w: 3, h: 0.3, margin: 0, fontFace: MONO, fontSize: 11, bold: true, color: C.text1, charSpacing: 2, isTextBox: true } } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.8, w: CW, h: 1.25, fontSize: 32, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: W - MX - 0.6, y: 7.0, w: 0.6, h: 0.3, fontFace: MONO, fontSize: 10, color: C.accent6, align: "right" },
});

function rotulo(slide, txt) { // rótulo de seção no canto superior direito
  slide.addText(txt.toUpperCase(), { x: W - MX - 6, y: 0.33, w: 6, h: 0.3, margin: 0, align: "right", fontFace: MONO, fontSize: 10, color: C.accent6, charSpacing: 1, isTextBox: true, objectName: "Rótulo da seção" });
}
function eyebrow(slide, txt, x, y, w) {
  slide.addText(txt.toUpperCase(), { x, y, w, h: 0.3, margin: 0, fontFace: MONO, fontSize: 10, bold: true, color: C.accent1, charSpacing: 1, isTextBox: true, objectName: "Sobretítulo" });
}
function nota(slide, runs, x, y, w, h, name) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, objectName: name || "Nota" });
  slide.addText(runs, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h - 0.24, margin: 0, fontSize: 14, color: C.text2, valign: "top", isTextBox: true, objectName: (name || "Nota") + " texto" });
}
function legenda(slide, txt, x, y, w, h = 0.5) {
  slide.addText(txt, { x, y, w, h, margin: 0, fontSize: 11, color: C.accent6, valign: "top", isTextBox: true, objectName: "Legenda" });
}

// ---------- mapa nativo ----------
function parsePath(d) {
  const rings = [];
  for (const seg of d.split("Z").filter(Boolean)) {
    const pts = seg.replace(/^M/, "").split("L").map(p => p.split(",").map(Number));
    rings.push(pts);
  }
  return rings;
}
const RINGS = Object.fromEntries(Object.entries(MAPA.ufs).map(([uf, g]) => [uf, parsePath(g.d)]));
const LBL = { GO: [-14, 12], DF: [2, -2], ES: [18, 0], RJ: [10, 0], SE: [16, 6], AL: [20, 4], PB: [26, 0], RN: [22, 0] };
const PEQ = ["SE", "AL", "PB", "RN", "PE", "DF", "ES", "RJ", "SC"];

function mapa(slide, metric, breaks, labels, x0, y0, H) {
  const s = H / MAPA.h;
  const cls = v => { let i = 0; while (i < breaks.length && v >= breaks[i]) i++; return i; };
  // DF por último, por cima de GO
  const ordem = Object.keys(RINGS).filter(u => u !== "DF").concat(["DF"]);
  for (const uf of ordem) {
    const rings = RINGS[uf];
    const xs = rings.flat().map(p => p[0]), ys = rings.flat().map(p => p[1]);
    const minx = Math.min(...xs), miny = Math.min(...ys), maxx = Math.max(...xs), maxy = Math.max(...ys);
    const points = [];
    for (const r of rings) {
      r.forEach((p, i) => points.push({ x: (p[0] - minx) * s, y: (p[1] - miny) * s, ...(i === 0 ? { moveTo: true } : {}) }));
      points.push({ close: true });
    }
    const c = cls(UF[uf][metric]);
    slide.addShape(pres.shapes.CUSTOM_GEOMETRY, {
      x: x0 + minx * s, y: y0 + miny * s, w: Math.max(0.02, (maxx - minx) * s), h: Math.max(0.02, (maxy - miny) * s),
      points, fill: { color: HEX.q[c] }, line: { color: HEX.white, width: 0.75 }, objectName: `UF ${uf} – ${NOMES[uf]}`,
    });
  }
  for (const uf of Object.keys(RINGS)) {
    const g = MAPA.ufs[uf]; const [dx, dy] = LBL[uf] || [0, 0];
    const c = cls(UF[uf][metric]);
    const cx = x0 + (g.cx + dx) * s, cy = y0 + (g.cy + dy) * s;
    slide.addText(uf, { x: cx - 0.25, y: cy - 0.1, w: 0.5, h: 0.2, margin: 0, align: "center", valign: "middle", fontFace: MONO, bold: true,
      fontSize: PEQ.includes(uf) ? 7 : 9, color: c >= 3 ? HEX.white : HEX.ink, isTextBox: true, objectName: `Sigla ${uf}` });
  }
  // legenda
  let lx = x0, ly = y0 + H + 0.1;
  labels.forEach((l, i) => {
    if (i === 3) { lx = x0; ly += 0.28; }
    slide.addShape(pres.shapes.RECTANGLE, { x: lx, y: ly + 0.04, w: 0.28, h: 0.15, fill: { color: HEX.q[i] }, line: { type: "none" }, objectName: `Legenda cor ${i + 1}` });
    slide.addText(l, { x: lx + 0.34, y: ly, w: 1.3, h: 0.24, margin: 0, fontFace: MONO, fontSize: 10, color: C.text2, isTextBox: true, objectName: `Legenda texto ${i + 1}` });
    lx += 1.7;
  });
}

// ---------- gráficos de barras horizontais ----------
function barras(slide, labels, values, opt) {
  slide.addChart(pres.charts.BAR, [{ name: opt.name, labels, values }], {
    x: opt.x, y: opt.y, w: opt.w, h: opt.h, barDir: "bar", barGapWidthPct: 45,
    chartColors: [opt.color || HEX.petrol], catAxisOrientation: "maxMin",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: opt.fmt, dataLabelFontSize: 11,
    dataLabelColor: HEX.ink2, dataLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, catAxisLabelColor: HEX.ink, catAxisLabelFontFace: "+mn-lt",
    catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    valAxisMaxVal: opt.max, valAxisMinVal: 0, showLegend: false, showTitle: false, objectName: opt.name,
  });
}

// ---------- tabelas ----------
const TH = t => ({ text: t, options: { fontFace: MONO, fontSize: 9, bold: true, color: HEX.muted, border: [{ type: "none" }, { type: "none" }, { pt: 1.5, color: HEX.ink }, { type: "none" }] } });
const THR = t => ({ text: t, options: { ...TH(t).options, align: "right" } });
const BORDER = [{ type: "none" }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }];
const TD = (t, o = {}) => ({ text: t, options: { fontSize: 12, color: HEX.ink, border: BORDER, valign: "top", ...o } });
const TDN = (t, o = {}) => TD(t, { fontFace: MONO, fontSize: 11, align: "right", ...o });
const TD2 = (main, sub) => ({ text: [{ text: main, options: { fontSize: 12, color: HEX.ink, breakLine: true } }, { text: sub, options: { fontSize: 10, color: HEX.muted } }], options: { border: BORDER, valign: "top" } });

const UNIDADES = {
  unidade: ["UNIDAD", "UNIDADE", "UNID", "UND", "UN", "UNI", "UM", "EA", "U"], peça: ["PEÇA", "PECA", "PEÇ", "PEC", "PC", "PÇ"],
  metro: ["METRO", "MT", "M"], kit: ["KIT", "KT"], caixa: ["CX", "CAIXA"], kg: ["KG"], frasco: ["FR", "FA", "FRASCO"],
  pacote: ["PCT", "PACOTE"], litro: ["L", "LT", "LITRO"], "m²": ["M2"], "m³": ["M3"], teste: ["TST", "TS", "TESTE"],
  envelope: ["ENV"], conjunto: ["CJ", "CONJ"], comprimido: ["CP", "COMP", "CPR"], exemplar: ["EX"], ampola: ["AMP"],
};
const PLURAL = { unidade: "unidades", peça: "peças", metro: "metros", kit: "kits", caixa: "caixas", frasco: "frascos", pacote: "pacotes", litro: "litros", teste: "testes", envelope: "envelopes", conjunto: "conjuntos", comprimido: "comprimidos", exemplar: "exemplares", ampola: "ampolas" };
function unidade(raw, qtd) {
  const u = String(raw || "").trim().toUpperCase();
  const nome = Object.keys(UNIDADES).find(k => UNIDADES[k].includes(u));
  if (!nome) return u.toLowerCase();
  return qtd === 1 ? nome : (PLURAL[nome] || nome);
}

// =================================================================== slides
// 1. Capa
{
  pres.addSection({ title: "Abertura" });
  const s = pres.addSlide({ masterName: "CAPA", sectionTitle: "Abertura" });
  s.addText("OBSERVATÓRIO CIDADÃO DE COMPRAS PÚBLICAS", { x: W - MX - 6, y: 0.33, w: 6, h: 0.3, margin: 0, align: "right", fontFace: MONO, fontSize: 10, color: C.accent6, charSpacing: 1, isTextBox: true, objectName: "Rótulo" });
  eyebrow(s, `Relatório de varredura · NF-e destinadas à União · ${periodo}`, MX, 1.35, 10);
  s.addText([{ text: "Onde as notas", options: { breakLine: true } }, { text: "não fecham", options: { highlight: HEX.mark } }],
    { x: MX, y: 1.7, w: 10, h: 2.1, margin: 0, fontSize: 64, lineSpacingMultiple: 0.9, bold: true, color: C.text1, valign: "top", isTextBox: true, objectName: "Título da capa" });
  s.addText(`Cruzamos ${n0(T.notas)} notas fiscais eletrônicas e ${n0(T.itens)} itens publicados no Portal da Transparência para achar o que merece uma segunda olhada.`,
    { x: MX, y: 3.95, w: 8.2, h: 0.9, margin: 0, fontSize: 18, color: C.text2, valign: "top", isTextBox: true, objectName: "Subtítulo" });
  // chave de acesso de exemplo
  const parts = [["25", "UF"], ["2603", "Ano/mês"], ["11222333000181", "CNPJ do emitente"], ["55", "Mod."], ["001", "Série"], ["000004812", "Número"], ["1", "Tipo"], ["70531964", "Código num."]];
  const base = parts.map(p => p[0]).join(""); let soma = 0;
  for (let i = base.length - 1, w = 2; i >= 0; i--, w = w === 9 ? 2 : w + 1) soma += +base[i] * w;
  let dv = 11 - (soma % 11); if (dv >= 10) dv = 0; parts.push([String(dv), "DV"]);
  let x = MX; const CH = 0.215, y = 5.05;
  parts.forEach((p, i) => {
    const w = Math.max(p[0].length * CH + 0.16, 0.62);
    if (i === 0) s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.48, fill: { color: HEX.mark, transparency: 35 }, line: { type: "none" }, objectName: "Destaque UF" });
    s.addText(p[0], { x, y, w, h: 0.48, margin: 0, align: "center", valign: "middle", fontFace: MONO, fontSize: 20, color: C.text1, isTextBox: true, objectName: `Chave ${p[1]}` });
    s.addShape(pres.shapes.LINE, { x, y: y + 0.5, w, h: 0, line: { color: i === 0 ? HEX.alert : HEX.ink, width: 1.5 }, objectName: `Traço ${p[1]}` });
    s.addText(p[1].toUpperCase(), { x, y: y + 0.56, w: w + 0.3, h: 0.22, margin: 0, fontFace: MONO, fontSize: 8, bold: i === 0, color: i === 0 ? HEX.alert : HEX.muted, isTextBox: true, objectName: `Rótulo ${p[1]}` });
    x += w + 0.08;
  });
  s.addText([{ text: "Exemplo de chave de acesso de 44 dígitos. Os dois primeiros dizem " }, { text: "25 (Paraíba)", options: { highlight: HEX.mark } },
    { text: `, mas o emitente está cadastrado no DF: esse é um dos ${n0(D.regras.find(r => r.id.startsWith("R3")).notas)} casos desse tipo encontrados na base.` }],
    { x: MX, y: 6.05, w: 11, h: 0.6, margin: 0, fontSize: 13, color: C.text2, valign: "top", isTextBox: true, objectName: "Nota da chave" });
}

// 1b. Premissas
{
  pres.addSection({ title: "Premissas" });
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Premissas" });
  rotulo(s, "Premissas");
  s.addText("O que esta base mostra", { placeholder: "title" });
  const semInfo = (D.orgaos.find(o => o.orgao_sup === "Sem informação") || { notas: 0 }).notas;
  const linhas = [
    ["Escopo", `Notas fiscais eletrônicas de mercadorias (modelo 55) emitidas por fornecedores para órgãos do governo federal, quase todos do Poder Executivo, ${periodo}.`],
    ["Fora da base", "Compras de estados e municípios, e serviços (limpeza, vigilância, consultoria, obras), cobrados por nota fiscal de serviço municipal."],
    ["Nota não é pagamento", "A nota mostra o que o fornecedor faturou. Não confirma empenho, entrega nem pagamento, que estão nos dados de despesa do portal."],
    ["Uma linha, uma nota", `No arquivo de notas, cada linha é uma nota fiscal: ${n0(T.notas)} depois de remover as duplicadas. Destas, ${n0(semInfo)} vêm com órgão "Sem informação".`],
    ["Item é linha de produto", `Os ${n0(T.itens)} itens são as linhas de produto dentro das notas (média de ${fmt(T.itens / T.notas, 1)} por nota), não a quantidade comprada. As unidades são misturadas, por isso as quantidades não são somadas.`],
    ["Valor × soma dos itens", "Valor da nota é o total do cabeçalho. Soma dos itens é a soma do valor total de cada linha de produto da mesma nota."],
  ];
  const rows = linhas.map((l, i) => [
    { text: l[0], options: { bold: true, fontSize: 15, color: i === linhas.length - 1 ? HEX.alert : HEX.ink, border: i === 0 ? [{ pt: 1.5, color: HEX.ink }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }] : BORDER, valign: "top" } },
    { text: l[1], options: { fontSize: 14, color: HEX.ink2, border: i === 0 ? [{ pt: 1.5, color: HEX.ink }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }] : BORDER, valign: "top" } },
  ]);
  s.addTable(rows, { x: MX, y: 2.05, w: CW, colW: [2.9, CW - 2.9], fontFace: "Arial", margin: [0.09, 0.15, 0.09, 0], objectName: "Tabela de premissas" });
}

// 2. Base
{
  pres.addSection({ title: "Panorama" });
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Panorama" });
  rotulo(s, "A base analisada");
  s.addText(`${brl(T.valor)} em compras federais, nota por nota`, { placeholder: "title" });
  s.addText(brl(T.valor), { x: MX, y: 2.15, w: 5.9, h: 0.9, margin: 0, fontSize: 54, bold: true, color: C.accent1, isTextBox: true, objectName: "KPI valor" });
  s.addText(`valor somado das NF-e emitidas para órgãos federais, ${periodo}`, { x: MX, y: 3.05, w: 5.9, h: 0.3, margin: 0, fontSize: 13, color: C.text2, isTextBox: true, objectName: "KPI valor rótulo" });
  const kpis = [[fmt(T.notas / 1e6, 2) + " mi", "notas únicas"], [fmt(T.itens / 1e6, 2) + " mi", "itens de produto"], [n0(Math.round(T.fornecedores / 1000)) + " mil", "fornecedores (CPF/CNPJ)"],
    [n0(T.orgaos), "órgãos destinatários"], [brl(T.ticket_mediano), "valor mediano por nota"], [String(D.mensal.length), "competências mensais"]];
  kpis.forEach((k, i) => {
    const x = MX + (i % 3) * 2.0, y = 3.45 + Math.floor(i / 3) * 1.05;
    s.addText(k[0], { x, y, w: 1.95, h: 0.6, margin: 0, fontSize: 28, bold: true, color: C.text1, isTextBox: true, objectName: `KPI ${k[1]}` });
    s.addText(k[1], { x, y: y + 0.62, w: 1.9, h: 0.5, margin: 0, fontSize: 12, color: C.text2, valign: "top", isTextBox: true, objectName: `KPI rótulo ${k[1]}` });
  });
  const X2 = 7.0, W2 = W - MX - X2;
  eyebrow(s, "Antes de procurar erro nos fornecedores", X2, 2.2, W2);
  s.addText("O próprio arquivo público tinha repetições. Limpamos isso primeiro, para não acusar ninguém por um erro de publicação.",
    { x: X2, y: 2.55, w: W2, h: 0.8, margin: 0, fontSize: 16, color: C.text2, valign: "top", isTextBox: true, objectName: "Lede" });
  nota(s, [{ text: `${n0(D.qualidade.linhas_dup_notas)} linhas de nota`, options: { bold: true, color: HEX.ink } }, { text: " e " },
    { text: `${n0(D.qualidade.linhas_dup_itens)} linhas de item`, options: { bold: true, color: HEX.ink } },
    { text: " aparecem duplicadas, idênticas, nos CSVs mensais. Foram removidas antes de qualquer cálculo." }], X2, 3.5, W2, 1.0, "Nota duplicadas");
  nota(s, "CPFs de emitentes chegam mascarados pelo portal (***.123.456-**) e continuam mascarados aqui.", X2, 4.65, W2, 0.75, "Nota CPF");
  nota(s, [{ text: "Nota sinalizada: ", options: { bold: true, color: HEX.ink } }, { text: "nota que caiu em pelo menos uma checagem de divergência (valor × itens, preço unitário, chave de acesso, CFOP, datas, duplicidade ou publicação incompleta). Precisa ser conferida, não é necessariamente irregular." }], MX, 5.4, 5.95, 1.2, "Definição nota sinalizada");
  legenda(s, `Fonte: Portal da Transparência, NF-e modelo 55, arquivos ${D.periodo.inicio} a ${D.periodo.fim}.`, MX, 6.8, 9, 0.25);
}

// 3 e 4. Mapas
const BR_VAL = [1e8, 3e8, 1e9, 3e9, 1e10];
const LB_VAL = ["< 100 mi", "100–300 mi", "300 mi–1 bi", "1–3 bi", "3–10 bi", "> 10 bi"];
function slideMapa(metric, rot, titulo, texto, eb) {
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Produção estadual" });
  rotulo(s, rot);
  s.addText(titulo, { placeholder: "title" });
  mapa(s, metric, BR_VAL, LB_VAL, MX, 2.1, 4.3);
  const rows = [...D.por_uf].sort((a, b) => b[metric] - a[metric]).slice(0, 12);
  const X2 = 6.2, W2 = W - MX - X2;
  eyebrow(s, eb, X2, 2.15, W2);
  s.addText(texto, { x: X2, y: 2.5, w: W2, h: 0.7, margin: 0, fontSize: 14, color: C.text2, valign: "top", isTextBox: true, objectName: "Texto do mapa" });
  barras(s, rows.map(u => u.uf), rows.map(u => u[metric] / 1e9), { name: "Valor por UF (R$ bi)", x: X2 - 0.1, y: 3.2, w: W2 + 0.1, h: 3.55, fmt: '"R$ "0.00" bi"', max: rows[0][metric] / 1e9 * 1.22 });
  return s;
}
pres.addSection({ title: "Produção estadual" });
{
  const tot = T.valor, r = [...D.por_uf].sort((a, b) => b.valor_emit - a.valor_emit);
  const top12 = r.slice(0, 12).reduce((a, u) => a + u.valor_emit, 0);
  slideMapa("valor_emit", "Produção estadual",
    `São Paulo e Rio de Janeiro emitem ${pct((UF.SP.valor_emit + UF.RJ.valor_emit) / tot, 0)} do valor vendido à União`,
    `Fornecedores sediados em São Paulo emitiram ${brl(UF.SP.valor_emit)}. Os 12 primeiros estados somam ${pct(top12 / tot, 0)} do total.`,
    "Valor emitido por UF do fornecedor");
  slideMapa("valor_dest", "Para onde a nota vai",
    `Pelo destino, o Espírito Santo recebe ${brl(UF.ES.valor_dest)} em notas, mas emite só ${brl(UF.ES.valor_emit)}`,
    "ES e DF sobem porque concentram órgãos compradores centrais. O destino mostra onde o órgão está registrado, não onde o produto foi consumido.",
    "Valor recebido por UF do órgão comprador");
}

// 5. Fluxos
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Produção estadual" });
  rotulo(s, "Quem vende para quem");
  s.addText("São Paulo abastece ES, DF e RJ nos três maiores fluxos interestaduais", { placeholder: "title" });
  const fl = D.fluxo.slice(0, 10);
  eyebrow(s, "Maiores fluxos interestaduais · origem → destino", MX, 2.15, 5.8);
  barras(s, fl.map(f => `${f.uf_emit} → ${f.uf_dest}`), fl.map(f => f.valor / 1e9), { name: "Fluxos interestaduais (R$ bi)", x: MX - 0.1, y: 2.5, w: 5.9, h: 4.2, fmt: '"R$ "0.00" bi"', max: fl[0].valor / 1e9 * 1.25 });
  const it = D.por_uf.filter(u => u.valor_emit >= 1e9).sort((a, b) => b.interestadual_pct - a.interestadual_pct);
  const X2 = 7.0, W2 = W - MX - X2;
  eyebrow(s, "Quanto da produção de cada UF sai do estado", X2, 2.15, W2);
  barras(s, it.map(u => u.uf), it.map(u => u.interestadual_pct), { name: "% de notas para outra UF", x: X2 - 0.1, y: 2.5, w: W2 + 0.1, h: 3.3, fmt: "0%", max: 1, color: HEX.q[3] });
  legenda(s, "Estados com ao menos R$ 1 bilhão emitido. Percentual do número de notas com UF do destinatário diferente da UF do emitente.", X2, 5.95, W2, 0.6);
}

// 6. Mensal
{
  pres.addSection({ title: "Evolução" });
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Evolução" });
  rotulo(s, "Evolução mensal");
  const ms = D.mensal, pico = ms.reduce((a, m) => m.valor > a.valor ? m : a);
  const mesNome = c => MESES[+c.slice(4) - 1] + "/" + c.slice(2, 4);
  const nomeLongo = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][+pico.competencia.slice(4) - 1];
  s.addText(`${nomeLongo}/${pico.competencia.slice(0, 4)} concentrou ${brl(pico.valor)}, o dobro de um mês típico`, { placeholder: "title" });
  s.addChart(pres.charts.BAR, [{ name: "Valor (R$ bi)", labels: ms.map(m => mesNome(m.competencia)), values: ms.map(m => m.valor / 1e9) }], {
    x: MX - 0.1, y: 2.15, w: CW + 0.2, h: 3.6, barDir: "col", barGapWidthPct: 60,
    chartColors: ms.map(m => m === pico ? HEX.petrol : HEX.q[2]), varyColors: false,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0', dataLabelFontSize: 11, dataLabelColor: HEX.ink2, dataLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, catAxisLabelColor: HEX.ink2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontSize: 10, valAxisLabelColor: HEX.muted, valAxisLabelFontFace: "+mn-lt",
    valAxisMinVal: 0, valAxisMaxVal: 14, valAxisMajorUnit: 4, valGridLine: { color: HEX.rule, size: 0.5 }, catGridLine: { style: "none" },
    showValAxisTitle: true, valAxisTitle: "R$ bilhões", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.muted,
    showLegend: false, objectName: "Valor mensal",
  });
  const rows = [[TH("MÊS")].concat(ms.map(m => ({ text: mesNome(m.competencia), options: { ...TH("").options, align: "center" } }))),
    [TD("% sinalizadas", { fontSize: 11, color: HEX.muted })].concat(ms.map(m => TDN(pct(m.divergentes / m.notas), { align: "center" })))];
  s.addTable(rows, { x: MX, y: 5.95, w: CW, colW: [1.5].concat(ms.map(() => (CW - 1.5) / ms.length)), fontFace: "Arial", rowH: 0.3, objectName: "Tabela % sinalizadas" });
  const taxas = ms.map(m => m.divergentes / m.notas);
  legenda(s, `Barras: valor total das notas emitidas no mês, em R$ bilhões. Linha "% sinalizadas": parte das notas do mês com pelo menos uma divergência, estável entre ${pct(Math.min(...taxas))} e ${pct(Math.max(...taxas))}. O pico de março é de valor, não de mais notas com problema.`, MX, 6.62, CW, 0.45);
}

// 7. Preços atípicos
{
  pres.addSection({ title: "Casos" });
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Casos" });
  rotulo(s, "Preço unitário fora da curva");
  const tripe = D.top_r9.find(r => r.produto.trim().toUpperCase().startsWith("TRIPE"));
  s.addText([{ text: "Um tripé de " }, { text: brl(tripe.vunit), options: { highlight: HEX.mark } }, { text: ` entre ${n0(D.r9_itens)} itens com preço dez vezes acima do comum` }], { placeholder: "title" });
  const rows = [[TH("ITEM · EMITENTE"), TH("ÓRGÃO"), THR("QUANTIDADE"), THR("PREÇO UNITÁRIO"), THR("MEDIANA SIMILAR"), THR("VEZES A MEDIANA"), THR("EXCESSO")]];
  D.top_r9.slice(0, 6).forEach(r => rows.push([
    TD2(r.produto, `${r.emitente} · ${r.uf}${r.notas > 1 ? ` · ${r.notas} notas` : ""}`), TD(r.orgao, { fontSize: 11 }),
    TDN(`${n0(r.qtd)} ${unidade(r.unid, r.qtd)}`), TDN("R$ " + fmt(r.vunit, 2)), TDN("R$ " + fmt(r.mediana, 2)),
    TDN(n0(r.razao) + "×", { color: HEX.alert, bold: true }), TDN(brl(r.excesso))]));
  s.addTable(rows, { x: MX, y: 2.1, w: CW, colW: [3.15, 1.85, 1.45, 1.85, 1.6, 1.05, 1.18], fontFace: "Arial", margin: [0.05, 0.08, 0.05, 0], objectName: "Tabela preços atípicos" });
  legenda(s, `Item comparado com a mediana de itens do mesmo NCM, mesma unidade e mesma palavra inicial da descrição (mínimo de 30 itens). Excesso somado dos ${n0(D.r9_itens)} itens: ${brl(D.r9_excesso)}. Parte provavelmente é erro de unidade ou de digitação.`, MX, 6.55, CW, 0.5);
}

// 8. Nota x itens
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Casos" });
  rotulo(s, "Valor da nota × soma dos itens");
  s.addText("Notas de R$ 150 milhões cujos itens somam mais que o dobro", { placeholder: "title" });
  const rows = [[TH("EMITENTE"), TH("ÓRGÃO"), TH("EMISSÃO"), THR("VALOR DA NOTA"), THR("SOMA DOS ITENS"), THR("RAZÃO")]];
  D.top_r1.slice(0, 6).forEach(r => rows.push([TD2(r.emitente, `${r.uf} · nota nº ${r.numero}`), TD(r.orgao.replace(" - Unidades com vínculo direto", ""), { fontSize: 11 }), TDN(r.data, { align: "left" }),
    TDN(brl(r.valor)), TDN(brl(r.soma_itens)), TDN(fmt(r.valor / r.soma_itens, 2), { color: HEX.alert, bold: true })]));
  s.addTable(rows, { x: MX, y: 2.1, w: CW, colW: [4.4, 2.3, 1.35, 1.35, 1.6, 1.13], fontFace: "Arial", margin: [0.05, 0.08, 0.05, 0], objectName: "Tabela nota x itens" });
  nota(s, "Os maiores casos são medicamentos de alto custo para o Ministério da Saúde. Os dois descontos legais explicam uma razão de até cerca de 0,64 (próximo slide). Razões de 0,13 a 0,42 pedem o contrato e o empenho para conferir.", MX, 5.75, 9.2, 0.95, "Nota medicamentos");
}

// 8b. ICMS e CAP
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Casos" });
  rotulo(s, "Descontos legais em medicamentos");
  s.addText("Desoneração de ICMS e desconto CAP explicam uma nota até cerca de 36% menor que os itens", { placeholder: "title" });
  const cw = (CW - 0.4) / 2, y = 2.15, h = 2.75;
  const cards = [
    ["Imposto estadual", "Desoneração de ICMS", "O Convênio ICMS 87/2002 isenta de ICMS os medicamentos de uma lista anexa vendidos a órgãos da administração pública direta e indireta, federal, estadual e municipal. O valor dispensado precisa ser deduzido do preço e demonstrado na nota (§ 6º): sai como desconto do tamanho da alíquota de ICMS, que varia por estado e operação."],
    ["Regra da CMED", "Desconto CAP", "O Coeficiente de Adequação de Preço é um desconto mínimo obrigatório sobre o Preço Fábrica na venda ao governo de medicamentos de uma lista da CMED: PMVG = Preço Fábrica × (1 − CAP). Foi de 21,53% até 23/09/2026 e passou a 19,79% a partir de 24/09/2026 (Resolução CTE-CMED nº 1/2026)."],
  ];
  cards.forEach((c, i) => {
    const x = MX + i * (cw + 0.4);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Cartão ${c[1]}` });
    s.addText(c[0].toUpperCase(), { x: x + 0.3, y: y + 0.22, w: cw - 0.6, h: 0.28, margin: 0, fontFace: MONO, fontSize: 10, bold: true, color: C.accent1, charSpacing: 1, isTextBox: true, objectName: `${c[1]} sobretítulo` });
    s.addText(c[1], { x: x + 0.3, y: y + 0.52, w: cw - 0.6, h: 0.42, margin: 0, fontSize: 19, bold: true, color: C.text1, isTextBox: true, objectName: `${c[1]} título` });
    s.addText(c[2], { x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: h - 1.15, margin: 0, fontSize: 14, color: C.text2, valign: "top", isTextBox: true, objectName: `${c[1]} texto` });
  });
  const passos = [["R$ 100,00", "item a Preço Fábrica"], ["R$ 78,47", "após CAP de 21,53%"], ["R$ 64,35", "após desonerar ICMS de 18%"], ["0,64", "razão nota ÷ itens"]];
  const pw = 2.35, gap = 0.55, py = 5.2;
  passos.forEach((p, i) => {
    const x = MX + i * (pw + gap);
    s.addShape(pres.shapes.LINE, { x, y: py, w: pw, h: 0, line: { color: HEX.ink, width: 1.5 }, objectName: `Conta traço ${i + 1}` });
    s.addText(p[0], { x, y: py + 0.08, w: pw, h: 0.45, margin: 0, fontFace: MONO, fontSize: 22, color: i === 3 ? C.accent1 : C.text1, bold: i === 3, isTextBox: true, objectName: `Conta valor ${i + 1}` });
    s.addText(p[1], { x, y: py + 0.55, w: pw, h: 0.3, margin: 0, fontSize: 12, color: C.accent6, isTextBox: true, objectName: `Conta rótulo ${i + 1}` });
    if (i < 3) s.addText(i === 2 ? "=" : "→", { x: x + pw + 0.05, y: py + 0.08, w: gap - 0.1, h: 0.45, margin: 0, align: "center", fontFace: MONO, fontSize: 20, color: C.accent6, isTextBox: true, objectName: `Conta seta ${i + 1}` });
  });
  legenda(s, "Exemplo ilustrativo com ICMS de 18%. Os descontos só aparecem como diferença entre nota e itens quando os itens vêm a preço cheio e o desconto entra no total. Os casos do slide anterior (0,13 a 0,42) ficam bem abaixo desse limite.", MX, 6.3, CW, 0.5);
}

// 9. Órgãos
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Casos" });
  rotulo(s, "Quem compra");
  s.addText("Saúde concentra metade do valor; Desenvolvimento Social, a maior taxa de sinalização", { placeholder: "title" });
  const curto = t => t.replace(/^Ministério (da|do|das|de) /, "").replace("Desenvolvimento e Assistência Social, Família e Combate à Fome", "Desenvolvimento Social")
    .replace("Integração e do Desenvolvimento Regional", "Integração Regional").replace("Meio Ambiente e Mudança do Clima", "Meio Ambiente")
    .replace("Justiça e Segurança Pública", "Justiça e Seg. Pública").replace("Agricultura e Pecuária", "Agricultura");
  const os = D.orgaos;
  eyebrow(s, "Valor das notas por órgão superior", MX, 2.15, 5.8);
  barras(s, os.map(o => curto(o.orgao_sup)), os.map(o => o.valor / 1e9), { name: "Valor por órgão (R$ bi)", x: MX - 0.1, y: 2.5, w: 5.9, h: 4.3, fmt: '"R$ "0.00" bi"', max: os[0].valor / 1e9 * 1.3 });
  const ot = [...os].sort((a, b) => b.divergentes / b.notas - a.divergentes / a.notas);
  const X2 = 7.0, W2 = W - MX - X2;
  eyebrow(s, "% de notas sinalizadas no órgão", X2, 2.15, W2);
  barras(s, ot.map(o => curto(o.orgao_sup)), ot.map(o => o.divergentes / o.notas), { name: "% sinalizadas por órgão", x: X2 - 0.1, y: 2.5, w: W2 + 0.1, h: 4.3, fmt: "0.0%", max: 0.24, color: HEX.alert });
}

// 11. (próximos passos removido)
pres.addSection({ title: "Encerramento" });

// 12. Encerramento
{
  const s = pres.addSlide({ masterName: "CAPA", sectionTitle: "Encerramento" });
  s.addText([{ text: "Toda nota pública merece uma " }, { text: "contraprova", options: { highlight: HEX.mark } }, { text: "." }],
    { x: MX, y: 2.9, w: 11, h: 1.9, margin: 0, fontSize: 54, bold: true, color: C.text1, valign: "bottom", isTextBox: true, objectName: "Frase final" });
  s.addText(`${n0(T.notas)} notas lidas. ${n0(T.notas_divergentes)} separadas para verificação. ${n0(D.r9_itens)} preços para explicar.`,
    { x: MX, y: 5.0, w: 11, h: 0.45, margin: 0, fontSize: 18, color: C.text2, isTextBox: true, objectName: "Resumo final" });
  legenda(s, `Contraprova é uma organização fictícia criada para este estudo. Dados públicos do Portal da Transparência (NF-e destinadas a órgãos do Poder Executivo Federal), ${periodo}. Sinalizações são indícios estatísticos e de consistência, sujeitos a verificação.`, MX, 6.1, 11.5, 0.6);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await aplicarTema(OUT, THEME);
  console.log("ok", OUT);
})();
