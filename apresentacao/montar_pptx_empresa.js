// Gera um PowerPoint separado com os slides "Quem somos" e "Ferramentas e achados" da Contraprova.
// Uso: node apresentacao/montar_pptx_empresa.js dados/resumo.json dados/mapa.json apresentacao/<saida>.pptx
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
pres.title = "Contraprova · Quem somos e ferramentas";
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
pres.addSection({ title: "Quem somos" });
// 1a. Quem somos
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Quem somos" });
  rotulo(s, "Quem somos");
  s.addText("Uma organização sem fins lucrativos que confere as compras públicas com dados abertos", { placeholder: "title" });
  s.addText("A Contraprova é uma associação civil do terceiro setor. Lê os dados que o governo já publica, aponta o que não fecha e devolve o resultado à sociedade e aos próprios órgãos.",
    { x: MX, y: 2.05, w: 10.5, h: 0.65, margin: 0, fontSize: 16, color: C.text2, valign: "top", isTextBox: true, objectName: "Apresentação da organização" });
  const pilares = [["Lê", "Dados abertos", "Trabalha só com informação pública: Portal da Transparência e pedidos pela Lei de Acesso à Informação."],
    ["Aponta", "Divergências", "Checagens automáticas mostram onde a nota não fecha. Sinalizar é pedir conferência, não acusar."],
    ["Devolve", "Resultado aberto", "Publica relatórios e código para qualquer pessoa repetir a análise, e envia os casos aos órgãos responsáveis."]];
  const cw = (CW - 0.6) / 3, ch = 1.85, y0 = 2.9;
  pilares.forEach((p, i) => {
    const x = MX + i * (cw + 0.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y: y0, w: cw, h: ch, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Pilar ${p[1]}` });
    s.addText(p[0].toUpperCase(), { x: x + 0.25, y: y0 + 0.2, w: cw - 0.5, h: 0.26, margin: 0, fontFace: MONO, fontSize: 10, bold: true, color: C.accent1, charSpacing: 1, isTextBox: true, objectName: `Pilar ${p[1]} verbo` });
    s.addText(p[1], { x: x + 0.25, y: y0 + 0.48, w: cw - 0.5, h: 0.38, margin: 0, fontSize: 18, bold: true, color: C.text1, isTextBox: true, objectName: `Pilar ${p[1]} título` });
    s.addText(p[2], { x: x + 0.25, y: y0 + 0.9, w: cw - 0.5, h: ch - 1.0, margin: 0, fontSize: 13, color: C.text2, valign: "top", isTextBox: true, objectName: `Pilar ${p[1]} texto` });
  });
  const princ = [["Sem fins lucrativos", "Todo recurso vai para a análise e a publicação. Não há distribuição de lucro."],
    ["Independência", "Não recebe recursos de fornecedores do governo nem de partidos."],
    ["Método aberto", "Regras, código e dados tratados ficam disponíveis no GitHub."]];
  const rows = princ.map((l, i) => [
    { text: l[0], options: { bold: true, fontSize: 14, color: HEX.ink, border: i === 0 ? [{ pt: 1.5, color: HEX.ink }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }] : BORDER } },
    { text: l[1], options: { fontSize: 14, color: HEX.ink2, border: i === 0 ? [{ pt: 1.5, color: HEX.ink }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }] : BORDER } }]);
  s.addTable(rows, { x: MX, y: 4.95, w: CW, colW: [2.9, CW - 2.9], fontFace: "Arial", margin: [0.07, 0.15, 0.07, 0], objectName: "Princípios" });
  legenda(s, "Organização fictícia, criada para este estudo acadêmico (pós-graduação em Gestão de TI).", MX, 6.75, 10, 0.25);
}

// 11. Ferramentas e achados
pres.addSection({ title: "Ferramentas e achados" });
{
  const s = pres.addSlide({ masterName: "CONTEUDO", sectionTitle: "Ferramentas e achados" });
  rotulo(s, "Resumo do trabalho");
  s.addText("As ferramentas que usamos e o que os dados mostraram", { placeholder: "title" });
  const W2 = (CW - 0.5) / 2, X2 = MX + W2 + 0.5;
  eyebrow(s, "Ferramentas", MX, 2.05, W2);
  eyebrow(s, "O que encontramos", X2, 2.05, W2);
  const borda = i => i === 0 ? [{ pt: 1.5, color: HEX.ink }, { type: "none" }, { pt: 0.5, color: HEX.rule }, { type: "none" }] : BORDER;
  const ferr = [
    ["Portal da Transparência", "Fonte dos dados: 10 arquivos mensais de NF-e (jan–out/2026), cerca de 3,9 GB."],
    ["Python 3.15", "pandas (leitura, limpeza, agregações), NumPy (cálculos vetorizados) e biblioteca padrão (glob, json, re, math)."],
    ["HTML · CSS · JavaScript", "Apresentação interativa com mapa do Brasil por UF e dropdowns."],
    ["Node.js", "pptxgenjs e JSZip para gerar este PowerPoint com mapa e gráficos editáveis."],
    ["Git · GitHub · GitHub Pages", "Versionamento pelo GitHub Desktop e publicação do site."],
    ["Playwright · LibreOffice", "Conferência visual das páginas e dos slides."],
    ["Claude Opus 5.5 (Anthropic)", "Apoio de IA no tratamento dos dados, no código e na montagem das apresentações."],
  ];
  s.addTable(ferr.map((l, i) => [{ text: l[0], options: { bold: true, fontSize: 12, color: HEX.ink, border: borda(i), valign: "top" } },
    { text: l[1], options: { fontSize: 11, color: HEX.ink2, border: borda(i), valign: "top" } }]),
    { x: MX, y: 2.4, w: W2, colW: [2.15, W2 - 2.15], fontFace: "Arial", margin: [0.06, 0.12, 0.06, 0], objectName: "Tabela de ferramentas" });
  const tops = D.r9_itens;
  const ach = [[brl(T.valor), `em ${n0(T.notas)} notas de ${n0(Math.round(T.fornecedores / 1000))} mil fornecedores para ${n0(T.orgaos)} órgãos federais.`],
    [pct((UF.SP.valor_emit + UF.RJ.valor_emit) / T.valor, 0), "do valor emitido por fornecedores de São Paulo e Rio de Janeiro."],
    [n0(T.notas_divergentes), `notas sinalizadas (${pct(T.notas_divergentes / T.notas)} da base), somando ${brl(T.valor_divergente)}.`],
    [n0(tops), `itens com preço 10 vezes acima da mediana, com excesso estimado de ${brl(D.r9_excesso)}.`],
    [n0(D.qualidade.linhas_dup_notas), "linhas de nota publicadas em duplicidade pelo próprio portal."]];
  s.addTable(ach.map((l, i) => [{ text: l[0], options: { bold: true, fontFace: MONO, fontSize: 14, color: HEX.petrol, border: borda(i), valign: "top" } },
    { text: l[1], options: { fontSize: 12, color: HEX.ink2, border: borda(i), valign: "top" } }]),
    { x: X2, y: 2.4, w: W2, colW: [1.85, W2 - 1.85], fontFace: "Arial", margin: [0.08, 0.12, 0.08, 0], objectName: "Tabela de achados" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await aplicarTema(OUT, THEME);
  console.log("ok", OUT);
})();
