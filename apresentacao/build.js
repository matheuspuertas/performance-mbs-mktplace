// Performance Marketplaces - Setembro 2026, identidade Puertas Marketing
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("C:/Users/mathe/.claude/skills/synced/091d6316-ee3f-4f62-8652-364905227600_d392708a-8365-4750-a078-b2cd1d41f2c0/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Puertas Marketing",
  headFontFace: "Montserrat",
  bodyFontFace: "Manrope",
  colors: {
    dk1: "0A2B5C", lt1: "FFFFFF", dk2: "051730", lt2: "F4F3EF",
    accent1: "0F5BD9", accent2: "1E7CFF", accent3: "4C95FF", accent4: "B9C6DC",
    accent5: "4F6180", accent6: "E6EFFD", hlink: "0F5BD9", folHlink: "4F6180",
  },
};
const K = { navy: "0A2B5C", deep: "051730", blue: "0F5BD9", bright: "1E7CFF", sky: "4C95FF",
  mist: "B9C6DC", muted: "4F6180", bg: "F4F3EF", line: "E2E0D8", pale: "E6EFFD", white: "FFFFFF",
  card: "0E3470", alert: "C2410C" };
const HF = "Montserrat", BF = "Manrope";
const W = 13.333, H = 7.5, MX = 0.89, CW = W - 2 * MX;
const FOOT = "Puertas Marketing · MBS Pro Grooming · Setembro 2026";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Performance Marketplaces - Setembro 2026";
pres.author = "Puertas Marketing";
pres.company = "Puertas Marketing";
pres.theme = { headFontFace: HF, bodyFontFace: BF };

const footer = (color) => [
  { text: { text: FOOT, options: { x: MX, y: 6.92, w: 8, h: 0.3, fontFace: BF, fontSize: 10, color, margin: 0 } } },
];
pres.defineSlideMaster({ title: "CLARO", background: { color: K.bg }, objects: footer(K.muted),
  slideNumber: { x: W - MX - 0.6, y: 6.92, w: 0.6, h: 0.3, fontFace: BF, fontSize: 10, color: K.muted, align: "right", margin: 0 } });
pres.defineSlideMaster({ title: "ESCURO", background: { color: K.deep }, objects: footer(K.mist),
  slideNumber: { x: W - MX - 0.6, y: 6.92, w: 0.6, h: 0.3, fontFace: BF, fontSize: 10, color: K.mist, align: "right", margin: 0 } });
pres.defineSlideMaster({ title: "CAPA", background: { color: K.navy } });

// ---------- helpers ----------
const brl = (v, d = 0) => "R$ " + v.toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const num = (v, d = 0) => v.toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = (v, d = 1) => num(v * 100, d) + "%";

function header(s, eyebrow, title, dark) {
  s.addText(eyebrow.toUpperCase(), { x: MX, y: 0.62, w: CW, h: 0.28, fontFace: BF, fontSize: 12, bold: true,
    color: dark ? K.sky : K.blue, charSpacing: 1, margin: 0, isTextBox: true });
  s.addText(title, { x: MX, y: 0.95, w: CW, h: 0.75, fontFace: HF, fontSize: 32, bold: true,
    color: dark ? K.bg : K.navy, margin: 0, valign: "top", isTextBox: true });
}
function card(s, x, y, w, h, dark) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08,
    fill: { color: dark ? K.card : K.white }, line: { color: dark ? K.card : K.line, width: 0.75 } });
}
function stat(s, x, y, w, h, value, label, tag, dark, opts = {}) {
  card(s, x, y, w, h, dark);
  const p = 0.25;
  if (tag) s.addText(tag.toUpperCase(), { x: x + p, y: y + 0.25, w: w - 2 * p, h: 0.25, fontFace: BF, fontSize: 10, bold: true,
    color: opts.tagColor || (dark ? K.sky : K.blue), charSpacing: 1, margin: 0, isTextBox: true });
  s.addText(value, { x: x + p, y: y + (tag ? 0.55 : 0.3), w: w - 2 * p, h: 0.75, fontFace: HF, fontSize: opts.size || 36, bold: true,
    color: opts.valueColor || (dark ? K.sky : K.navy), margin: 0, valign: "middle", isTextBox: true });
  s.addText(label, { x: x + p, y: y + (tag ? 1.35 : 1.1), w: w - 2 * p, h: h - (tag ? 1.5 : 1.25), fontFace: BF, fontSize: 14,
    color: dark ? K.mist : K.muted, margin: 0, valign: "top", isTextBox: true });
}
function note(s, text, y, dark) {
  s.addText(text, { x: MX, y, w: CW, h: 0.4, fontFace: BF, fontSize: 14, color: dark ? K.mist : K.muted, margin: 0, isTextBox: true });
}
function brackets(s) { // motivo da capa das apresentacoes Puertas: cantos em L
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.14, w: 1.94, h: 0.14, fill: { color: K.bright }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 1.94, w: 0.14, h: 1.94, fill: { color: K.bright }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: W - 1.94, y: 0, w: 1.94, h: 0.14, fill: { color: K.blue }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: W - 0.14, y: 0, w: 0.14, h: 1.94, fill: { color: K.blue }, line: { type: "none" } });
}
const axisFont = { catAxisLabelFontFace: BF, valAxisLabelFontFace: BF, dataLabelFontFace: BF, legendFontFace: BF,
  catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 10, legendFontSize: 11 };

// ---------- dados (planilha Performance MBS MKTPLACE, set/2026) ----------
const ML_FAT = [["set/25", 430257], ["out/25", 536046], ["nov/25", 789016], ["dez/25", 657000], ["jan/26", 492000],
  ["fev/26", 484560], ["mar/26", 658863], ["abr/26", 748644], ["mai/26", 707533], ["jun/26", 657713],
  ["jul/26", 727470], ["ago/26", 737072], ["set/26", 768014]];
const ML_TICKET = [["jan/26", 174.05], ["fev/26", 168.89], ["mar/26", 171.71], ["abr/26", 175.94], ["mai/26", 166.05],
  ["jun/26", 160.81], ["jul/26", 157.19], ["ago/26", 164.97], ["set/26", 146.62]];
const SH_ACOS = [["jan/26", 10.03], ["fev/26", 12.03], ["mar/26", 11.95], ["abr/26", 7.92], ["mai/26", 7.64],
  ["jun/26", 9.27], ["jul/26", 6.58], ["ago/26", 6.07], ["set/26", 5.76]];
const SH_ROAS = [6.41, 6.96, 8.86, 12.63, 13.08, 10.79, 15.19, 14.14, 17.35];

// ---------- 1. capa ----------
let s = pres.addSlide({ masterName: "CAPA" });
brackets(s);
s.addImage({ path: "logo_branco.png", x: MX, y: 0.89, w: 2.92, h: 0.68, altText: "Puertas Marketing Digital" });
// logo do cliente em cartao branco, como na proposta Bertachini
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: W - MX - 2.7, y: 0.72, w: 2.7, h: 1.05, rectRadius: 0.1, fill: { color: K.white }, line: { type: "none" } });
s.addImage({ path: "logo_mbs.png", x: W - MX - 2.7 + 0.3, y: 0.72 + 0.135, w: 2.1, h: 0.823, altText: "MBS Pro Grooming" });
s.addText("RELATÓRIO MENSAL · MBS PRO GROOMING", { x: MX, y: 2.45, w: 9, h: 0.3, fontFace: BF, fontSize: 12, bold: true, color: K.sky, charSpacing: 1, margin: 0, isTextBox: true });
s.addText("Performance Marketplaces", { x: MX, y: 2.85, w: 11, h: 0.9, fontFace: HF, fontSize: 48, bold: true, color: K.white, margin: 0, isTextBox: true });
s.addText("Setembro de 2026", { x: MX, y: 3.75, w: 11, h: 0.8, fontFace: HF, fontSize: 40, bold: true, color: K.sky, margin: 0, isTextBox: true });
s.addText("Mercado Livre e Shopee: faturamento, mídia, afiliados e comparação com setembro de 2025.", { x: MX, y: 4.75, w: 8.5, h: 0.6, fontFace: BF, fontSize: 16, color: K.mist, margin: 0, isTextBox: true });
s.addText("puertasmarketing.com.br", { x: MX, y: 6.0, w: 5, h: 0.3, fontFace: BF, fontSize: 11, color: K.mist, margin: 0, isTextBox: true });

// ---------- 2. consolidado ----------
s = pres.addSlide({ masterName: "ESCURO" });
header(s, "Consolidado", "O maior mês da história dos dois canais juntos", true);
s.addText("Faturamento somado de Mercado Livre e Shopee em setembro", { x: MX, y: 1.8, w: CW, h: 0.4, fontFace: BF, fontSize: 16, color: K.mist, margin: 0, isTextBox: true });
{
  const y = 2.55, h = 2.6, g = 0.25, w3 = (CW - 2 * g) / 3;
  stat(s, MX, y, w3, h, "R$ 1,05 mi", "R$ 1.049.375 somando os dois canais: R$ 768 mil no ML e R$ 281 mil na Shopee", "Faturamento", true);
  stat(s, MX + w3 + g, y, w3, h, "+105%", "contra setembro de 2025, quando os dois canais somaram R$ 513 mil", "Ano a ano", true);
  stat(s, MX + 2 * (w3 + g), y, w3, h, "+5,9%", "sobre agosto (R$ 991 mil). Supera julho, o recorde anterior, com R$ 1,04 mi", "Mês a mês", true);
}
note(s, "Segunda vez na história acima de R$ 1 milhão no mês. A primeira foi julho de 2026.", 5.55, true);

// ---------- 3. divisor ML ----------
s = pres.addSlide({ masterName: "ESCURO" });
s.addText("01", { x: MX, y: 2.2, w: 3, h: 1.2, fontFace: HF, fontSize: 72, bold: true, color: K.sky, margin: 0, isTextBox: true });
s.addText("Mercado Livre", { x: MX, y: 3.4, w: CW, h: 0.9, fontFace: HF, fontSize: 44, bold: true, color: K.bg, margin: 0, isTextBox: true });
s.addText("Volume recorde, com mídia e afiliados puxando o crescimento", { x: MX, y: 4.3, w: CW, h: 0.5, fontFace: BF, fontSize: 18, color: K.mist, margin: 0, isTextBox: true });

// ---------- 4. ML volume ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Mercado Livre · Volume", "Recorde de vendas, unidades e visitas");
{
  const y = 2.05, h = 2.0, g = 0.25, w2 = (CW - g) / 2;
  const items = [
    ["5.238", "vendas no mês, ante 4.468 em agosto (+17%)", "Recorde"],
    ["5.525", "unidades vendidas, ante 4.728 em agosto (+17%)", "Recorde"],
    ["73.923", "visitas nos anúncios, ante 65.177 em agosto (+13%)", "Recorde"],
    ["4.102", "compradores únicos, sendo 1.134 recorrentes, os dois também recorde", "Recorde"],
  ];
  items.forEach((it, i) => stat(s, MX + (i % 2) * (w2 + g), y + Math.floor(i / 2) * (h + g), w2, h, it[0], it[1], it[2], false, { size: 32 }));
}
note(s, "Conversão de 7,09%, contra 6,86% em agosto.", 6.4);

// ---------- 5. ML faturamento ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Mercado Livre · Faturamento", "R$ 768 mil, segundo maior mês da série");
s.addText("Vendas brutas em R$ mil · set/25 a set/26 · só novembro de 2025, mês da Black Friday, foi maior", { x: MX, y: 1.75, w: CW, h: 0.35, fontFace: BF, fontSize: 14, color: K.muted, margin: 0, isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "Vendas brutas (R$ mil)", labels: ML_FAT.map(r => r[0]), values: ML_FAT.map(r => Math.round(r[1] / 1000)) }], {
  x: MX, y: 2.25, w: CW, h: 4.45, barDir: "col", barGapWidthPct: 45,
  chartColors: ML_FAT.map((r, i) => (i === ML_FAT.length - 1 ? K.blue : K.mist)),
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: K.navy, dataLabelFormatCode: "#,##0",
  catAxisLabelColor: K.muted, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  catAxisLineShow: false, showLegend: false, ...axisFont, dataLabelFontSize: 11,
});

// ---------- 6. ML midia e afiliados ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Mercado Livre · Mídia e afiliados", "Publicidade e afiliados batem recorde de receita");
{
  const y = 2.0, h = 2.35, g = 0.25, w4 = (CW - 3 * g) / 4;
  stat(s, MX, y, w4, h, "R$ 409 mil", "receita de Ads, ante R$ 401 mil em agosto", "Recorde", false, { size: 28 });
  stat(s, MX + (w4 + g), y, w4, h, "R$ 72,6 mil", "receita de afiliados, +72% sobre agosto (R$ 42,2 mil)", "Recorde", false, { size: 28 });
  stat(s, MX + 2 * (w4 + g), y, w4, h, "17,10", "ROAS, ante 17,45 em agosto. Investimento subiu 4% (R$ 23.913)", "Retorno", false, { size: 28 });
  stat(s, MX + 3 * (w4 + g), y, w4, h, "5,85%", "ACOS, ante 5,73% em agosto. TACOS estável em 3,1%", "Custo", false, { size: 28 });
}
card(s, MX, 4.65, CW, 1.4, false);
s.addText([
  { text: "Leitura do mês", options: { bold: true, color: K.navy, fontFace: HF, fontSize: 16, breakLine: true } },
  { text: "Ads gerou 53% do faturamento e afiliados 9,5%, empatado com junho como a maior fatia de afiliados da série. A mídia ficou levemente mais cara (ROAS caiu 2%), mas o volume extra compensou: a receita de Ads cresceu mesmo assim.", options: { color: K.muted, fontSize: 14 } },
], { x: MX + 0.3, y: 4.85, w: CW - 0.6, h: 1.05, fontFace: BF, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true });

// ---------- 7. ML ponto de atencao ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Mercado Livre · Ponto de atenção", "Ticket médio caiu ao menor nível da série");
{
  const lw = 3.6;
  stat(s, MX, 2.0, lw, 2.1, "R$ 146,62", "valor médio por venda em setembro, ante R$ 164,97 em agosto (−11%)", "Menor da série", false, { size: 32, tagColor: K.alert });
  card(s, MX, 4.35, lw, 2.05, false);
  s.addText("Mais vendas de menor valor: o preço médio por unidade caiu de R$ 155,90 para R$ 139,01. Vale cruzar com o mix de produtos e as promoções do mês.", {
    x: MX + 0.25, y: 4.55, w: lw - 0.5, h: 1.7, fontFace: BF, fontSize: 14, color: K.muted, margin: 0, valign: "top", isTextBox: true });
  const cx = MX + lw + 0.35, cw = CW - lw - 0.35;
  card(s, cx, 2.0, cw, 4.4, false);
  s.addText("Valor médio por venda (R$) · 2026", { x: cx + 0.3, y: 2.2, w: cw - 0.6, h: 0.3, fontFace: BF, fontSize: 12, bold: true, color: K.navy, margin: 0, isTextBox: true });
  s.addChart(pres.charts.LINE, [{ name: "Valor médio por venda", labels: ML_TICKET.map(r => r[0]), values: ML_TICKET.map(r => r[1]) }], {
    x: cx + 0.15, y: 2.55, w: cw - 0.3, h: 3.7, chartColors: [K.blue], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8,
    showValue: true, dataLabelPosition: "t", dataLabelColor: K.navy, dataLabelFormatCode: "0",
    valAxisMinVal: 120, valAxisMaxVal: 190, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelColor: K.muted, showLegend: false, ...axisFont,
  });
}

// ---------- 8. divisor Shopee ----------
s = pres.addSlide({ masterName: "ESCURO" });
s.addText("02", { x: MX, y: 2.2, w: 3, h: 1.2, fontFace: HF, fontSize: 72, bold: true, color: K.sky, margin: 0, isTextBox: true });
s.addText("Shopee", { x: MX, y: 3.4, w: CW, h: 0.9, fontFace: HF, fontSize: 44, bold: true, color: K.bg, margin: 0, isTextBox: true });
s.addText("A mídia mais eficiente desde o início da operação, pelo segundo mês seguido", { x: MX, y: 4.3, w: CW, h: 0.5, fontFace: BF, fontSize: 18, color: K.mist, margin: 0, isTextBox: true });

// ---------- 9. Shopee publicidade ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Shopee · Publicidade", "Gastamos menos e vendemos mais com Ads");
{
  const y = 2.0, h = 2.35, g = 0.25, w4 = (CW - 3 * g) / 4;
  stat(s, MX, y, w4, h, "17,35", "ROAS, ante 14,14 em agosto", "Recorde", false, { size: 30 });
  stat(s, MX + (w4 + g), y, w4, h, "13,55", "ROAS direto, ante 10,67 em agosto", "Recorde", false, { size: 30 });
  stat(s, MX + 2 * (w4 + g), y, w4, h, "5,76%", "ACOS, ante 6,07% em agosto", "Menor da série", false, { size: 30 });
  stat(s, MX + 3 * (w4 + g), y, w4, h, "−16%", "despesa de mídia: de R$ 14.414 para R$ 12.050", "Investimento", false, { size: 30 });
}
card(s, MX, 4.65, CW, 1.4, false);
s.addText([
  { text: "Leitura do mês", options: { bold: true, color: K.navy, fontFace: HF, fontSize: 16, breakLine: true } },
  { text: "Com 16% menos investimento, o GMV de Ads subiu 2,6% (R$ 209 mil) e as conversões diretas bateram recorde (1.301). A loja fechou com R$ 281 mil em pedidos pagos, segundo maior mês, e ticket médio de R$ 151,60.", options: { color: K.muted, fontSize: 14 } },
], { x: MX + 0.3, y: 4.85, w: CW - 0.6, h: 1.05, fontFace: BF, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true });

// ---------- 10. Shopee eficiencia ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Shopee · Eficiência", "ACOS caiu para menos da metade em 2026");
s.addText("Custo de mídia sobre a receita gerada por Ads, em % · de 12,0% em fevereiro para 5,8% em setembro", { x: MX, y: 1.75, w: CW, h: 0.35, fontFace: BF, fontSize: 14, color: K.muted, margin: 0, isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "ACOS (%)", labels: SH_ACOS.map(r => r[0]), values: SH_ACOS.map(r => r[1]) }], {
  x: MX, y: 2.25, w: CW, h: 4.45, barDir: "col", barGapWidthPct: 45,
  chartColors: SH_ACOS.map((r, i) => (i === SH_ACOS.length - 1 ? K.blue : K.mist)),
  showValue: true, dataLabelPosition: "outEnd", dataLabelColor: K.navy, dataLabelFormatCode: "0.0",
  catAxisLabelColor: K.muted, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  catAxisLineShow: false, showLegend: false, ...axisFont, dataLabelFontSize: 11,
});

// ---------- 11. ano contra ano ----------
s = pres.addSlide({ masterName: "CLARO" });
header(s, "Setembro 2025 vs setembro 2026", "Os dois canais seguem multiplicando de tamanho");
{
  const lw = 3.6;
  stat(s, MX, 2.0, lw, 2.1, "+79%", "Mercado Livre: de R$ 430 mil para R$ 768 mil", "Mercado Livre", false, { size: 36 });
  stat(s, MX, 4.3, lw, 2.1, "+240%", "Shopee: de R$ 83 mil para R$ 281 mil", "Shopee", false, { size: 36 });
  const cx = MX + lw + 0.35, cw = CW - lw - 0.35;
  card(s, cx, 2.0, cw, 4.4, false);
  s.addText("Faturamento em R$ mil", { x: cx + 0.3, y: 2.2, w: cw - 0.6, h: 0.3, fontFace: BF, fontSize: 12, bold: true, color: K.navy, margin: 0, isTextBox: true });
  s.addChart(pres.charts.BAR, [
    { name: "set/25", labels: ["Mercado Livre", "Shopee"], values: [430, 83] },
    { name: "set/26", labels: ["Mercado Livre", "Shopee"], values: [768, 281] },
  ], {
    x: cx + 0.15, y: 2.55, w: cw - 0.3, h: 3.7, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [K.mist, K.blue], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: K.navy, dataLabelFormatCode: "#,##0",
    catAxisLabelColor: K.navy, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendColor: K.muted, ...axisFont, catAxisLabelFontSize: 13,
  });
}

// ---------- 12. resumo ----------
s = pres.addSlide({ masterName: "ESCURO" });
header(s, "Resumo", "O que setembro mostrou", true);
{
  const y = 2.05, h = 3.6, g = 0.25, w3 = (CW - 2 * g) / 3;
  const items = [
    ["01", "Escala recorde", "R$ 1,05 mi somados, maior mês da história, com recorde de vendas e visitas no Mercado Livre"],
    ["02", "Mídia mais eficiente", "Shopee com o menor ACOS e o maior ROAS da série, gastando 16% menos"],
    ["03", "Atenção ao ticket", "Ticket médio do ML no menor nível (R$ 146,62): crescemos em volume, com vendas de menor valor"],
  ];
  items.forEach((it, i) => {
    const x = MX + i * (w3 + g);
    card(s, x, y, w3, h, true);
    s.addText(it[0], { x: x + 0.3, y: y + 0.3, w: 1.2, h: 0.7, fontFace: HF, fontSize: 32, bold: true, color: K.sky, margin: 0, isTextBox: true });
    s.addText(it[1], { x: x + 0.3, y: y + 1.1, w: w3 - 0.6, h: 0.5, fontFace: HF, fontSize: 18, bold: true, color: K.bg, margin: 0, isTextBox: true });
    s.addText(it[2], { x: x + 0.3, y: y + 1.7, w: w3 - 0.6, h: 1.7, fontFace: BF, fontSize: 14, color: K.mist, margin: 0, valign: "top", isTextBox: true });
  });
}

// ---------- 13. encerramento ----------
s = pres.addSlide({ masterName: "CAPA" });
s.background = { color: K.blue };
s.addImage({ path: "logo_todo_branco.png", x: MX, y: 0.89, w: 2.92, h: 0.68, altText: "Puertas Marketing Digital" });
s.addText("Obrigado!", { x: MX, y: 2.9, w: CW, h: 1.0, fontFace: HF, fontSize: 48, bold: true, color: K.white, margin: 0, isTextBox: true });
s.addText("Próximo relatório: outubro de 2026.", { x: MX, y: 3.9, w: CW, h: 0.5, fontFace: BF, fontSize: 16, color: K.pale, margin: 0, isTextBox: true });
s.addText("WHATSAPP", { x: MX, y: 5.6, w: 3, h: 0.25, fontFace: BF, fontSize: 10, bold: true, color: K.pale, margin: 0, isTextBox: true });
s.addText("+55 11 91689-6015", { x: MX, y: 5.88, w: 3.5, h: 0.4, fontFace: HF, fontSize: 16, bold: true, color: K.white, margin: 0, isTextBox: true });
s.addText("E-MAIL", { x: MX + 3.9, y: 5.6, w: 3, h: 0.25, fontFace: BF, fontSize: 10, bold: true, color: K.pale, margin: 0, isTextBox: true });
s.addText("matheuspuertas@puertasmarketing.com.br", { x: MX + 3.9, y: 5.88, w: 6.5, h: 0.4, fontFace: HF, fontSize: 16, bold: true, color: K.white, margin: 0, isTextBox: true });

(async () => {
  await pres.writeFile({ fileName: "Performance_Marketplaces_Setembro_2026.pptx" });
  await applyTheme("Performance_Marketplaces_Setembro_2026.pptx", THEME);
  console.log("ok");
})();
