// Gera o deck "Business Plan Reviews 2027" a partir dos números da planilha
// "Blueprint @ Budget Reviews 2027" (aba SaaS Plan Template).
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "Business_Plan_Reviews_2027.pptx";

// ---------------------------------------------------------------- dados
const M = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
const sum = (a) => a.reduce((s, v) => s + v, 0);
const eopContratos = [396, 400, 423, 449, 488, 540, 592, 649, 705, 754, 769, 770];
const eopCheckout = [450, 427, 405, 384, 364, 345, 327, 310, 294, 278, 264, 250];
const newDed = [3, 8, 18, 18, 27, 36, 37, 46, 53, 52, 26, 15];
const newShr = [1, 4, 13, 16, 21, 26, 26, 22, 16, 12, 4, 1];
const churnCt = [8, 8, 8, 8, 9, 10, 11, 12, 13, 14, 15, 15];
const revCt = [705600, 712800, 719919, 761591, 807761, 877601, 972189, 1066370, 1167803, 1268219, 1357605, 1384903];
const revCk = [34128, 34128, 32353, 30671, 29076, 27564, 26131, 24772, 23484, 22263, 21105, 20008];
const revGross = revCt.map((v, i) => v + revCk[i]);
const revNet = revGross.map((v) => v * 0.78);
const reviewsCt = [301950, 304966, 322618, 342177, 371762, 411830, 451726, 494694, 537232, 575097, 586660, 587363];
// Corrigido: reviews de Checkout sobre a base de Checkout (a planilha usa a base total)
const reviewsCk = eopCheckout.map((c) => (c * 9000) / 475);
const reviewsTot = reviewsCt.map((v, i) => v + reviewsCk[i]);
const sales = [42000, 42000, 42000, 84000, 84000, 84000, 126000, 126000, 126000, 126000, 126000, 126000];
const paid = [10000, 12000, 14000, 16000, 18000, 20000, 22000, 24000, 26000, 28000, 30000, 32000];
const events = [0, 0, 0, 40000, 150000, 55000, 105000, 75000, 55000, 0, 0, 0];
const mktPayroll = Array(12).fill(30000); // Growth Engineer (pós-encargos)
const mkt = paid.map((v, i) => v + events[i] + mktPayroll[i]);
const pd = Array(12).fill(185449 + 30000); // + Data Engineer (pós-encargos)
const ga = Array(12).fill(40000 / 12);
const opex = M.map((_, i) => sales[i] + mkt[i] + pd[i] + ga[i]);
const ebitda = M.map((_, i) => revNet[i] - opex[i]);
const ebitdaSheet = 5328490;
const marginM = ebitda.map((v, i) => v / revGross[i]);
const leadsGen = [171, 206, 240, 959, 2878, 1285, 2176, 1696, 1388, 480, 514, 548];
const leadsReq = [100, 297, 779, 853, 1194, 1558, 1578, 1705, 1719, 1594, 756, 408];

const T = {
  gross: sum(revGross), net: sum(revNet), tax: sum(revGross) * 0.22,
  sales: sum(sales), mkt: sum(mkt), mktSpend: sum(paid) + sum(events), pd: sum(pd), ga: sum(ga),
  opex: sum(opex), ebitda: sum(ebitda), newDed: sum(newDed), newShr: sum(newShr), churn: sum(churnCt),
  reviews: sum(reviewsTot), reviewsCt: sum(reviewsCt),
};
T.newTot = T.newDed + T.newShr;
T.ir = T.ebitda * 0.22;
T.result = T.ebitda - T.ir;
T.cac = (T.sales + T.mkt) / T.newTot;
T.cacDed = (T.sales + T.mkt) / T.newDed;
T.ltv = (1800 * 0.7) / 0.02;
T.cogsHyp = T.net * 0.295;

const q = (arr) => [0, 1, 2, 3].map((k) => sum(arr.slice(k * 3, k * 3 + 3)));

const fmtInt = (v) => Math.round(v).toLocaleString("pt-BR");
const mi = (v, d = 1) => (v / 1e6).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const mil = (v) => Math.round(v / 1000).toLocaleString("pt-BR");
const pct = (v, d = 0) => (v * 100).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d }) + "%";

console.log({
  gross: T.gross, net: T.net, opex: T.opex, ebitda: T.ebitda, margin: T.ebitda / T.gross, result: T.result,
  cac: T.cac, cacDed: T.cacDed, reviews: T.reviews, sharePd: T.pd / T.opex, shareSales: T.sales / T.opex,
  shareMkt: T.mkt / T.opex, marginM: marginM.map((m) => m.toFixed(3)).join(" "),
});

// ---------------------------------------------------------------- tema
const THEME = {
  name: "Reviews 2027",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2420", lt1: "FFFFFF", dk2: "003330", lt2: "E6F1EB",
    accent1: "007535", accent2: "90B823", accent3: "5CA77E", accent4: "D99A00",
    accent5: "8AC0A2", accent6: "B83A26", hlink: "007535", folHlink: "005931",
  },
};
const HEX = {
  green: "007535", lime: "90B823", mid: "5CA77E", soft: "8AC0A2", dark: "003330", amber: "D99A00",
  red: "B83A26", grid: "E3E8E5", axis: "55635C", tint: "E6F1EB",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "Reviews · Reclame AQUI";
pres.title = "Business Plan Reviews 2027";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------------------------------------------------------------- ícones
async function icon(Comp, color = "FFFFFF") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------------------------------------------------------------- layouts
const footerObjs = (dark) => [
  {
    text: {
      text: "Business Plan Reviews 2027 · Reclame AQUI",
      options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: dark ? C.background2 : C.accent3, margin: 0 },
    },
  },
];

pres.defineSlideMaster({
  title: "CAPA",
  background: { color: HEX.dark },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.9, w: 11.7, h: 0.5, fontSize: 16, bold: true, color: C.accent2, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.45, w: 11.7, h: 1.4, fontSize: 60, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.95, w: 10.5, h: 1.1, fontSize: 24, color: C.background2, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 14, color: C.accent5, margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "SECAO",
  background: { color: HEX.dark },
  objects: [
    ...footerObjs(true),
    { placeholder: { options: { name: "num", type: "body", x: 0.8, y: 2.0, w: 3, h: 1.1, fontSize: 66, bold: true, color: C.accent2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 3.15, w: 11.5, h: 1.0, fontSize: 44, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.25, w: 11.5, h: 0.8, fontSize: 20, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.background2, align: "right" },
});

pres.defineSlideMaster({
  title: "CONTEUDO",
  background: { color: "FFFFFF" },
  margin: [0.5, 0.6, 0.7, 0.6],
  objects: [
    ...footerObjs(false),
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.4, w: 12.1, h: 0.32, fontSize: 12, bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.75, w: 12.1, h: 0.95, fontSize: 28, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent3, align: "right" },
});

pres.defineSlideMaster({
  title: "FECHAMENTO",
  background: { color: HEX.dark },
  objects: [
    ...footerObjs(true),
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.4, w: 12.1, h: 0.32, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.75, w: 12.1, h: 0.95, fontSize: 28, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.background2, align: "right" },
});

// ---------------------------------------------------------------- helpers
let SECTION = "";
function content(kicker, title, layout = "CONTEUDO") {
  const s = pres.addSlide({ masterName: layout, sectionTitle: SECTION });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function section(num, title, body) {
  SECTION = title;
  pres.addSection({ title });
  const s = pres.addSlide({ masterName: "SECAO", sectionTitle: SECTION });
  s.addText(num, { placeholder: "num" });
  s.addText(title, { placeholder: "title" });
  s.addText(body, { placeholder: "body" });
  return s;
}
let objN = 0;
const on = (p) => `${p}-${++objN}`;

function card(s, x, y, w, h, fill = C.background2) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: 0.12, objectName: on("card") });
}
function stat(s, x, y, w, h, big, label, sub, opts = {}) {
  card(s, x, y, w, h, opts.fill || C.background2);
  s.addText(big, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.75, fontSize: opts.bigSize || 36, bold: true, color: opts.bigColor || C.accent1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true, objectName: on("stat-num"), valign: "middle" });
  s.addText(label, { x: x + 0.25, y: y + 0.98, w: w - 0.5, h: 0.42, fontSize: 15, bold: true, color: opts.labelColor || C.text1, margin: 0, isTextBox: true, objectName: on("stat-label"), valign: "top" });
  if (sub) s.addText(sub, { x: x + 0.25, y: y + 1.4, w: w - 0.5, h: h - 1.5, fontSize: 12, color: opts.subColor || C.text1, margin: 0, isTextBox: true, objectName: on("stat-sub"), valign: "top" });
}
function iconCircle(s, data, x, y, d = 0.6, fill = C.accent1) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: on("icon-bg") });
  s.addImage({ data, x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, objectName: on("icon"), altText: "ícone" });
}
function note(s, text, y = 6.55, dark = false) {
  s.addText(text, { x: 0.6, y, w: 12.1, h: 0.35, fontSize: 10, italic: true, color: dark ? C.background2 : C.accent3, margin: 0, isTextBox: true, objectName: on("nota") });
}
function chartBase(extra = {}) {
  return Object.assign({
    catAxisLabelColor: HEX.axis, valAxisLabelColor: HEX.axis,
    catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontSize: 10, legendFontSize: 11,
    valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" },
    catAxisLineShow: false, valAxisLineShow: false,
    showTitle: true, titleFontSize: 13, titleColor: HEX.dark, titleBold: true,
    dataLabelColor: HEX.dark, legendColor: HEX.axis,
  }, extra);
}

// ---------------------------------------------------------------- slides
(async () => {
  const I = {
    star: await icon(fa.FaStar), user: await icon(fa.FaUserCheck), store: await icon(fa.FaStore),
    contract: await icon(fa.FaFileSignature), cart: await icon(fa.FaShoppingCart), chart: await icon(fa.FaChartLine),
    team: await icon(fa.FaUsers), bull: await icon(fa.FaBullhorn), code: await icon(fa.FaCode), coins: await icon(fa.FaCoins),
    warn: await icon(fa.FaExclamationTriangle), check: await icon(fa.FaCheck), clock: await icon(fa.FaClock),
    handshake: await icon(fa.FaHandshake), funnel: await icon(fa.FaFilter), arrow: await icon(fa.FaArrowRight, HEX.mid),
    server: await icon(fa.FaServer), sync: await icon(fa.FaSyncAlt),
  };

  // 1 — Capa
  SECTION = "Abertura";
  pres.addSection({ title: SECTION });
  let s = pres.addSlide({ masterName: "CAPA", sectionTitle: SECTION });
  s.addText("BUSINESS PLAN · BUDGET 2027", { placeholder: "kicker" });
  s.addText("Reviews 2027", { placeholder: "title" });
  s.addText("De 400 para 770 contratos e R$ 12,1 mi de receita bruta, com o time comercial triplicando", { placeholder: "body" });
  s.addText("Reclame AQUI · Outubro de 2026", { placeholder: "date" });
  s.addNotes("Abertura. Este é o plano de negócio de Reviews para 2027, construído a partir do Blueprint de Budget. A mensagem central: com um time comercial dedicado que cresce de 3 para 9 Sales Specialists, a base de contratos quase dobra e a receita bruta chega a R$ 12,1 milhões no ano.");

  // 2 — Resumo executivo
  s = content("RESUMO EXECUTIVO", "Em 2027, Reviews quase dobra e se paga com folga");
  const sw = 2.85, sg = 0.233;
  [
    ["R$ 12,1 mi", "Receita bruta 2027", "MRR 1,9x: R$ 740 mil em jan → R$ 1,4 mi em dez"],
    ["770", "Contratos em dez/27", "+93% vs. 400 contratos de partida"],
    ["~5,4 mi", "Reviews no ano", "~310 mil → ~590 mil reviews por mês"],
    [`R$ ${mi(T.ebitda)} mi`, "EBITDA antes de COGS", `${pct(T.ebitda / T.gross)} da receita bruta; COGS a fechar`],
  ].forEach((d, i) => stat(s, 0.6 + i * (sw + sg), 1.85, sw, 2.05, d[0], d[1], d[2], { bigSize: 28 }));
  [
    [I.handshake, "Motor de crescimento", "Venda consultiva de Contratos (ticket de R$ 1.800/mês): 501 novos contratos, 2/3 vindos do time dedicado"],
    [I.coins, "Investimento", `R$ ${mi(T.opex)} mi de opex no ano: P&D ${pct(T.pd / T.opex)}, Vendas ${pct(T.sales / T.opex)}, Marketing ${pct(T.mkt / T.opex)}`],
    [I.check, "O que pedimos", "Aprovar o budget e o plano de contratação em 3 ondas: nov/26, abr/27 e jul/27"],
  ].forEach((d, i) => {
    const y = 4.25 + i * 0.75;
    iconCircle(s, d[0], 0.6, y, 0.55);
    s.addText([
      { text: d[1] + "  ", options: { bold: true, color: C.text2 } },
      { text: d[2], options: { color: C.text1 } },
    ], { x: 1.35, y, w: 11.35, h: 0.55, fontSize: 15, valign: "middle", margin: 0, isTextBox: true, objectName: on("resumo") });
  });
  s.addNotes(`Quatro números resumem o plano: R$ 12,1 mi de receita bruta, 770 contratos ativos no fim do ano, cerca de 5,4 milhões de reviews coletados e R$ ${mi(T.ebitda)} mi de EBITDA antes de COGS. O EBITDA aqui já inclui a folha do Growth Engineer e do Data Engineer, que estão no plano de contratação mas ficaram fora do total da planilha. COGS ainda não foi estimado e é tratado na sensibilidade.`);

  // 3 — Por que Reviews
  s = content("POR QUE REVIEWS", "Reviews transforma opinião real em ferramenta de venda");
  const cw = 3.6, cx = [0.6, 4.85, 9.1];
  [
    [I.user, "Consumidor", "cliente de impacto", "Quer adquirir com confiança, ajudando outros consumidores. Ao avaliar uma compra, gera um review."],
    [I.star, "Opinião real", "a matéria-prima", "Review ou reclamação avaliada: experiência humana e verdadeira, transformada em dado público e verificável."],
    [I.store, "Marca", "cliente econômico", "Quer vender mais e ser reconhecida como mais confiável, usando reputação e reviews como diferencial."],
  ].forEach((d, i) => {
    card(s, cx[i], 2.0, cw, 3.9, i === 1 ? C.text2 : C.background2);
    iconCircle(s, d[0], cx[i] + 0.35, 2.35, 0.8, i === 1 ? C.accent2 : C.accent1);
    s.addText(d[1], { x: cx[i] + 0.35, y: 3.35, w: cw - 0.7, h: 0.5, fontSize: 22, bold: true, color: i === 1 ? C.background1 : C.text2, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[2], { x: cx[i] + 0.35, y: 3.85, w: cw - 0.7, h: 0.35, fontSize: 13, italic: true, color: i === 1 ? C.accent2 : C.accent1, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[3], { x: cx[i] + 0.35, y: 4.3, w: cw - 0.7, h: 1.4, fontSize: 15, color: i === 1 ? C.background1 : C.text1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  });
  [4.27, 8.52].forEach((x) => s.addImage({ data: I.arrow, x, y: 3.7, w: 0.45, h: 0.45, objectName: on("seta"), altText: "seta" }));
  note(s, "Quem gera o valor (consumidor) não é quem paga por ele (marca). Fonte: memória corporativa RA — JTBD e Opinião Real.", 6.3);
  s.addNotes("Reviews conecta os dois jobs do RA. O consumidor quer comprar com confiança e, ao avaliar, gera opinião real. A marca compra o uso dessa opinião real como ferramenta de venda. A assimetria importa: a qualidade da experiência do consumidor é a condição de existência do ativo que a marca paga.");

  // 4 — Ponto de partida
  s = content("PONTO DE PARTIDA · AGO/26", "Hoje: 875 clientes em dois modelos, e Contratos é o motor");
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 14, align: "center" } });
  const rowsPP = [
    ["Clientes pagantes", "400", "475"],
    ["Reviews por mês", "305 mil", "9 mil"],
    ["Reviews por cliente/mês", "762", "19"],
    ["ARPU mensal", "R$ 1.800", "R$ 76"],
    ["Churn mensal", "2,0%", "5,2%"],
    ["MRR", "R$ 720 mil", "R$ 36 mil"],
  ];
  s.addTable([
    [{ text: "", options: { fill: { color: C.text2 } } }, hdr("Contratos"), hdr("Checkout")],
    ...rowsPP.map((r, i) => r.map((t, j) => ({ text: t, options: { fontSize: 15, bold: j === 1, color: j === 1 ? C.accent1 : C.text1, align: j ? "center" : "left", fill: { color: i % 2 ? "FFFFFF" : HEX.tint } } }))),
  ], { x: 0.6, y: 1.95, w: 7.4, colW: [3.0, 2.2, 2.2], rowH: 0.58, border: { type: "none" }, margin: [0, 0.15, 0, 0.15], valign: "middle", objectName: on("tabela") });
  card(s, 8.5, 1.95, 4.2, 4.06, C.text2);
  s.addText("Contratos já é", { x: 8.8, y: 2.2, w: 3.6, h: 0.35, fontSize: 14, color: C.background2, margin: 0, isTextBox: true, objectName: on("t") });
  [["95%", "do MRR"], ["97%", "dos reviews"]].forEach((d, i) => {
    s.addText(d[0], { x: 8.8, y: 2.6 + i * 0.85, w: 1.7, h: 0.7, fontSize: 36, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[1], { x: 10.5, y: 2.6 + i * 0.85, w: 2.0, h: 0.7, fontSize: 18, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: on("t") });
  });
  s.addText("Por cliente: 24x o ticket, 40x os reviews e 2,6x menos churn que o Checkout", { x: 8.8, y: 4.45, w: 3.6, h: 1.3, fontSize: 15, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  s.addNotes("Partimos de 400 contratos e 475 clientes de Checkout. Contratos concentra 95% do MRR e 97% dos reviews. Por cliente, um contrato vale 24 vezes o ticket do Checkout, gera 40 vezes mais reviews e tem churn 2,6 vezes menor. Por isso o plano de 2027 concentra toda a aquisição em Contratos.");

  // 5 — Seção 1
  section("01", "Crescimento da base", "Clientes, aquisição, capacidade comercial, marketing e reviews");

  // 6 — Base de clientes
  s = content("BASE DE CLIENTES", "Contratos quase dobram; Checkout encolhe sem novas vendas");
  s.addChart(pres.charts.LINE, [
    { name: "Contratos", labels: M, values: eopContratos },
    { name: "Checkout", labels: M, values: eopCheckout },
  ], chartBase({ x: 0.6, y: 1.85, w: 8.2, h: 4.6, title: "Clientes pagantes no fim do mês", chartColors: [HEX.green, HEX.soft], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 900, valAxisMajorUnit: 150 }));
  [
    ["400 → 770", "Contratos", "+93% no ano", C.accent1],
    ["475 → 250", "Checkout", "−47%: sem aquisição, churn de 5,2%/mês", C.accent3],
    ["875 → ~1.020", "Base total", "+17% de clientes pagantes", C.text2],
  ].forEach((d, i) => {
    card(s, 9.2, 1.85 + i * 1.6, 3.5, 1.45);
    s.addText(d[0], { x: 9.45, y: 1.95 + i * 1.6, w: 3.1, h: 0.6, fontSize: 26, bold: true, color: d[3], margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText([{ text: d[1], options: { bold: true, breakLine: true } }, { text: d[2] }], { x: 9.45, y: 2.55 + i * 1.6, w: 3.1, h: 0.7, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  });
  s.addNotes("A base de contratos sai de 400 e termina o ano em 770, alta de 93%. O Checkout não recebe aquisição no plano e encolhe de 475 para 250 clientes por causa do churn de 5,2% ao mês. A base total cresce 17%, mas o que importa é o mix: ele migra para o modelo de maior valor.");

  // 7 — Aquisição
  s = content("AQUISIÇÃO", "501 novos contratos no ano, 2/3 vindos do time dedicado");
  s.addChart(pres.charts.BAR, [
    { name: "Operação dedicada", labels: M, values: newDed },
    { name: "Operação compartilhada", labels: M, values: newShr },
  ], chartBase({ x: 0.6, y: 1.85, w: 8.2, h: 4.6, barDir: "col", barGrouping: "stacked", title: "Novos contratos por mês", chartColors: [HEX.green, HEX.lime], showLegend: true, legendPos: "b", showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", dataLabelFontSize: 9, barGapWidthPct: 45, valAxisMaxVal: 80, valAxisMajorUnit: 20 }));
  [
    [String(T.newTot), "novos contratos", C.accent1],
    [`${T.newDed} (${pct(T.newDed / T.newTot)})`, "operação dedicada", C.accent1],
    [`${T.newShr} (${pct(T.newShr / T.newTot)})`, "operação compartilhada", C.accent2],
    [`−${T.churn}`, "churn (2%/mês) → saldo +370", C.accent6],
  ].forEach((d, i) => {
    s.addText(d[0], { x: 9.2, y: 1.95 + i * 1.08, w: 3.5, h: 0.55, fontSize: 28, bold: true, color: d[2], margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[1], { x: 9.2, y: 2.5 + i * 1.08, w: 3.5, h: 0.35, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: on("t") });
  });
  note(s, "Sazonalidade: jan e dez a 25% da produtividade; fev e nov a 50%. Pico de vendas entre mai e out.");
  s.addNotes(`São ${T.newTot} novos contratos no ano: ${T.newDed} da operação dedicada e ${T.newShr} da operação compartilhada, que é o time comercial do RA vendendo Reviews junto com outros produtos. Descontado o churn de ${T.churn} contratos, o saldo líquido é de +370. As vendas se concentram de maio a outubro; janeiro, fevereiro, novembro e dezembro têm sazonalidade forte.`);

  // 8 — Capacidade comercial
  s = content("CAPACIDADE COMERCIAL", "Time dedicado cresce em três ondas: 3 → 6 → 9 vendedores");
  s.addShape(pres.shapes.LINE, { x: 1.2, y: 2.75, w: 10.9, h: 0, line: { color: C.accent5, width: 2 }, objectName: on("linha-tempo") });
  [
    ["nov/26", "3 SS", "Onda 1: time inicial + Growth Engineer + Data Engineer"],
    ["abr/27", "6 SS", "Onda 2: +3 Sales Specialists, a tempo do ciclo de eventos"],
    ["jul/27", "9 SS", "Onda 3: +3 Sales Specialists para o pico do 2º semestre"],
  ].forEach((d, i) => {
    const x = 1.2 + i * 4.0;
    s.addShape(pres.shapes.OVAL, { x: x - 0.22, y: 2.53, w: 0.44, h: 0.44, fill: { color: C.accent1 }, line: { color: C.background1, width: 3 }, objectName: on("marco") });
    s.addText(d[0], { x: x - 0.22, y: 1.9, w: 2.5, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[1], { x: x - 0.22, y: 3.1, w: 2.5, h: 0.55, fontSize: 28, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[2], { x: x - 0.22, y: 3.65, w: 3.6, h: 0.7, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  });
  [
    ["5 vendas/mês", "produtividade base do Sales Specialist dedicado (0,25 por dia útil)"],
    ["2 vendas/mês", "por vendedor ativo da operação compartilhada (2 → 12 → 2 ativos)"],
    ["0% · 50% · 100%", "ramp-up do vendedor novo no 1º, 2º e 3º mês"],
    ["4,2 e 1,9", "vendas/mês realizadas em média: dedicado e compartilhado"],
  ].forEach((d, i) => {
    const x = 0.6 + i * 3.07;
    card(s, x, 4.65, 2.85, 1.75);
    s.addText(d[0], { x: x + 0.2, y: 4.8, w: 2.5, h: 0.55, fontSize: 22, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[1], { x: x + 0.2, y: 5.4, w: 2.5, h: 0.9, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  });
  s.addNotes("A capacidade comercial é o que dirige o plano: as vendas são calculadas a partir de quantos vendedores temos e de quanto cada um produz. Começamos com 3 Sales Specialists dedicados em novembro de 2026, vamos para 6 em abril e 9 em julho. A produtividade base é de 5 vendas por mês, com ramp-up de 0, 50 e 100% nos três primeiros meses, e ganho de produtividade de 5% a 20% ao longo dos primeiros meses. Na média do ano, o dedicado entrega 4,2 vendas por mês e o compartilhado 1,9.");

  // 9 — Marketing
  s = content("MARKETING", "R$ 732 mil para ~12,5 mil leads, a CPL de R$ 58");
  const fun = [["~12,5 mil", "leads", 4.6], ["~2,5 mil", "leads qualificados", 3.7], ["501", "novos contratos", 2.8]];
  fun.forEach((d, i) => {
    const y = 1.95 + i * 1.45, x = 0.6 + (4.6 - d[2]) / 2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: d[2], h: 1.05, fill: { color: [C.accent5, C.accent3, C.accent1][i] }, line: { color: "FFFFFF", width: 0 }, rectRadius: 0.1, objectName: on("funil") });
    s.addText([{ text: d[0], options: { bold: true, fontSize: 24, breakLine: true } }, { text: d[1], options: { fontSize: 13 } }], { x, y, w: d[2], h: 1.05, align: "center", valign: "middle", color: i === 0 ? C.text2 : C.background1, margin: 0, isTextBox: true, objectName: on("t") });
    if (i < 2) s.addText("20%", { x: 4.2, y: y + 1.07, w: 1.0, h: 0.35, fontSize: 12, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: on("t") });
  });
  s.addChart(pres.charts.BAR, [
    { name: "Mídia paga", labels: M, values: paid.map((v) => v / 1000) },
    { name: "Eventos", labels: M, values: events.map((v) => v / 1000) },
  ], chartBase({ x: 5.6, y: 1.85, w: 7.1, h: 4.55, barDir: "col", barGrouping: "stacked", title: "Investimento em marketing por mês (R$ mil)", chartColors: [HEX.mid, HEX.lime], showLegend: true, legendPos: "b", barGapWidthPct: 45, valAxisMaxVal: 180, valAxisMajorUnit: 30 }));
  note(s, "Eventos somam R$ 480 mil (66% do investimento), entre abr e set. Mídia paga começa em R$ 10 mil e cresce R$ 2 mil por mês.");
  s.addNotes("O funil foi desenhado de trás para frente: para entregar 501 contratos com conversão de 20% em cada etapa, precisamos de cerca de 12,5 mil leads. Com R$ 732 mil de mídia e eventos, o custo por lead máximo é de R$ 58. Dois terços do investimento estão em eventos, concentrados de abril a setembro. Isso cria um descasamento no primeiro e no quarto trimestres, que tratamos nos riscos.");

  // 10 — Reviews
  s = content("REVIEWS", "Reviews por mês quase dobram: ~310 mil → ~590 mil");
  s.addChart(pres.charts.BAR, [{ name: "Reviews (mil)", labels: M, values: reviewsTot.map((v) => Math.round(v / 1000)) }],
    chartBase({ x: 0.6, y: 1.85, w: 8.2, h: 4.6, barDir: "col", title: "Reviews coletados por mês (mil)", chartColors: [HEX.green], showLegend: false, showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 40, valAxisMaxVal: 700, valAxisMajorUnit: 100 }));
  stat(s, 9.2, 1.85, 3.5, 1.4, "~5,4 mi", "reviews no ano", null);
  stat(s, 9.2, 3.45, 3.5, 1.4, pct(T.reviewsCt / T.reviews), "via Contratos", null);
  stat(s, 9.2, 5.05, 3.5, 1.4, "762", "reviews por contrato/mês", null);
  s.addNotes("Cada contrato gera em média 762 reviews por mês, então a curva de reviews acompanha a base de contratos: de cerca de 310 mil para 590 mil por mês, perto de 5,4 milhões no ano, 99% via Contratos. A planilha mostra 5,5 milhões porque calcula os reviews de Checkout sobre a base total de clientes; corrigimos para a base de Checkout. Os subtotais por modelo da planilha também somam só parte dos meses.");

  // 11 — Seção 2
  section("02", "Resultado financeiro", "Receita, DRE, estrutura de custos e unit economics");

  // 12 — Receita
  s = content("RECEITA", "R$ 12,1 mi de receita bruta e MRR de saída de R$ 1,4 mi");
  s.addChart(pres.charts.BAR, [
    { name: "Contratos", labels: M, values: revCt.map((v) => Math.round(v / 1000)) },
    { name: "Checkout", labels: M, values: revCk.map((v) => Math.round(v / 1000)) },
  ], chartBase({ x: 0.6, y: 1.85, w: 8.2, h: 4.6, barDir: "col", barGrouping: "stacked", title: "Receita bruta por mês (R$ mil)", chartColors: [HEX.green, HEX.soft], showLegend: true, legendPos: "b", barGapWidthPct: 40, valAxisMaxVal: 1600, valAxisMajorUnit: 400, valAxisLabelFormatCode: "#,##0" }));
  [
    [`R$ ${mi(T.gross)} mi`, "receita bruta no ano"],
    [`R$ ${mi(T.net)} mi`, "receita líquida (tributos de 22%)"],
    ["97%", "da receita vem de Contratos"],
    ["~R$ 16,9 mi", "ARR de saída (MRR de dez × 12)"],
  ].forEach((d, i) => {
    s.addText(d[0], { x: 9.2, y: 1.95 + i * 1.08, w: 3.5, h: 0.55, fontSize: 28, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[1], { x: 9.2, y: 2.5 + i * 1.08, w: 3.5, h: 0.35, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: on("t") });
  });
  note(s, "A receita de Contratos é faturada sobre a base do mês anterior. Tributos sobre vendas de 22% (pós-reforma tributária).");
  s.addNotes("A receita bruta vai de R$ 740 mil em janeiro para R$ 1,4 milhão em dezembro, R$ 12,1 milhões no ano. Depois de 22% de tributos sobre vendas, a receita líquida é de R$ 9,5 milhões. Saímos do ano com ARR de quase R$ 17 milhões, base para 2028.");

  // 13 — DRE
  s = content("DRE 2027", `EBITDA de R$ ${mi(T.ebitda)} mi antes de COGS e lucro de R$ ${mi(T.result)} mi`);
  const Q = (a) => [...q(a), sum(a)];
  const ebitdaQ = Q(ebitda), grossQ = Q(revGross);
  const lines = [
    ["Receita bruta", Q(revGross), "b"],
    ["(−) Tributos sobre vendas (22%)", Q(revGross.map((v) => -v * 0.22))],
    ["Receita líquida", Q(revNet), "b"],
    ["(−) Vendas", Q(sales.map((v) => -v))],
    ["(−) Marketing (mídia, eventos e folha)", Q(mkt.map((v) => -v))],
    ["(−) P&D", Q(pd.map((v) => -v))],
    ["(−) G&A", Q(ga.map((v) => -v))],
    ["(−) COGS", null],
    ["EBITDA antes de COGS", ebitdaQ, "h"],
    ["Margem EBITDA (% receita bruta)", ebitdaQ.map((v, i) => v / grossQ[i]), "p"],
    ["(−) IRPJ + CSLL (22%)", Q(ebitda.map((v) => -v * 0.22))],
    ["Resultado", Q(ebitda.map((v) => v * 0.78)), "b"],
  ];
  const fm = (v) => (v < 0 ? "(" + mil(-v) + ")" : mil(v));
  const tRows = [[
    { text: "R$ mil", options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "left" } },
    ...["1T27", "2T27", "3T27", "4T27", "2027"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "right" } })),
  ]];
  lines.forEach(([label, vals, kind]) => {
    const hl = kind === "h";
    const fill = { color: hl ? HEX.tint : "FFFFFF" };
    const base = { bold: kind === "b" || hl, color: hl ? C.accent1 : C.text1, fill, border: [{ type: "none" }, { type: "none" }, { pt: 0.5, color: HEX.grid }, { type: "none" }] };
    tRows.push([
      { text: label, options: { ...base, align: "left" } },
      ...(vals ? vals.map((v, j) => ({ text: kind === "p" ? pct(v) : fm(v), options: { ...base, align: "right", bold: base.bold || j === 4 } }))
        : [{ text: "a definir — ver sensibilidade", options: { ...base, align: "right", italic: true, color: C.accent4, colspan: 5 } }]),
    ]);
  });
  s.addTable(tRows, { x: 0.6, y: 1.8, w: 12.1, colW: [4.6, 1.5, 1.5, 1.5, 1.5, 1.5], rowH: 0.34, fontSize: 13, margin: [0, 0.12, 0, 0.12], valign: "middle", objectName: on("dre") });
  note(s, "Inclui a folha do Growth Engineer e do Data Engineer (R$ 360 mil/ano cada, pós-encargos), previstas no plano mas fora do total da planilha.", 6.35);
  s.addNotes(`A DRE trimestral mostra a receita crescendo trimestre a trimestre e o EBITDA acompanhando. O segundo trimestre tem a menor margem porque concentra eventos. No ano, EBITDA antes de COGS de R$ ${mi(T.ebitda)} milhões, ${pct(T.ebitda / T.gross)} da receita bruta, e resultado de R$ ${mi(T.result)} milhões após IRPJ e CSLL. A linha de COGS ainda não foi estimada; o impacto aparece no slide de sensibilidade.`);

  // 14 — Custos e margem
  s = content("CUSTOS E MARGEM", `P&D é custo fixo: margem EBITDA sobe de ${pct(marginM[0])} para ${pct(marginM[11])}`);
  s.addChart(pres.charts.BAR, [
    { name: "P&D", labels: M, values: pd.map((v) => Math.round(v / 1000)) },
    { name: "Vendas", labels: M, values: sales.map((v) => Math.round(v / 1000)) },
    { name: "Marketing", labels: M, values: mkt.map((v) => Math.round(v / 1000)) },
    { name: "G&A", labels: M, values: ga.map((v) => Math.round(v / 1000)) },
  ], chartBase({ x: 0.6, y: 1.85, w: 6.4, h: 4.5, barDir: "col", barGrouping: "stacked", title: "Opex por mês (R$ mil)", chartColors: [HEX.dark, HEX.green, HEX.lime, HEX.soft], showLegend: true, legendPos: "b", barGapWidthPct: 40, valAxisMaxVal: 500, valAxisMajorUnit: 100 }));
  s.addChart(pres.charts.LINE, [{ name: "Margem EBITDA", labels: M, values: marginM.map((v) => Math.round(v * 1000) / 1000) }],
    chartBase({ x: 7.3, y: 1.85, w: 5.4, h: 4.5, title: "Margem EBITDA (% da receita bruta)", chartColors: [HEX.green], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: false, showValue: true, dataLabelPosition: "t", dataLabelFormatCode: "0%", valAxisLabelFormatCode: "0%", valAxisMinVal: 0, valAxisMaxVal: 0.6, valAxisMajorUnit: 0.1 }));
  note(s, `P&D responde por ${pct(T.pd / T.opex)} do opex. A queda de maio (${pct(marginM[4])}) vem dos R$ 150 mil em eventos do mês.`, 6.45);
  s.addNotes(`O opex vai de cerca de R$ ${mil(opex[0])} mil para R$ ${mil(opex[11])} mil por mês, enquanto a receita quase dobra. P&D é ${pct(T.pd / T.opex)} do custo e é fixo: não cresce com a base. Por isso a margem EBITDA sobe de ${pct(marginM[0])} em janeiro para ${pct(marginM[11])} em dezembro. A queda de maio é pontual e vem do maior evento do calendário.`);

  // 15 — Unit economics
  s = content("UNIT ECONOMICS", "Cada contrato paga seu CAC em cerca de 3,5 meses");
  [
    [`R$ ${(T.cac / 1000).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} mil`, "CAC", "Vendas + marketing ÷ 501 novos contratos"],
    [`R$ ${Math.round(T.ltv / 1000)} mil`, "LTV", "R$ 1.800 × margem de contribuição de 70% ÷ churn de 2%"],
    [`${Math.round(T.ltv / T.cac)}x`, "LTV / CAC", "Referência de mercado para SaaS B2B: > 3x"],
    [`${(T.cac / 1260).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} meses`, "Payback do CAC", "CAC ÷ margem de contribuição mensal (R$ 1.260)"],
  ].forEach((d, i) => stat(s, 0.6 + i * (sw + sg), 1.9, sw, 2.3, d[0], d[1], d[2], { bigSize: 28 }));
  card(s, 0.6, 4.55, 12.1, 1.75, C.text2);
  iconCircle(s, I.warn, 0.95, 4.9, 0.6, C.accent2);
  s.addText([
    { text: "Teste de estresse  ", options: { bold: true, color: C.accent2 } },
    { text: `Mesmo atribuindo todo o custo de vendas e marketing só às ${T.newDed} vendas da operação dedicada, o CAC sobe para R$ ${(T.cacDed / 1000).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} mil, o payback para ${(T.cacDed / 1260).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} meses e o LTV/CAC fica em ~${Math.round(T.ltv / T.cacDed)}x.`, options: { color: C.background1 } },
  ], { x: 1.8, y: 4.75, w: 10.6, h: 1.35, fontSize: 16, valign: "middle", margin: 0, isTextBox: true, objectName: on("t") });
  s.addNotes(`O CAC de R$ ${(T.cac / 1000).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} mil considera toda a folha de vendas, mídia, eventos e o Growth Engineer, dividido pelos 501 novos contratos. Com LTV de R$ 63 mil, a relação é de ${Math.round(T.ltv / T.cac)} vezes. Como a operação compartilhada não tem custo alocado no plano, fizemos o teste de estresse: mesmo jogando todo o custo só nas vendas dedicadas, a relação fica perto de ${Math.round(T.ltv / T.cacDed)} vezes. A margem de contribuição de 70% é premissa da planilha e precisa ser confirmada junto com o COGS.`);

  // 16 — Time
  s = content("TIME", "Time de 13 pessoas em jan/27, chegando a 19 em jul/27");
  [
    [I.code, "P&D", "9 pessoas", `R$ ${mi(T.pd, 2)} mi/ano`, "Tech Manager II, Associate PM, 3 devs (Jr, Pl, Especialista), QA Pleno, 2 Suporte Técnico (Jr, Pl) e Data Engineer Sênior"],
    [I.handshake, "Vendas", "3 → 6 → 9 pessoas", `R$ ${mi(T.sales, 2)} mi/ano`, "Sales Specialists dedicados a R$ 7 mil/mês, em três ondas: nov/26, abr/27 e jul/27"],
    [I.bull, "Marketing", "1 pessoa", `R$ ${mi(360000, 2)} mi/ano`, "Growth Engineer Pleno a partir de nov/26, operando mídia paga, eventos e funil"],
  ].forEach((d, i) => {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.9, 3.85, 3.95);
    iconCircle(s, d[0], x + 0.3, 2.15, 0.7);
    s.addText(d[1], { x: x + 1.15, y: 2.15, w: 2.5, h: 0.7, fontSize: 22, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[2], { x: x + 0.3, y: 3.05, w: 3.3, h: 0.5, fontSize: 24, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: on("t"), fontFace: THEME.headFontFace });
    s.addText(d[3], { x: x + 0.3, y: 3.55, w: 3.3, h: 0.4, fontSize: 14, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[4], { x: x + 0.3, y: 4.05, w: 3.3, h: 2.1, fontSize: 13, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: on("t") });
  });
  note(s, "Folha pós-encargos com premissa de encargos de 100% sobre o salário. G&A: R$ 40 mil/ano para dois offsites do time.", 6.1);
  s.addNotes("O time começa 2027 com 13 pessoas: 9 em P&D, 3 Sales Specialists e 1 Growth Engineer, e chega a 19 em julho com as duas ondas de vendas. A folha considera encargos de 100% sobre o salário, uma premissa conservadora que vale validar com RH. O G&A é pequeno: dois offsites por ano.");

  // 17 — Seção 3
  section("03", "Riscos e decisões", "Sensibilidade, riscos do plano e o que precisamos aprovar");

  // 18 — Sensibilidade (cascata)
  const ebitdaCogs = T.ebitda - T.cogsHyp;
  s = content("SENSIBILIDADE", `COGS é a incógnita: a 29,5%, o EBITDA cai para R$ ${mi(ebitdaCogs)} mi`);
  const wf = [
    ["EBITDA da planilha", ebitdaSheet, 0, "t"],
    ["Folha Growth Eng.", -360000, ebitdaSheet, "d"],
    ["Folha Data Eng.", -360000, ebitdaSheet - 360000, "d"],
    ["EBITDA antes de COGS", T.ebitda, 0, "t"],
    ["COGS (hipótese 29,5%)", -T.cogsHyp, T.ebitda, "d"],
    ["EBITDA com COGS", ebitdaCogs, 0, "t"],
  ];
  const baseY = 6.0, scale = 3.3 / 5.5e6, bw = 1.1, gap = 0.4, x0 = 0.9;
  s.addShape(pres.shapes.LINE, { x: 0.7, y: baseY, w: 8.9, h: 0, line: { color: HEX.grid, width: 1 }, objectName: on("eixo") });
  wf.forEach(([label, v, start, kind], i) => {
    const x = x0 + i * (bw + gap);
    const top = kind === "t" ? baseY - v * scale : baseY - start * scale;
    const h = Math.abs(v) * scale;
    const fill = kind === "t" ? (i === 5 ? C.accent4 : C.accent1) : C.accent6;
    s.addShape(pres.shapes.RECTANGLE, { x, y: top, w: bw, h, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: on("barra") });
    s.addText((v < 0 ? "−" : "") + mi(Math.abs(v), 2), { x: x - 0.2, y: top - 0.4, w: bw + 0.4, h: 0.35, fontSize: 14, bold: true, align: "center", color: kind === "t" ? C.text2 : C.accent6, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(label, { x: x - 0.1, y: baseY + 0.08, w: bw + 0.2, h: 0.45, fontSize: 11, align: "center", color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: on("t") });
  });
  s.addText("R$ mi, 2027", { x: 0.7, y: 1.85, w: 3, h: 0.3, fontSize: 11, color: C.accent3, margin: 0, isTextBox: true, objectName: on("t") });
  card(s, 10.1, 1.9, 2.6, 2.6);
  s.addText([
    { text: "Para fechar até dez/26", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "GCP por review", options: { bullet: true, breakLine: true } },
    { text: "New Relic e Cloudflare", options: { bullet: true, breakLine: true } },
    { text: "E-mail marketing (R$ 0,01/envio)", options: { bullet: true, breakLine: true } },
    { text: "Rateio de COGS do RA", options: { bullet: true } },
  ], { x: 10.35, y: 2.15, w: 2.15, h: 2.2, fontSize: 13, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: on("t") });
  s.addNotes(`Reconciliação do EBITDA: a planilha mostra R$ 5,33 milhões, mas deixa fora do total a folha do Growth Engineer e do Data Engineer, R$ 360 mil cada no ano. Com elas, R$ ${mi(T.ebitda, 2)} milhões antes de COGS. A linha de COGS está zerada; a planilha só anota uma referência de 29,5% de COGS geral do RA. Se aplicarmos 29,5% sobre a receita líquida, o EBITDA cai para R$ ${mi(ebitdaCogs, 2)} milhões, ${pct(ebitdaCogs / T.gross)} da receita bruta. Ainda positivo, mas precisamos fechar o COGS real antes da aprovação final.`);

  // 19 — Riscos
  s = content("RISCOS", "Quatro riscos principais e como mitigá-los");
  [
    [I.team, "Contratação e ramp-up", "9 Sales Specialists até jul/27, com 2 meses até produtividade plena. Um atraso de onda empurra vendas para o fim do ano, quando a sazonalidade é mais fraca.", "Abrir as vagas de abr e jul com 60 dias de antecedência"],
    [I.funnel, "Demanda descasada do calendário", "No 1T, o investimento gera ~620 leads contra ~1.180 necessários; eventos concentram 66% do budget entre abr e set.", "Antecipar mídia paga para o 1T e usar o pipeline de 2026"],
    [I.sync, "Churn de contratos", "O plano assume churn de 2% ao mês. A 3%, terminamos com ~713 contratos (−57) e ~R$ 0,5 mi a menos de receita.", "Meta de retenção por coorte e CS dedicado aos primeiros 90 dias"],
    [I.handshake, "Dependência da operação compartilhada", "32% das vendas vêm de um time sem custo alocado e com foco dividido; o Checkout encolhe 47% no ano.", "Meta e incentivo explícitos de Reviews para o time compartilhado"],
  ].forEach((d, i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 2.4;
    card(s, x, y, 5.95, 2.2);
    iconCircle(s, d[0], x + 0.25, y + 0.25, 0.55);
    s.addText(d[1], { x: x + 0.95, y: y + 0.25, w: 4.8, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[2], { x: x + 0.25, y: y + 0.9, w: 5.5, h: 0.8, fontSize: 12.5, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: on("t") });
    s.addText([{ text: "Mitigação: ", options: { bold: true } }, { text: d[3] }], { x: x + 0.25, y: y + 1.7, w: 5.5, h: 0.4, fontSize: 12.5, color: C.accent1, valign: "top", margin: 0, isTextBox: true, objectName: on("t") });
  });
  s.addNotes("Quatro riscos. Primeiro, a contratação: o plano depende de ter 9 Sales Specialists produtivos no segundo semestre. Segundo, o calendário de demanda: no primeiro trimestre o investimento gera cerca da metade dos leads necessários, então dependemos do pipeline de 2026. Terceiro, o churn: a 3% ao mês perdemos 57 contratos e cerca de meio milhão de receita. Quarto, um terço das vendas vem da operação compartilhada, que não tem custo nem meta específica no plano.");

  // 20 — Decisões
  s = content("PRÓXIMOS PASSOS", "O que precisamos aprovar", "FECHAMENTO");
  [
    ["1", "Budget 2027", `R$ ${mi(T.opex)} mi de opex para gerar R$ ${mi(T.gross)} mi de receita bruta e R$ ${mi(T.ebitda)} mi de EBITDA antes de COGS`],
    ["2", "Plano de contratação", "Nov/26: 3 Sales Specialists, Growth Engineer e Data Engineer · Abr/27: +3 SS · Jul/27: +3 SS"],
    ["3", "Calendário de marketing", "R$ 252 mil de mídia paga e R$ 480 mil em eventos entre abr e set, com CPL máximo de R$ 58"],
    ["4", "Fechar até dez/26", "COGS real (infra por review, e-mail), margem de contribuição e meta de churn por coorte"],
  ].forEach((d, i) => {
    const y = 1.95 + i * 1.18;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.75, h: 0.75, fill: { color: C.accent2 }, line: { color: C.accent2, width: 0 }, objectName: on("num") });
    s.addText(d[0], { x: 0.6, y, w: 0.75, h: 0.75, fontSize: 24, bold: true, align: "center", valign: "middle", color: C.text2, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[1], { x: 1.65, y: y - 0.02, w: 10.8, h: 0.42, fontSize: 20, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[2], { x: 1.65, y: y + 0.4, w: 10.8, h: 0.45, fontSize: 15, color: C.background2, margin: 0, isTextBox: true, objectName: on("t") });
  });
  s.addNotes("Pedimos quatro decisões: aprovar o budget, o plano de contratação em três ondas, o calendário de marketing, e o compromisso de fechar o COGS real e as metas de retenção até dezembro de 2026, antes do início do ano.");

  // 21 — Apêndice: premissas
  SECTION = "Apêndice";
  pres.addSection({ title: SECTION });
  s = content("APÊNDICE A", "Premissas do modelo");
  const prem = [
    ["Comercial", "Produtividade dedicado", "0,25 venda/dia útil (5/mês)"],
    ["Comercial", "Produtividade compartilhado", "0,1 venda/dia útil (2/mês)"],
    ["Comercial", "Ramp-up do vendedor novo", "M1 0% · M2 50% · M3 100%"],
    ["Comercial", "Ganho de produtividade", "M1 5% · M2 10% · M3 15% · M4 20%"],
    ["Comercial", "Sazonalidade", "jan e dez 25% · fev e nov 50% · demais 100%"],
    ["Funil", "Conversões", "Lead → qualificado 20% · qualificado → contrato 20%"],
    ["Funil", "CPL máximo", "R$ 58 (investimento ÷ leads necessários)"],
    ["Clientes", "Churn mensal", "Contratos 2,0% · Checkout 5,2%"],
    ["Receita", "ARPU mensal", "Contratos R$ 1.800 · Checkout R$ 76"],
    ["Receita", "Tributos", "22% sobre vendas · IRPJ + CSLL 22%"],
    ["Custos", "Encargos", "100% sobre o salário"],
  ];
  s.addTable([
    ["Tema", "Premissa", "Valor"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } })),
    ...prem.map((r, i) => r.map((t, j) => ({ text: t, options: { color: j === 0 ? C.accent1 : C.text1, bold: j === 0, fill: { color: i % 2 ? "FFFFFF" : HEX.tint } } }))),
  ], { x: 0.6, y: 1.8, w: 12.1, colW: [2.0, 4.3, 5.8], rowH: 0.38, fontSize: 13, border: { type: "none" }, margin: [0, 0.12, 0, 0.12], valign: "middle", objectName: on("premissas") });
  s.addNotes("Premissas como estão na aba SaaS Plan Template da planilha Blueprint @ Budget Reviews 2027.");

  // 22 — Apêndice: ajustes
  s = content("APÊNDICE B", "Ajustes feitos sobre a planilha para esta apresentação");
  const adj = [
    ["Folha fora do total", "Growth Engineer (Marketing) e Data Engineer (P&D) estão listados, mas não entram nas folhas totais nem no EBITDA. Incluídos aqui: −R$ 720 mil/ano."],
    ["COGS zerado", "GCP, New Relic, Cloudflare e e-mail sem valores; a nota de 29,5% de COGS do RA não foi aplicada. Tratado como sensibilidade."],
    ["Reviews de Checkout", "A fórmula usa a base total de clientes, não a de Checkout (209 mil vs. ~78 mil no ano), e os subtotais por modelo somam só parte dos meses. Corrigido: ~5,4 mi no ano."],
    ["Margens anuais", "As colunas de total de margem EBITDA (53%) e líquida (42%) repetem dezembro. Aqui: média do ano sobre a receita bruta."],
    ["Headcount", "A linha de headcount mostra 11 → 17 e omite Growth Engineer e Data Engineer. Aqui: 13 → 19."],
    ["Erros de referência", "#REF! em dez/27 (ARPAC, payback de contratos e aba Gráficos deck); a coluna de SQLs da aba Gráficos deck não bate com o modelo."],
  ];
  adj.forEach((d, i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 1.85 + Math.floor(i / 2) * 1.55;
    card(s, x, y, 5.95, 1.4);
    s.addText(d[0], { x: x + 0.25, y: y + 0.15, w: 5.5, h: 0.38, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: on("t") });
    s.addText(d[1], { x: x + 0.25, y: y + 0.53, w: 5.5, h: 0.8, fontSize: 12, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: on("t") });
  });
  s.addNotes("Lista dos ajustes aplicados sobre a planilha original, para facilitar a reconciliação dos números com o time de FP&A.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})();
