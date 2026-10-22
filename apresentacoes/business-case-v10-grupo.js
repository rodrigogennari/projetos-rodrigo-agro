const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const C = {
  navy: "1B3A8C", navyDark: "12295F", magenta: "C8106E", magLight: "FBE7F1",
  teal: "0096B4", tealDark: "26668F", aqua: "00A8A8", lime: "C5D300",
  text: "4A5560", muted: "7A858F", line: "D9DEE3", bg: "F4F6F8", white: "FFFFFF",
};
const FONT = "Calibri";
const OUT = process.argv[2] || "deck.pptx";

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function wave(alpha) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="260">
    <path d="M0,260 C500,215 900,60 1600,18 L1600,260 Z" fill="#C8106E" fill-opacity="${alpha}"/>
    <path d="M380,260 C850,232 1150,120 1600,92 L1600,260 Z" fill="#C8106E"/></svg>`;
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Rodrigo Gennari";
  pres.title = "Business Case v2.0 TN Manager and TN Manager Web";

  const WAVE = await wave(0.55);
  const I = {
    user: await icon(fa.FaUserTie, C.white), hat: await icon(fa.FaHardHat, C.white),
    tech: await icon(fa.FaUserMd, C.white), bldg: await icon(fa.FaBuilding, C.white),
    hand: await icon(fa.FaHandshake, C.white), search: await icon(fa.FaSearch, C.white),
    clip: await icon(fa.FaClipboardList, C.white), web: await icon(fa.FaLaptop, C.white),
    warn: await icon(fa.FaExclamationTriangle, C.magenta), check: await icon(fa.FaCheck, C.white),
    flag: await icon(fa.FaFlag, C.white),
  };

  let n = 0;
  function content(title, sub, source) {
    const s = pres.addSlide();
    n++;
    s.background = { color: C.white };
    s.addText(title, { x: 0.5, y: 0.28, w: 9, h: 0.5, fontFace: FONT, fontSize: 24, bold: true, color: C.navy, margin: 0 });
    if (sub) s.addText(sub, { x: 0.5, y: 0.78, w: 8.6, h: 0.32, fontFace: FONT, fontSize: 13, italic: true, color: C.navy, margin: 0 });
    s.addImage({ data: WAVE, x: 5.0, y: 4.81, w: 5.0, h: 0.8125 });
    s.addText(String(n), { x: 9.3, y: 5.28, w: 0.45, h: 0.25, fontFace: FONT, fontSize: 9, color: C.white, align: "right", margin: 0 });
    if (source) s.addText(source, { x: 0.5, y: 5.24, w: 5.2, h: 0.25, fontFace: FONT, fontSize: 8, color: C.muted, margin: 0 });
    return s;
  }
  const T = (s, text, o) => s.addText(text, Object.assign({ fontFace: FONT, color: C.text, margin: 0, valign: "top" }, o));
  const R = (s, o) => s.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { type: "none" } }, o));
  const RR = (s, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ line: { type: "none" }, rectRadius: 0.08 }, o));
  const OV = (s, o) => s.addShape(pres.shapes.OVAL, Object.assign({ line: { type: "none" } }, o));
  const LN = (s, x1, y1, x2, y2, color, width = 1.25, dash) => {
    s.addShape(pres.shapes.LINE, {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001,
      flipV: (x2 > x1) !== (y2 > y1) && y1 !== y2 && x1 !== x2,
      line: { color, width, dashType: dash || "solid" },
    });
  };
  const tag = (s, text, x, y, w, color = C.magenta) => {
    RR(s, { x, y, w, h: 0.24, fill: { color }, rectRadius: 0.12 });
    T(s, text, { x, y, w, h: 0.24, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle" });
  };

  const J = {
    mobile: await icon(fa.FaMobileAlt, C.white), paper: await icon(fa.FaFileAlt, C.white),
    cloud: await icon(fa.FaCloud, C.white), chart: await icon(fa.FaChartLine, C.white),
    user: await icon(fa.FaUserTie, C.white), wifi: await icon(fa.FaWifi, C.navy),
    pc: await icon(fa.FaDesktop, C.white), arrow: await icon(fa.FaArrowRight, C.muted),
  };
  const bullets = (items, o) => items.map((it, j) => ({ text: it, options: Object.assign({ bullet: { indent: 9 }, breakLine: j < items.length - 1, paraSpaceAfter: 3 }, o || {}) }));
  const head = (s, text, x, y, w, color, h = 0.36, size = 11) => {
    R(s, { x, y, w, h, fill: { color } });
    T(s, text, { x: x + 0.1, y, w: w - 0.2, h, fontSize: size, bold: true, color: C.white, valign: "middle" });
  };

  C.text = "33404C";
  const CAPS = ["", "CONTEXTO E PROPOSTA", "MERCADO", "SOLUÇÃO E DIFERENCIAL", "CONCORRÊNCIA", "MODELO DE NEGÓCIO", "PLANO E DECISÕES"];
  const chap = (s, c, dark) => T(s, "CAPÍTULO " + c + "  ·  " + CAPS[c], { x: 5.3, y: 0.07, w: 4.5, h: 0.2, fontSize: 8, bold: true, color: dark ? C.lime : C.magenta, align: "right", charSpacing: 1.5 });
  const card = (s, x, y, w, h, col, title, text, o) => {
    o = o || {};
    R(s, { x, y, w, h, fill: { color: o.fill || C.bg } });
    R(s, { x, y, w, h: 0.06, fill: { color: col } });
    T(s, title, { x: x + 0.12, y: y + 0.14, w: w - 0.24, h: o.th || 0.3, fontSize: o.ts || 13, bold: true, color: col });
    if (text) T(s, text, { x: x + 0.12, y: y + 0.14 + (o.th || 0.3) + 0.04, w: w - 0.24, h: h - (o.th || 0.3) - 0.24, fontSize: o.fs || 11 });
  };

  const IMG = __dirname + "/../img/";
  const K = {
    chart: await icon(fa.FaChartBar, C.white), search: await icon(fa.FaSearchPlus, C.white),
    bench: await icon(fa.FaBalanceScale, C.white), trophy: await icon(fa.FaTrophy, C.white),
    star: await icon(fa.FaStar, C.white), medal: await icon(fa.FaMedal, C.white),
  };
  const KC = {
    search: await icon(fa.FaSearchPlus, C.tealDark), chart: await icon(fa.FaChartBar, C.navy),
    bench: await icon(fa.FaBalanceScale, C.navy), trophy: await icon(fa.FaTrophy, C.magenta),
  };

  // G1 Grupo de Desenvolvimento
  {
    const s = content("Grupo de Desenvolvimento: comunidade movida a dados", "Análise, benchmarking e premiação engajam as fazendas a participar e os técnicos a trazer mais fazendas.");
    chap(s, 5);
    [["Análise de dados", "Analistas online auditam, validam e interpretam os lançamentos", KC.search, C.tealDark],
     ["Indicadores", "Dashboards semanais e relatório trimestral com comparações", KC.chart, C.navy],
     ["Benchmarking", "Minha fazenda x meta x média do grupo x melhores do grupo", KC.bench, C.navy],
     ["Premiação", "Melhores fazendas e melhores técnicos no evento anual", KC.trophy, C.magenta]].forEach((p, i) => {
      const x = 0.5 + i * 2.27, w = 2.17;
      R(s, { x, y: 1.2, w, h: 1.62, fill: { color: p[3] } });
      OV(s, { x: x + 0.14, y: 1.32, w: 0.46, h: 0.46, fill: { color: C.white } });
      s.addImage({ data: p[2], x: x + 0.24, y: 1.42, w: 0.26, h: 0.26 });
      T(s, p[0], { x: x + 0.7, y: 1.32, w: w - 0.8, h: 0.46, fontSize: 13.5, bold: true, color: C.white, valign: "middle" });
      T(s, p[1], { x: x + 0.14, y: 1.9, w: w - 0.26, h: 0.86, fontSize: 11, bold: true, color: C.white });
    });
    [["Para a fazenda", C.navy, ["Saber onde está em relação ao grupo", "Evolução medida safra a safra", "Reconhecimento no evento"]],
     ["Para o técnico TN", C.magenta, ["Prêmio de técnico destaque", "Conta: fazendas ativas e lançamentos em dia", "Mais fazendas no grupo, mais resultado"]],
     ["Para a TN", C.tealDark, ["Fidelização pelo grupo, não pelo sistema", "Base de dados própria para benchmark", "Novas fazendas querendo entrar"]]].forEach((b, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      head(s, b[0], x, 2.98, w, b[1], 0.34, 12);
      R(s, { x, y: 3.32, w, h: 1.0, fill: { color: C.bg } });
      T(s, bullets(b[2], { paraSpaceAfter: 2 }), { x: x + 0.1, y: 3.38, w: w - 0.2, h: 0.92, fontSize: 10.5, bold: true });
    });
    R(s, { x: 0.5, y: 4.42, w: 9.0, h: 0.36, fill: { color: C.navy } });
    T(s, [{ text: "Ciclo: ", options: { color: C.lime } }, { text: "mais fazendas, mais dados, benchmark mais forte, mais valor, mais fazendas" }], { x: 0.68, y: 4.42, w: 8.7, h: 0.36, fontSize: 12, bold: true, color: C.white, valign: "middle" });
    s.addNotes("Fala: 'O Grupo de Desenvolvimento é uma comunidade movida a dados. Os analistas online auditam, validam e interpretam os lançamentos. Cada fazenda recebe indicadores e se compara com a meta, com a média do grupo e com as melhores. No fim da safra, premiamos as melhores fazendas e também os melhores técnicos. O prêmio do técnico é o que faz o consultor trazer mais fazendas e manter os lançamentos em dia. Quanto mais fazendas, mais forte o benchmark, e mais valor para todo mundo.' Inspirado no Fazenda Nota 10, que o Rodrigo criou em 2020. Ponte: 'O que cada fazenda recebe ao longo da safra?'");
  }

  // G2 Entregas
  {
    const s = content("Entregas do Grupo: do trimestre à Reportagem de Safra", "Análises ao longo da safra e um fechamento completo no final, com comparações.", "Valores ilustrativos, apenas para mostrar o formato das entregas.");
    chap(s, 5);
    // capa da Reportagem de Safra
    s.addImage({ path: IMG + "gado.jpg", x: 0.5, y: 1.2, w: 2.4, h: 1.6, sizing: { type: "cover", w: 2.4, h: 1.6 } });
    R(s, { x: 0.5, y: 2.8, w: 2.4, h: 1.95, fill: { color: C.navy } });
    R(s, { x: 0.5, y: 2.8, w: 2.4, h: 0.06, fill: { color: C.lime } });
    T(s, "REPORTAGEM DE SAFRA", { x: 0.65, y: 2.95, w: 2.15, h: 0.24, fontSize: 9.5, bold: true, color: C.lime, charSpacing: 2 });
    T(s, "Desempenho anual da fazenda", { x: 0.65, y: 3.2, w: 2.15, h: 0.6, fontSize: 15, bold: true, color: C.white });
    T(s, "Safra 2027/2028", { x: 0.65, y: 3.85, w: 2.15, h: 0.28, fontSize: 12, bold: true, color: C.white });
    T(s, "Grupo de Desenvolvimento TN", { x: 0.65, y: 4.15, w: 2.15, h: 0.24, fontSize: 9.5, color: "C9D3EA" });
    T(s, "Ranking, indicadores e evolução safra a safra", { x: 0.65, y: 4.38, w: 2.15, h: 0.32, fontSize: 9, italic: true, color: "C9D3EA" });
    // gráfico de benchmark
    R(s, { x: 3.1, y: 1.2, w: 3.45, h: 1.95, fill: { color: C.white }, line: { color: C.line, width: 1 } });
    T(s, "GMD no confinamento (kg/dia)", { x: 3.22, y: 1.26, w: 3.2, h: 0.26, fontSize: 11, bold: true, color: C.navy });
    const bars = [["Minha fazenda", 1.62, "E07A1F"], ["Meta", 1.50, C.teal], ["Média do grupo", 1.48, "1E9A5A"], ["Top rentáveis", 1.55, C.navy], ["Top indicadores", 1.70, "7C5CD6"]];
    bars.forEach((b, i) => {
      const y = 1.58 + i * 0.3, w = 1.55 * b[1] / 1.7;
      T(s, b[0], { x: 3.18, y, w: 1.1, h: 0.24, fontSize: 9, bold: true, color: C.text, align: "right", valign: "middle" });
      R(s, { x: 4.33, y: y + 0.03, w, h: 0.18, fill: { color: b[2] } });
      RR(s, { x: 4.33 + w + 0.04, y: y + 0.01, w: 0.5, h: 0.22, fill: { color: b[2] }, rectRadius: 0.11 });
      T(s, b[1].toFixed(2).replace(".", ","), { x: 4.33 + w + 0.04, y: y + 0.01, w: 0.5, h: 0.22, fontSize: 9, bold: true, color: C.white, align: "center", valign: "middle" });
    });
    // evolução trimestral
    R(s, { x: 3.1, y: 3.25, w: 3.45, h: 1.5, fill: { color: C.white }, line: { color: C.line, width: 1 } });
    T(s, "Custo da @ produzida por trimestre (R$)", { x: 3.22, y: 3.3, w: 3.2, h: 0.26, fontSize: 11, bold: true, color: C.navy });
    const tri = [["T1", 268, 262], ["T2", 255, 260], ["T3", 246, 258], ["T4", 238, 255]];
    const base = 4.45, sc = 0.55 / 270;
    tri.forEach((t, i) => {
      const x = 3.45 + i * 0.78;
      const h1 = t[1] * sc, h2 = t[2] * sc;
      R(s, { x, y: base - h1, w: 0.26, h: h1, fill: { color: "E07A1F" } });
      R(s, { x: x + 0.28, y: base - h2, w: 0.26, h: h2, fill: { color: "1E9A5A" } });
      T(s, String(t[1]), { x: x - 0.06, y: base - h1 - 0.18, w: 0.38, h: 0.16, fontSize: 7.5, bold: true, color: "B85E10", align: "center" });
      T(s, String(t[2]), { x: x + 0.22, y: base - h2 - 0.18, w: 0.38, h: 0.16, fontSize: 7.5, bold: true, color: "1E7A48", align: "center" });
      T(s, t[0], { x, y: base + 0.02, w: 0.54, h: 0.18, fontSize: 9, bold: true, color: C.text, align: "center" });
    });
    R(s, { x: 3.24, y: 3.6, w: 0.12, h: 0.12, fill: { color: "E07A1F" } });
    T(s, "Minha fazenda", { x: 3.4, y: 3.57, w: 1.0, h: 0.18, fontSize: 8, bold: true });
    R(s, { x: 4.45, y: 3.6, w: 0.12, h: 0.12, fill: { color: "1E9A5A" } });
    T(s, "Média do grupo", { x: 4.61, y: 3.57, w: 1.1, h: 0.18, fontSize: 8, bold: true });
    // calendário de entregas
    T(s, "Calendário de entregas", { x: 6.8, y: 1.18, w: 2.7, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    [["Semanal", "Dashboard e análise do analista", C.tealDark], ["Trimestral", "Relatório com comparação entre trimestres e com o grupo", C.navy], ["Fim da safra", "Reportagem de Safra: fechamento, ranking e evolução", C.magenta], ["Anual", "Evento com análise dos dados e premiação", C.magenta]].forEach((r, i) => {
      const y = 1.52 + i * 0.66;
      R(s, { x: 6.8, y, w: 2.7, h: 0.6, fill: { color: C.bg } });
      R(s, { x: 6.8, y, w: 0.06, h: 0.6, fill: { color: r[2] } });
      T(s, r[0], { x: 6.95, y: y + 0.04, w: 2.5, h: 0.22, fontSize: 11, bold: true, color: r[2] });
      T(s, r[1], { x: 6.95, y: y + 0.26, w: 2.5, h: 0.32, fontSize: 9, bold: true, color: C.text });
    });
    T(s, "Indicadores: GMD, conversão alimentar, custo da diária, custo da @ produzida, rendimento de carcaça, dias de cocho e margem por cabeça.", { x: 6.8, y: 4.18, w: 2.7, h: 0.6, fontSize: 8.5, italic: true, color: C.muted });
    s.addNotes("Fala: 'Toda semana a fazenda tem o dashboard e a análise do analista. A cada trimestre, um relatório compara o trimestre atual com os anteriores e com o grupo. No fim da safra, a Reportagem de Safra fecha o ano: indicadores, ranking e evolução safra a safra. É esse fechamento que alimenta o evento anual.' Os números do slide são ilustrativos, só para mostrar o formato. As comparações seguem o padrão: minha fazenda, meta, média do grupo, top rentáveis (as mais lucrativas) e top indicadores (as melhores em desempenho técnico). Ponte: 'E uma vez por ano, o grupo se encontra.'");
  }

  // G3 Evento anual
  {
    const s = content("Evento anual: análise dos dados e premiação", "O encontro do grupo: resultados da safra, troca entre confinadores e reconhecimento.", "Imagem ilustrativa de um evento de análise de resultados.");
    chap(s, 5);
    s.addImage({ path: IMG + "evento.jpg", x: 0.5, y: 1.2, w: 5.25, h: 3.44, sizing: { type: "cover", w: 5.25, h: 3.44 } });
    R(s, { x: 0.5, y: 4.3, w: 5.25, h: 0.34, fill: { color: C.navy, transparency: 15 } });
    T(s, "Resultados da safra apresentados para o grupo", { x: 0.65, y: 4.3, w: 5.0, h: 0.34, fontSize: 11, bold: true, color: C.white, valign: "middle" });
    T(s, "Categorias de premiação", { x: 6.0, y: 1.18, w: 3.5, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    [["Fazenda mais rentável", "Maior margem por cabeça na safra", K.trophy, C.magenta],
     ["Maior evolução", "Quem mais melhorou de uma safra para outra", K.star, C.navy],
     ["Melhores indicadores", "GMD, conversão e custo da @ produzida", K.medal, C.tealDark],
     ["Técnico destaque", "Mais fazendas ativas e lançamentos em dia", K.medal, C.magenta]].forEach((c, i) => {
      const y = 1.52 + i * 0.66;
      R(s, { x: 6.0, y, w: 3.5, h: 0.6, fill: { color: C.bg } });
      OV(s, { x: 6.1, y: y + 0.09, w: 0.42, h: 0.42, fill: { color: c[3] } });
      s.addImage({ data: c[2], x: 6.2, y: y + 0.19, w: 0.22, h: 0.22 });
      T(s, c[0], { x: 6.62, y: y + 0.04, w: 2.85, h: 0.26, fontSize: 12, bold: true, color: c[3] });
      T(s, c[1], { x: 6.62, y: y + 0.3, w: 2.85, h: 0.26, fontSize: 9.5, bold: true, color: C.text });
    });
    R(s, { x: 6.0, y: 4.22, w: 3.5, h: 0.56, fill: { color: C.magLight } });
    T(s, "Quem vê o reconhecimento quer participar: novas fazendas e novos técnicos entram no grupo.", { x: 6.1, y: 4.22, w: 3.35, h: 0.56, fontSize: 10, bold: true, color: C.magenta, valign: "middle" });
    s.addNotes("Fala: 'Uma vez por ano, o grupo se encontra. Apresentamos os resultados da safra, as fazendas trocam experiência e premiamos as melhores: a mais rentável, a que mais evoluiu, a de melhores indicadores e o técnico destaque. O reconhecimento é o que traz novas fazendas e novos técnicos para o grupo, e é o que fideliza.' A foto é ilustrativa; os logotipos foram removidos. Ponte: 'Como isso entra no calendário?'");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
