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

  {
    const s = content("Onde vamos vencer: suporte e inteligência de dados", "O sistema é o meio. O que fazemos com os dados da fazenda é o diferencial.", "Fontes: avaliação interna TN e reclamações de clientes sobre o suporte dos fornecedores atuais (validar no discovery); licenças pagas pela TN; Embrapa (2018): sistemas vistos como complexos e caros.");
    head(s, "Onde os concorrentes falham", 0.5, 1.18, 3.3, "7A858F", 0.32, 10.5);
    head(s, "Onde a TN vai vencer", 4.2, 1.18, 5.3, C.magenta, 0.32, 10.5);
    const rows = [
      ["Suporte fraco", "Depois da venda, o cliente fica sozinho com o sistema", "Suporte excelente", "Analista online que responde rápido, técnico TN no campo e acompanhamento toda semana"],
      ["Implantação cara e longa", "Licença de R$ 10 mil a 40 mil por ano e projeto pesado para começar", "Implantação simples", "Começa em dias, sem licença para o cliente e com lançamento manual"],
      ["Dado parado no sistema", "A ferramenta é entregue, mas ninguém transforma o número em decisão", "Dashboards criativos e análise", "Painéis inovadores para decidir, lidos pelo analista junto com o gestor"],
    ];
    rows.forEach((r, i) => {
      const y = 1.58 + i * 0.76;
      R(s, { x: 0.5, y, w: 3.3, h: 0.68, fill: { color: "E3E7EC" } });
      T(s, r[0], { x: 0.62, y: y + 0.07, w: 3.1, h: 0.22, fontSize: 10.5, bold: true, color: C.muted });
      T(s, r[1], { x: 0.62, y: y + 0.3, w: 3.1, h: 0.36, fontSize: 9, color: C.muted });
      s.addImage({ data: J.arrow, x: 3.9, y: y + 0.24, w: 0.2, h: 0.2 });
      R(s, { x: 4.2, y, w: 5.3, h: 0.68, fill: { color: C.magLight } });
      R(s, { x: 4.2, y, w: 0.06, h: 0.68, fill: { color: C.magenta } });
      T(s, r[2], { x: 4.38, y: y + 0.07, w: 5.0, h: 0.22, fontSize: 11, bold: true, color: C.magenta });
      T(s, r[3], { x: 4.38, y: y + 0.3, w: 5.0, h: 0.36, fontSize: 9.5, color: C.text });
    });
    T(s, "GRUPO DE DESENVOLVIMENTO: DO SISTEMA AO RESULTADO", { x: 0.5, y: 3.92, w: 9, h: 0.2, fontSize: 8.5, bold: true, color: C.navy, charSpacing: 2 });
    const st = [["Sistema", "o meio", "9AA3AC"], ["Dados da fazenda", "lançados no dia a dia", C.tealDark], ["Analistas", "auditam, validam e interpretam", C.navy], ["Relatório de safra", "e benchmark: o resultado", C.magenta]];
    const w = 9.0 / 4;
    st.forEach((p, i) => {
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: 0.5 + i * w, y: 4.17, w: w + 0.12, h: 0.6, fill: { color: p[2] }, line: { color: C.white, width: 2 } });
      T(s, [{ text: p[0], options: { bold: true, breakLine: true } }, { text: p[1], options: { fontSize: 8.5 } }], { x: 0.5 + i * w + (i === 0 ? 0.12 : 0.36), y: 4.17, w: w - 0.5, h: 0.6, fontSize: 10, color: C.white, align: "center", valign: "middle" });
    });
    s.addNotes("Fala sugerida: 'Todo concorrente tem sistema. Onde eles falham é no que vem depois: o suporte é fraco, a implantação é cara e demorada, e o dado fica parado. É exatamente aí que vamos vencer: suporte excelente, implantação simples e dashboards criativos, lidos por um analista junto com o gestor. No longo prazo, o Grupo de Desenvolvimento fecha o ciclo: os analistas auditam, validam e interpretam os dados de cada fazenda e entregam o relatório de safra e o benchmark. O sistema é o meio; o que fazemos com os dados é o diferencial.' Cuidado: 'suporte fraco' vem da avaliação interna e das reclamações dos nossos clientes (o slide de objetivos já cita o suporte fraco do TGC); confirmar no discovery com perguntas sobre a relação com o fornecedor. Ponte: 'Quanto custa entregar isso e como cobramos?'");
  }
  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
