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
    const s = content("Quem atende quem: três frentes, sem conflito", "Web para novos clientes pequenos e médios. Quem precisa de automação segue com a PecCode.");
    const cards = [
      ["AGORA", "TN Manager Web", C.navy, [["Para quem", "Novos clientes: confinamentos pequenos e médios, de 200 a 2.000 cabeças"], ["Como", "Lançamento manual, sem automação no momento, com analista TN"], ["Papel", "Abrir mercado e fidelizar"]]],
      ["CONTINUA", "PecCode (parceria)", C.tealDark, [["Para quem", "Clientes maiores que querem automação"], ["Como", "Balança, vagão e integrações. Parceria consolidada e rodando bem"], ["Papel", "Atender hoje quem precisa de automação"]]],
      ["FUTURO", "TN Manager Automação", C.magenta, [["Para quem", "Clientes maiores e fazendas Web que crescerem"], ["Como", "Retomado no futuro, trabalhando em paralelo com a versão Web"], ["Papel", "Completar a oferta digital da TN"]]],
    ];
    cards.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.88, y = 1.25;
      R(s, { x, y, w, h: 2.7, fill: { color: C.bg } });
      R(s, { x, y, w, h: 0.62, fill: { color: c[2] } });
      T(s, c[0], { x: x + 0.14, y: y + 0.06, w: w - 0.2, h: 0.2, fontSize: 8.5, bold: true, color: i === 0 ? C.lime : C.white, charSpacing: 2 });
      T(s, c[1], { x: x + 0.14, y: y + 0.27, w: w - 0.2, h: 0.3, fontSize: 14, bold: true, color: C.white });
      c[3].forEach((r, k) => {
        const yy = y + 0.74 + k * 0.64;
        T(s, r[0], { x: x + 0.14, y: yy, w: w - 0.26, h: 0.2, fontSize: 9.5, bold: true, color: c[2] });
        T(s, r[1], { x: x + 0.14, y: yy + 0.2, w: w - 0.26, h: 0.4, fontSize: 9.5 });
      });
    });
    R(s, { x: 0.5, y: 4.1, w: 9.0, h: 0.62, fill: { color: C.navy } });
    T(s, [{ text: "Regra simples: ", options: { bold: true, color: C.lime } }, { text: "sem automação, TN Manager Web. Com automação, PecCode hoje e TN Manager Automação no futuro, lado a lado com a versão Web." }], { x: 0.68, y: 4.1, w: 8.7, h: 0.62, fontSize: 11.5, color: C.white, valign: "middle" });
    s.addNotes("Fala sugerida: 'Para não haver dúvida: o TN Manager Web é para novos clientes, de porte pequeno e médio, e sem automação no momento. Os clientes maiores, que querem automação, continuam com a PecCode: a parceria está consolidada e rodando bem, e vamos mantê-la. O TN Manager com automação será retomado no futuro, trabalhando em paralelo com a versão Web.' Ponte: 'Em resumo, é isso que propomos.'");
  }
  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
