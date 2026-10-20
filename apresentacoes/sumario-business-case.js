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
  pres.title = "Sumário Business Case TN Manager Web";

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

  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    T(s, "Sumário", { x: 0.5, y: 0.28, w: 9, h: 0.5, fontSize: 24, bold: true, color: C.navy });
    T(s, "Business Case TN Manager Web: o que vamos ver hoje", { x: 0.5, y: 0.78, w: 8.6, h: 0.32, fontSize: 13, italic: true, color: C.navy });
    s.addImage({ data: WAVE, x: 5.0, y: 4.81, w: 5.0, h: 0.8125 });
    const items = [
      ["Resumo Executivo", "A recomendação em uma página"],
      ["Aderência Estratégica", "Como o projeto atende as prioridades da Trouw"],
      ["Panorama do confinamento no Brasil", "Dados de 2025 e onde está o nosso público"],
      ["Quem é nosso Target Customer", "Confinamentos de 200 a 2.000 cabeças"],
      ["O que muda da v1.0 para a v2.0", "De 2023 para o novo modelo"],
      ["Proposta de Valor", "O que o confinador ganha com o TN Manager Web"],
      ["Descrição do MVP", "O que entra na primeira versão"],
      ["Roadmap Alvo", "Como o produto evolui por etapas"],
      ["Go-to-Market", "Como chegamos às fazendas"],
      ["Próximos passos e decisões", "O que precisamos aprovar hoje"],
    ];
    items.forEach((it, i) => {
      const col = Math.floor(i / 5), row = i % 5;
      const x = 0.6 + col * 4.55, y = 1.35 + row * 0.68;
      OV(s, { x, y, w: 0.46, h: 0.46, fill: { color: C.magenta } });
      T(s, String(i + 1), { x, y, w: 0.46, h: 0.46, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, it[0], { x: x + 0.62, y: y - 0.02, w: 3.75, h: 0.27, fontSize: 13, bold: true, color: C.navy });
      T(s, it[1], { x: x + 0.62, y: y + 0.25, w: 3.75, h: 0.24, fontSize: 10, color: C.text });
    });
    LN(s, 5.0, 1.35, 5.0, 4.45, C.line, 0.75);
  }
  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
