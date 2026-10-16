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
  pres.title = "Panorama do Confinamento no Brasil";

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


  const FONTES = "Fontes: dsm-firmenich, Censo de Confinamento 2025 (fev/2026) e 2024; Embrapa Pecuária Sudeste (2023; 2024); Embrapa, Censos Agropecuários 2006 e 2017.";

  // Slide 1: raio-X
  {
    const s = content("Raio-X do confinamento no Brasil (2025)", "O mercado cresce rápido e se concentra nos grandes, mas a maioria das fazendas ainda é pequena.", FONTES);
    const kpis = [["9,25 mi", "cabeças confinadas em 2025, alta de 16% sobre 2024", C.navy], ["2.445", "propriedades mapeadas em 1.095 municípios", C.tealDark], ["45%", "das cabeças estão nos 100 maiores confinamentos", C.magenta]];
    kpis.forEach((k, i) => {
      const y = 1.3 + i * 1.12;
      R(s, { x: 0.5, y, w: 2.2, h: 1.0, fill: { color: C.bg } });
      R(s, { x: 0.5, y, w: 0.06, h: 1.0, fill: { color: k[2] } });
      T(s, k[0], { x: 0.68, y: y + 0.08, w: 1.95, h: 0.42, fontSize: 22, bold: true, color: k[2] });
      T(s, k[1], { x: 0.68, y: y + 0.5, w: 1.95, h: 0.48, fontSize: 9 });
    });
    // pirâmide por número de propriedades
    T(s, "Propriedades por porte", { x: 2.95, y: 1.22, w: 3.4, h: 0.25, fontSize: 11, bold: true, color: C.navy, align: "center" });
    const cx = 4.65;
    const lv = [
      [1.5, C.navy, "100 maiores", "45% das cabeças"],
      [2.5, C.tealDark, "cerca de 1.200", "1.000 cabeças ou mais"],
      [3.4, C.magenta, "1.131 propriedades", "menos de 1.000 cabeças (46%)"],
    ];
    lv.forEach((l, i) => {
      const y = 1.52 + i * 0.86;
      R(s, { x: cx - l[0] / 2, y, w: l[0], h: 0.8, fill: { color: l[1] } });
      T(s, l[2], { x: cx - l[0] / 2, y: y + 0.1, w: l[0], h: 0.3, fontSize: 12, bold: true, color: C.white, align: "center" });
      T(s, l[3], { x: cx - l[0] / 2, y: y + 0.42, w: l[0], h: 0.28, fontSize: 9, color: C.white, align: "center" });
    });
    RR(s, { x: cx - 1.7, y: 4.18, w: 3.4, h: 0.3, fill: { color: C.lime }, rectRadius: 0.15 });
    T(s, "Faixa do TN Manager Web: 200 a 2.000 cabeças", { x: cx - 1.7, y: 4.18, w: 3.4, h: 0.3, fontSize: 9.5, bold: true, color: C.navy, align: "center", valign: "middle" });
    // estados
    T(s, "Cabeças confinadas por estado (milhões)", { x: 6.6, y: 1.22, w: 3.0, h: 0.25, fontSize: 11, bold: true, color: C.navy });
    [["MT", 2.2, C.magenta], ["SP", 1.4, C.navy], ["GO", 1.4, C.navy], ["MS", 0.9, C.navy], ["MG", 0.8, C.navy], ["Outros", 2.55, C.teal]].forEach((b, i) => {
      const y = 1.55 + i * 0.4, maxW = 1.9;
      T(s, b[0], { x: 6.6, y, w: 0.6, h: 0.3, fontSize: 10, valign: "middle" });
      R(s, { x: 7.2, y: y + 0.03, w: maxW * b[1] / 2.6, h: 0.24, fill: { color: b[2] } });
      T(s, b[1].toFixed(b[1] === 2.55 ? 2 : 1).replace(".", ","), { x: 7.25 + maxW * b[1] / 2.6, y, w: 0.5, h: 0.3, fontSize: 9.5, bold: true, color: C.text, valign: "middle" });
    });
    T(s, "5 estados somam 72%. Centro-Oeste e o bioma Cerrado concentram a maior parte.", { x: 6.6, y: 4.02, w: 3.0, h: 0.45, fontSize: 9, italic: true, color: C.text });
    s.addNotes("Fonte principal: Censo de Confinamento 2025 da dsm-firmenich, divulgado em fevereiro de 2026 (9,25 milhões de cabeças, 2.445 propriedades, 1.095 municípios, 100 maiores com 45% das cabeças; MT 2,2 mi, SP 1,4, GO 1,4, MS 0,9, MG 0,8). 'Outros' = 9,25 menos a soma dos 5 estados. O número de 1.131 propriedades abaixo de 1.000 cabeças vem de reportagem sobre o mesmo censo (CompreRural): confirmar no relatório completo. A faixa do meio (cerca de 1.200; exatamente 1.214) é calculada: 2.445 menos 1.131 menos 100. O censo mapeia propriedades monitoradas, não é a contagem de todos os confinamentos do país. Bioma: não há tabela pública por bioma; o Cerrado é inferência pela concentração no Centro-Oeste, MG e SP. Embrapa (Censos 2006 e 2017): o número de estabelecimentos com confinamento cresceu cerca de 145% e aumentou em todos os biomas.");
  }

  // Slide 2: veredito
  {
    const s = content("Estamos no caminho certo?", "Sim, com uma ressalva: o público é o maior em número de fazendas, mas pequeno em volume de cabeças.", FONTES);
    const cols = [
      { t: "O que os dados confirmam", col: C.navy, items: [
        "Confinamentos com menos de 1.000 cabeças são o maior grupo: 1.131 propriedades, cerca de 46% do total",
        "Os grandes já estão concentrados e disputados: os 100 maiores têm 45% das cabeças",
        "A adoção digital cresce com escala e consultoria (Embrapa): o pequeno fica sem ferramenta, e o técnico TN é o elo",
        "A Embrapa indica que o vagão forrageiro se justifica a partir de cerca de 200 cabeças: coerente com o nosso piso",
      ] },
      { t: "O que pede cuidado", col: C.magenta, items: [
        "Em volume, o pequeno pesa pouco: estimamos menos de 10% das cabeças confinadas",
        "Abaixo de 10 mil cabeças o volume ficou estável em 2024 (3,3 mi), enquanto os grandes cresceram",
        "O valor vem do número de fazendas e da fidelização, não do Sales Profit por cliente",
      ] },
      { t: "O que ainda não sabemos", col: C.tealDark, items: [
        "Percentual nacional de confinamentos com automação: não há dado público",
        "Número total de confinamentos do país: o censo mapeia 2.445; o IBGE (Censo Agro 2017) é a única contagem universal",
        "Recorte de 1.000 a 2.000 cabeças e distribuição por bioma",
        "A pesquisa com 30 fazendas preenche essas lacunas",
      ] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.85;
      R(s, { x, y: 1.25, w, h: 0.42, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.12, y: 1.25, w: w - 0.24, h: 0.42, fontSize: 12, bold: true, color: C.white, valign: "middle" });
      R(s, { x, y: 1.67, w, h: 2.33, fill: { color: C.bg } });
      T(s, c.items.map((it, j) => ({ text: it, options: { bullet: { indent: 10 }, breakLine: j < c.items.length - 1, paraSpaceAfter: 4 } })), { x: x + 0.08, y: 1.77, w: w - 0.18, h: 2.18, fontSize: 9.5 });
    });
    R(s, { x: 0.5, y: 4.1, w: 9.0, h: 0.58, fill: { color: C.navy } });
    T(s, [{ text: "Conclusão: ", options: { bold: true, color: C.lime } }, { text: "o público de 200 a 2.000 cabeças é o maior em número de fazendas e o menos atendido. O caso se sustenta com custo baixo por cliente e escala de fazendas." }], { x: 0.68, y: 4.1, w: 8.7, h: 0.58, fontSize: 11, color: C.white, valign: "middle" });
    s.addNotes("Estimativa de 'menos de 10% das cabeças': 1.131 propriedades com média hipotética de 500 cabeças dão cerca de 0,57 milhão, ou 6% de 9,25 milhões. É uma conta nossa, não dado do censo. Dados de 2024 (dsm-firmenich): confinamentos acima de 10 mil cabeças passaram de 53,9% para 58,7% das cabeças; abaixo de 10 mil ficaram em 3,3 milhões. Embrapa Pecuária Sudeste: estudos com 184 confinamentos (2023) e 130 confinamentos (2024) mostram que escala, escolaridade e consultoria aumentam a adoção de tecnologia digital; adotantes de software têm eficiência técnica maior (114 confinamentos). A indicação do vagão a partir de cerca de 200 cabeças vem de publicação técnica antiga da Embrapa sobre confinamento de bovinos.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
