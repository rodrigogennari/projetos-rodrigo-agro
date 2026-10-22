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

  // N1 Agenda
  {
    const s = content("Agenda", "Seis blocos, do contexto às decisões de hoje");
    const items = [
      ["Contexto e proposta", "Resumo, nossa história e o cliente-alvo"],
      ["Mercado", "Bússola Farm, nossa base de clientes e a nova geração"],
      ["Solução e diferencial", "Suporte, dashboards e equipe técnica: como funciona na prática"],
      ["Concorrência", "Cargill, DSM, Ponta e Tecnobeef: onde vamos vencer"],
      ["Modelo de negócio", "Custos, SWOT, go-to-market e Grupo de Desenvolvimento"],
      ["Plano e decisões", "Roadmap, cronograma de 4 anos e o que aprovar hoje"],
    ];
    items.forEach((it, i) => {
      const x = 0.6, y = 1.28 + i * 0.57;
      OV(s, { x, y, w: 0.44, h: 0.44, fill: { color: i === 2 ? C.navy : C.magenta } });
      T(s, String(i + 1), { x, y, w: 0.44, h: 0.44, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, it[0], { x: x + 0.62, y: y - 0.03, w: 6.8, h: 0.28, fontSize: 15, bold: true, color: C.navy });
      T(s, it[1], { x: x + 0.62, y: y + 0.25, w: 6.8, h: 0.22, fontSize: 11, color: C.text });
      if (i < 5) LN(s, x + 0.62, y + 0.51, 8.3, y + 0.51, "E4E7EB", 0.5);
    });
    s.addNotes("Abertura sugerida: 'Vou mostrar por que mudamos a lógica do TN Manager, para quem é a versão Web, o que nos diferencia e o que precisamos decidir hoje.' O bloco 3 é o coração da conversa: o sistema é o meio; suporte, dashboards e equipe técnica são o diferencial. O anexo tem o material de apoio.");
  }

  // N2 Nossa história
  {
    const s = content("Nossa história: da automação para o resultado", "A v1.0 começou pelo fim. A versão Web começa pela rotina do curral e termina no resultado do pecuarista.", "Fontes: Probeef (95% dos confinamentos analisados usam software de gestão); Embrapa (2018): 106 softwares para pecuária de corte, vistos como complexos e caros.");
    chap(s, 1);
    const chain = (y, label, sub, steps, cols, txt) => {
      T(s, label, { x: 0.5, y: y + 0.04, w: 1.75, h: 0.3, fontSize: 13, bold: true, color: cols[0] === "9AA3AC" ? C.muted : C.navy });
      T(s, sub, { x: 0.5, y: y + 0.36, w: 1.75, h: 0.36, fontSize: 10, italic: true, color: C.muted });
      const x0 = 2.3, w = 7.2 / steps.length;
      steps.forEach((st, i) => {
        s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: x0 + i * w, y, w: w + 0.12, h: 0.78, fill: { color: cols[i] }, line: { color: C.white, width: 2 } });
        T(s, st, { x: x0 + i * w + (i === 0 ? 0.1 : 0.34), y, w: w - 0.42, h: 0.78, fontSize: 10, bold: true, color: txt ? txt[i] : C.white, align: "center", valign: "middle" });
      });
    };
    chain(1.22, "TN Manager v1.0", "2023: começou pelo fim", ["Automação e hardware", "Integrações", "Grandes confinamentos", "Implantação longa e cara"], ["9AA3AC", "9AA3AC", "9AA3AC", "9AA3AC"]);
    chain(2.15, "TN Manager Web", "2026: começa pelo começo", ["Rotina do curral e inputs manuais", "Sistema objetivo e flexível", "Análise técnica TN", "Resultado do pecuarista", "Automação: PecCode hoje, TN no futuro"], [C.tealDark, C.navy, C.navy, C.magenta, "E3E7EC"], [C.white, C.white, C.white, C.white, C.muted]);
    const pr = [["Simples", "Inputs manuais com poucos campos: 10 minutos por dia", C.tealDark], ["Objetivo", "Só os indicadores que mudam a decisão", C.navy], ["Flexível", "Celular ou papel, com ou sem internet no curral", C.magenta]];
    pr.forEach((p, i) => card(s, 0.5 + i * 1.8, 3.15, 1.7, 1.55, p[2], p[0], p[1], { ts: 15, fs: 11 }));
    R(s, { x: 5.95, y: 3.15, w: 3.55, h: 1.55, fill: { color: C.navy } });
    T(s, "MENOS É MAIS", { x: 6.15, y: 3.27, w: 3.2, h: 0.24, fontSize: 10, bold: true, color: C.lime, charSpacing: 3 });
    T(s, "Sistema virou commodity. O que manda é o resultado que ajudamos o pecuarista a alcançar.", { x: 6.15, y: 3.55, w: 3.2, h: 1.1, fontSize: 15, bold: true, color: C.white });
    s.addNotes("Fala sugerida: 'A primeira versão foi desenhada do fim para o começo: automação, integração com balança e vagão, foco nos grandes. Entrei para inverter essa lógica: começar pela rotina do curral, com inputs manuais e um sistema objetivo e flexível para o confinamento pequeno e médio. Menos é mais. Sistema virou commodity: o que manda é o resultado que ajudamos o pecuarista a alcançar.' Ponte: 'E para quem é isso?'");
  }

  // N3 Bússola
  {
    const s = content("Bússola Farm: 14.038 confinamentos mapeados", "85% são pequenos, mas somam só 23% da capacidade estática: muitas fazendas, pouco volume cada.", "Fonte: Bússola Farm, dados 2024/2025. Percentuais sobre os 13.214 confinamentos com faixa informada.");
    chap(s, 2);
    [["14.038", "confinamentos mapeados", C.navy, 0.5, 1.3], ["14.985.249", "cabeças de capacidade estática", C.tealDark, 4.6, 2.05]].forEach((k) => {
      T(s, k[0], { x: k[3], y: 1.15, w: k[4], h: 0.45, fontSize: 24, bold: true, color: k[2], valign: "middle" });
      T(s, k[1], { x: k[3] + k[4] + 0.1, y: 1.15, w: 2.4, h: 0.45, fontSize: 12, bold: true, color: C.text, valign: "middle" });
    });
    const fx = [
      ["Faixa 1: pequenos", "até 1.500 cabeças", "11.194", "3,5 mi", 84.7, 23, C.magenta],
      ["Faixa 2: médios", "até 6.900 cabeças", "1.664", "5,5 mi", 12.6, 37, C.tealDark],
      ["Faixa 3: grandes", "mais de 7.000 (média de 16 mil)", "356", "6,0 mi", 2.7, 40, C.navy],
    ];
    fx.forEach((f, i) => {
      const y = 1.72 + i * 0.86;
      R(s, { x: 0.5, y, w: 9.0, h: 0.78, fill: { color: i === 0 ? C.magLight : C.bg } });
      R(s, { x: 0.5, y, w: 0.07, h: 0.78, fill: { color: f[6] } });
      T(s, f[0], { x: 0.68, y: y + 0.1, w: 2.3, h: 0.3, fontSize: 14, bold: true, color: f[6] });
      T(s, f[1], { x: 0.68, y: y + 0.42, w: 2.3, h: 0.28, fontSize: 10.5 });
      T(s, f[2], { x: 2.95, y: y + 0.08, w: 1.4, h: 0.42, fontSize: 24, bold: true, color: f[6] });
      T(s, "confinamentos", { x: 2.95, y: y + 0.5, w: 1.4, h: 0.22, fontSize: 10, color: C.muted });
      T(s, f[3], { x: 4.4, y: y + 0.08, w: 1.4, h: 0.42, fontSize: 24, bold: true, color: C.navy });
      T(s, "cabeças", { x: 4.4, y: y + 0.5, w: 1.4, h: 0.22, fontSize: 10, color: C.muted });
      [[f[4], f[6], "dos confinamentos"], [f[5], "9AA3AC", "da capacidade"]].forEach((b, k) => {
        const by = y + 0.14 + k * 0.29, w = Math.max(0.05, 2.5 * b[0] / 100), inside = w > 1.9;
        R(s, { x: 5.95, y: by, w, h: 0.23, fill: { color: b[1] } });
        T(s, Math.round(b[0]) + "% " + b[2], { x: inside ? 6.03 : 5.95 + w + 0.06, y: by - 0.01, w: 1.9, h: 0.25, fontSize: 10, bold: true, color: inside ? C.white : C.text, valign: "middle" });
      });
    });
    RR(s, { x: 0.5, y: 4.38, w: 4.6, h: 0.36, fill: { color: C.lime }, rectRadius: 0.18 });
    T(s, "Público do TN Manager Web: Faixa 1 e início da Faixa 2", { x: 0.5, y: 4.38, w: 4.6, h: 0.36, fontSize: 11, bold: true, color: C.navy, align: "center", valign: "middle" });
    s.addNotes("Bússola Farm, dados 2024/2025: 14.038 confinamentos mapeados e 14.985.249 cabeças de capacidade estática. Faixa 1 (até 1.500 cabeças, unidades até 2,4 ha): 11.194 e 3,5 milhões. Faixa 2 (até 6.900, até 9 ha): 1.664 e 5,5 milhões. Faixa 3 (mais de 7.000, média de 16.000, acima de 10 ha): 356 e 6 milhões. As três faixas somam 13.214 confinamentos; os demais não vêm com faixa. Fala: 'São mais de 11 mil confinamentos pequenos: 85% do total, com só 23% da capacidade. O caso se sustenta com muitas fazendas e custo baixo por cliente.' Ponte: 'E onde estão os nossos clientes hoje?'");
  }

  // N4 Mapa da base TN
  {
    const s = content("Nossa base: 2.796 confinadores cadastrados", "Clientes no sistema da Trouw, de janeiro a outubro de 2026. RO, SP e MT somam 65%.", "Fonte: clientes cadastrados no sistema Trouw Nutrition (jan a out/2026); cabeças por estado: censo dsm-firmenich 2025. Mapa: @svg-maps/brazil (CC BY 4.0).");
    chap(s, 2);
    const CEN = JSON.parse(require("fs").readFileSync(__dirname + "/mapa-centroides.json", "utf8"));
    const Dv = { ac: 78, al: 3, am: 34, ba: 117, df: 4, es: 13, go: 93, ma: 1, mg: 200, ms: 161, mt: 341, pa: 77, pe: 3, pi: 3, pr: 94, rj: 8, rn: 11, ro: 1022, rr: 8, rs: 8, sc: 18, se: 5, sp: 461, to: 34 };
    const RED = "D7263D", mx = 0.45, my = 1.1, k = 3.85 / 613;
    s.addImage({ path: __dirname + "/mapa-brasil.png", x: mx, y: my, w: 613 * k, h: 639 * k });
    const tagAt = (uf, cx, cy, big) => {
      const txt = uf + " " + Dv[uf].toLocaleString("pt-BR"), fs = big ? 12 : 10;
      const w = 0.14 + txt.length * (big ? 0.085 : 0.07), h = big ? 0.27 : 0.23;
      RR(s, { x: cx - w / 2, y: cy - h / 2, w, h, fill: { color: C.white }, line: { color: big ? RED : "B8C2CC", width: big ? 1.25 : 0.75 }, rectRadius: 0.05 });
      T(s, [{ text: uf.toUpperCase() + " ", options: { bold: true, color: C.navy } }, { text: Dv[uf].toLocaleString("pt-BR"), options: { bold: true, color: RED } }], { x: cx - w / 2, y: cy - h / 2, w, h, fontSize: fs, align: "center", valign: "middle" });
    };
    const NUD = { go: [-10, 10], mg: [8, -6], sp: [6, 2], pa: [0, 10], am: [0, 10], ba: [6, 0], rs: [8, 0], ma: [-10, 4], pi: [-4, 6] };
    ["ro", "sp", "mt", "mg", "ms", "ba", "pr", "go", "ac", "pa", "am", "to", "sc", "rs", "rr", "ma", "pi"].forEach((uf) => {
      const d = NUD[uf] || [0, 0];
      tagAt(uf, mx + (CEN[uf][0] + d[0]) * k, my + (CEN[uf][1] + d[1]) * k, ["ro", "sp", "mt"].includes(uf));
    });
    let last = 0;
    ["rn", "pe", "al", "se", "df", "es", "rj"].forEach((uf) => {
      const px = mx + CEN[uf][0] * k, py = my + CEN[uf][1] * k, ty = Math.max(py, last + 0.28); last = ty;
      LN(s, px, py, 4.42, ty, "8A939C", 0.75);
      OV(s, { x: px - 0.03, y: py - 0.03, w: 0.06, h: 0.06, fill: { color: RED } });
      tagAt(uf, 4.73, ty, false);
    });
    [["1B3A8C", "mais de 300"], ["4A79BF", "101 a 300"], ["9DBBDF", "21 a 100"], ["DCE7F3", "1 a 20"], ["EEF0F3", "sem dado"]].forEach((l, i) => {
      R(s, { x: 0.45, y: 4.2 + i * 0.17, w: 0.24, h: 0.13, fill: { color: l[0] } });
      T(s, l[1], { x: 0.75, y: 4.17 + i * 0.17, w: 1.0, h: 0.18, fontSize: 8.5, valign: "middle" });
    });
    R(s, { x: 5.3, y: 1.18, w: 4.2, h: 0.72, fill: { color: C.navy } });
    T(s, "2.796", { x: 5.45, y: 1.18, w: 1.6, h: 0.72, fontSize: 30, bold: true, color: C.white, valign: "middle" });
    T(s, "confinadores cadastrados no sistema da Trouw", { x: 7.05, y: 1.18, w: 2.35, h: 0.72, fontSize: 11, bold: true, color: C.white, valign: "middle" });
    T(s, "Maiores estados", { x: 5.3, y: 2.0, w: 2.5, h: 0.24, fontSize: 11, bold: true, color: C.navy });
    [["RO", 1022], ["SP", 461], ["MT", 341], ["MG", 200], ["MS", 161], ["BA", 117], ["Outros 18", 495]].forEach((t, i) => {
      const y = 2.27 + i * 0.25, w = 2.3 * t[1] / 1022;
      T(s, t[0], { x: 5.3, y, w: 0.85, h: 0.22, fontSize: 10.5, bold: true, color: C.text, valign: "middle" });
      R(s, { x: 6.15, y: y + 0.03, w, h: 0.16, fill: { color: i === 6 ? C.teal : i < 3 ? C.navy : "4A79BF" } });
      T(s, t[1].toLocaleString("pt-BR"), { x: 6.2 + w, y, w: 0.6, h: 0.22, fontSize: 10.5, bold: true, color: RED, valign: "middle" });
    });
    R(s, { x: 5.3, y: 4.08, w: 4.2, h: 0.68, fill: { color: C.magLight } });
    R(s, { x: 5.3, y: 4.08, w: 0.06, h: 0.68, fill: { color: C.magenta } });
    T(s, [{ text: "Leitura: ", options: { bold: true, color: C.magenta } }, { text: "RO tem 37% da base, com operações menores: perfil da Web. GO está entre os 3 estados com mais cabeças e tem só 93 clientes: espaço para crescer." }], { x: 5.45, y: 4.08, w: 4.0, h: 0.68, fontSize: 10, color: C.text, valign: "middle" });
    s.addNotes("Base de clientes cadastrados no sistema da Trouw Nutrition, de janeiro a outubro de 2026: 2.796 confinadores. A soma dos estados dá 2.797: conferir. A sigla MD da planilha foi lida como MS. AP, CE e PB não têm cadastro. RO, SP e MT somam 1.824 (65%). Cruzamento com o censo dsm-firmenich 2025: MT 2,2 milhões de cabeças, SP 1,4, GO 1,4, MS 0,9, MG 0,8; RO não aparece entre os cinco, o que indica operações menores. Fala: 'Esta é a nossa base hoje. Rondônia concentra mais de um terço dos clientes, com confinamentos menores, o perfil da versão Web. E Goiás, um dos maiores estados em cabeças, ainda tem poucos clientes nossos.' Ponte: 'E quem está decidindo nessas fazendas?'");
  }

  // N5 Nova geração
  {
    const s = content("A nova geração quer dados para decidir", "Quem assume o confinamento busca praticidade, números e gestão de empresa.", "Fontes: Embrapa Pecuária Sudeste, Transformação digital da pecuária de corte intensiva no Brasil (2023), 184 confinamentos; Embrapa, Sebrae e Inpe, Agricultura digital no Brasil (2020).");
    chap(s, 2);
    const k = [
      ["Mais jovens", "com curso superior e consultoria são os que mais usam tecnologia digital no confinamento", "Embrapa, 184 confinamentos (2023)", C.navy],
      ["22%", "dos produtores usam aplicativo de gestão, embora 84% já usem alguma tecnologia digital", "Embrapa, Sebrae e Inpe (2020)", C.tealDark],
      ["61%", "apontam a conexão no campo como a principal barreira à tecnologia", "Embrapa, Sebrae e Inpe (2020)", C.magenta],
    ];
    k.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      R(s, { x, y: 1.2, w, h: 1.62, fill: { color: c[3] } });
      T(s, c[0], { x: x + 0.15, y: 1.28, w: w - 0.3, h: 0.5, fontSize: i === 0 ? 22 : 28, bold: true, color: C.white, valign: "middle" });
      T(s, c[1], { x: x + 0.15, y: 1.8, w: w - 0.3, h: 0.66, fontSize: 11, color: C.white });
      T(s, c[2], { x: x + 0.15, y: 2.5, w: w - 0.3, h: 0.24, fontSize: 9, italic: true, color: "DCE3F0" });
    });
    T(s, "O que a nova geração quer e como o TN Manager Web responde", { x: 0.5, y: 3.0, w: 9, h: 0.26, fontSize: 12, bold: true, color: C.navy });
    [["Praticidade", "10 minutos por dia, no celular ou no papel"], ["Dados", "Painel semanal e benchmark com fazendas parecidas"], ["Decisão", "Analista TN junto, toda semana"], ["Gestão de empresa", "Fechamento de safra e evolução ano a ano"]].forEach((c, i) => {
      card(s, 0.5 + i * 2.27, 3.32, 2.17, 1.38, [C.tealDark, C.navy, C.navy, C.magenta][i], c[0], c[1], { ts: 13, fs: 11 });
    });
    s.addNotes("Fontes verificadas. (1) Embrapa Pecuária Sudeste, 'Transformação digital da pecuária de corte intensiva no Brasil: tecnologias e perfil de adoção' (2023), com dados de 2021 de 184 confinamentos: a adoção mais intensa está ligada a produtores mais jovens, com ensino superior, com mais confiança na gestão e com acesso a consultoria especializada. Um segundo estudo da Embrapa com 185 fazendas confirma: escala, escolaridade, menor idade e consultoria aumentam a chance de usar tecnologia digital. (2) Embrapa, Sebrae e Inpe, pesquisa 'Agricultura digital no Brasil' (2020): 84,1% dos produtores usam pelo menos uma tecnologia digital, só 22,2% usam aplicativos de gestão da propriedade, e 61,4% apontam a conectividade como principal barreira. O número de 90% da versão anterior foi retirado: vinha de uma pesquisa com 119 jovens de vários setores, sem foco em pecuária. Fala: 'Quem assume a fazenda quer número e decisão. A Embrapa mostra que é justamente o produtor mais jovem, com formação e com consultoria, que adota tecnologia no confinamento. E só um em cada cinco produtores usa aplicativo de gestão: o espaço está aberto.' Ponte: 'Então, estamos no caminho certo?'");
  }

  // N6 Caminho certo
  {
    const s = content("Estamos no caminho certo? Sim.", "Três sinais confirmam o foco no confinamento pequeno e médio, com uma ressalva.", "Fontes: Bússola Farm 2024/2025; censo de confinamento dsm-firmenich 2025; base de clientes Trouw Nutrition (jan a out/2026).");
    chap(s, 2);
    const k = [
      ["85%", "dos confinamentos mapeados são pequenos, com até 1.500 cabeças", "Bússola Farm", C.magenta],
      ["45%", "das cabeças estão nos 100 maiores confinamentos, um mercado já disputado", "Censo dsm-firmenich 2025", C.navy],
      ["37%", "da nossa base está em Rondônia, com operações menores", "Base de clientes TN", C.tealDark],
    ];
    k.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      R(s, { x, y: 1.2, w, h: 1.55, fill: { color: C.bg } });
      R(s, { x, y: 1.2, w: 0.07, h: 1.55, fill: { color: c[3] } });
      T(s, c[0], { x: x + 0.2, y: 1.26, w: w - 0.3, h: 0.55, fontSize: 32, bold: true, color: c[3] });
      T(s, c[1], { x: x + 0.2, y: 1.82, w: w - 0.32, h: 0.62, fontSize: 11.5, bold: true, color: C.text });
      T(s, c[2], { x: x + 0.2, y: 2.47, w: w - 0.3, h: 0.22, fontSize: 9, italic: true, color: C.muted });
    });
    R(s, { x: 0.5, y: 2.9, w: 9.0, h: 0.6, fill: { color: C.magLight } });
    T(s, [{ text: "Ressalva: ", options: { bold: true, color: C.magenta } }, { text: "os pequenos somam só 23% da capacidade. O caso depende de muitas fazendas, custo baixo por cliente e fidelização." }], { x: 0.68, y: 2.9, w: 8.7, h: 0.6, fontSize: 12, color: C.text, valign: "middle" });
    R(s, { x: 0.5, y: 3.62, w: 9.0, h: 1.1, fill: { color: C.navy } });
    T(s, "PESQUISA COM 30 FAZENDAS: 09 A 27/11", { x: 0.68, y: 3.7, w: 8.6, h: 0.24, fontSize: 10, bold: true, color: C.lime, charSpacing: 2 });
    T(s, "Vai confirmar com números: as dores da rotina, o sistema usado hoje, quanto custa, a satisfação com o suporte do fornecedor atual e a disposição de usar e recomendar o TN Manager Web.", { x: 0.68, y: 3.96, w: 8.6, h: 0.72, fontSize: 12, color: C.white });
    s.addNotes("Fala: 'Sim, estamos no caminho certo. Os pequenos são a grande maioria dos confinamentos. Os grandes já estão concentrados e disputados. E a nossa própria base mostra o mesmo perfil. A ressalva: em volume, o pequeno pesa menos, então o caso depende de escala de fazendas e fidelização. A pesquisa com 30 fazendas, de 9 a 27 de novembro, vai confirmar isso com números, inclusive a satisfação com o suporte dos fornecedores atuais.' Dado dos 100 maiores: censo dsm-firmenich 2025 (9,25 milhões de cabeças, 2.445 propriedades monitoradas). Ponte: 'E o que entregamos para esse cliente?'");
  }

  // N7 Mapa de valor
  {
    const s = content("Mapa de valor: quem ganha o quê com o TN Manager Web", "Quatro públicos, quatro ganhos. Hipóteses que a pesquisa com 30 fazendas vai validar.");
    chap(s, 3);
    const q = [
      ["Gestor do confinamento", ["Custo e desempenho por lote, sem montar planilha", "Decidir com o dado da semana, não do mês"], C.navy, 0.5, 1.22, I.user],
      ["Técnico TN", ["Acompanhar várias fazendas à distância", "Recomendar ajuste de dieta com dado real"], C.tealDark, 5.15, 1.22, I.tech],
      ["Equipe do curral", ["Lançar trato e leitura de cocho em poucos minutos", "Tela simples, no celular ou no papel"], C.teal, 0.5, 3.0, I.hat],
      ["Trouw Nutrition", ["Fidelizar clientes em nutrição", "Porta de entrada para novos clientes"], C.magenta, 5.15, 3.0, I.bldg],
    ];
    q.forEach((c) => {
      const [t, items, col, x, y, ic] = c, w = 4.35, h = 1.68;
      R(s, { x, y, w, h, fill: { color: C.bg } });
      R(s, { x, y, w: 0.07, h, fill: { color: col } });
      OV(s, { x: x + 0.2, y: y + 0.15, w: 0.48, h: 0.48, fill: { color: col } });
      s.addImage({ data: ic, x: x + 0.31, y: y + 0.26, w: 0.26, h: 0.26 });
      T(s, t, { x: x + 0.8, y: y + 0.15, w: 3.4, h: 0.48, fontSize: 14, bold: true, color: col, valign: "middle" });
      T(s, bullets(items, { paraSpaceAfter: 5 }), { x: x + 0.22, y: y + 0.75, w: w - 0.4, h: 0.88, fontSize: 12 });
    });
    OV(s, { x: 4.12, y: 2.4, w: 1.76, h: 0.82, fill: { color: C.navy }, line: { color: C.white, width: 3 } });
    T(s, "TN Manager Web", { x: 4.12, y: 2.4, w: 1.76, h: 0.82, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addNotes("Fala: 'Todos ganham. O gestor decide com o dado da semana, a equipe do curral lança em minutos, o técnico acompanha várias fazendas à distância e a Trouw fideliza e abre novos clientes. Mas repare: o ganho não está na tela, está no que fazemos com o dado.' Ponte: 'É isso que nos diferencia.'");
  }

  // N8 Diferencial
  {
    const s = content("Nosso diferencial: o que fazemos com os dados", "O sistema é o meio. O resultado vem do suporte, dos indicadores e da equipe técnica.");
    chap(s, 3);
    R(s, { x: 0.5, y: 1.22, w: 1.9, h: 2.3, fill: { color: "E3E7EC" } });
    T(s, "Sistema", { x: 0.62, y: 1.35, w: 1.7, h: 0.32, fontSize: 15, bold: true, color: C.muted });
    T(s, "Item de prateleira: todo concorrente tem um. É o meio, não o fim.", { x: 0.62, y: 1.72, w: 1.7, h: 1.6, fontSize: 11.5, color: C.muted });
    s.addImage({ data: J.arrow, x: 2.47, y: 2.27, w: 0.22, h: 0.22 });
    const p = [
      ["Suporte excelente", "Analista online que responde rápido e técnico TN no campo", C.tealDark],
      ["Indicadores e dashboards", "Consumo, ganho de peso, custo da arroba e dias de cocho, toda semana", C.navy],
      ["Equipe técnica", "Ajustes de dieta e manejo baseados nos dados da fazenda", C.navy],
    ];
    p.forEach((c, i) => card(s, 2.78 + i * 2.27, 1.22, 2.17, 2.3, c[2], c[0], c[1], { ts: 14, th: 0.5, fs: 12 }));
    R(s, { x: 0.5, y: 3.9, w: 9.0, h: 0.85, fill: { color: C.magenta } });
    T(s, "RESULTADO", { x: 0.7, y: 3.98, w: 2.2, h: 0.24, fontSize: 10, bold: true, color: C.white, charSpacing: 3 });
    T(s, "Rentabilidade da fazenda: mais ganho, menor custo por arroba e decisão no tempo certo", { x: 0.7, y: 4.2, w: 8.6, h: 0.5, fontSize: 15, bold: true, color: C.white, valign: "middle" });
    s.addNotes("Fala: 'O sistema é o meio. Qualquer concorrente entrega um sistema. O nosso diferencial são três entregas em cima dele: suporte excelente, com analista online e técnico no campo; indicadores e dashboards que mostram toda semana o que importa; e uma equipe técnica que ajusta dieta e manejo com base nos dados. O resultado é o que o pecuarista quer: rentabilidade.' Ponte: 'Como isso funciona na rotina da fazenda?'");
  }

  // N9 Como funciona
  {
    const s = content("Como funciona na rotina da fazenda", "A fazenda registra (passos 1 a 3). A TN transforma em decisão (passos 4 e 5), toda semana.");
    chap(s, 3);
    R(s, { x: 5.48, y: 1.18, w: 4.17, h: 2.5, fill: { color: C.magLight } });
    T(s, "ENTREGA TÉCNICA TN", { x: 5.6, y: 1.24, w: 3.8, h: 0.22, fontSize: 10, bold: true, color: C.magenta, charSpacing: 2 });
    T(s, "A FAZENDA REGISTRA", { x: 0.5, y: 1.24, w: 3.8, h: 0.22, fontSize: 10, bold: true, color: C.muted, charSpacing: 2 });
    const st = [
      ["1", "Cadastrar", "uma vez", "Currais, lotes e dieta", C.muted],
      ["2", "Registrar", "todo dia", "Trato e leitura de cocho", C.muted],
      ["3", "Movimentar", "quando ocorre", "Entradas, pesagens, saídas", C.muted],
      ["4", "Analisar", "toda semana", "Dashboards e benchmark com o analista", C.navy],
      ["5", "Ajustar", "com o técnico TN", "Dieta, trato e manejo", C.magenta],
    ];
    st.forEach((p, i) => {
      const big = i >= 3, cx = 0.95 + i * 1.88 + (big ? 0.12 : 0), d = big ? 0.8 : 0.62, y = big ? 1.52 : 1.62;
      OV(s, { x: cx - d / 2 + 0.3, y, w: d, h: d, fill: { color: p[4] } });
      T(s, p[0], { x: cx - d / 2 + 0.3, y, w: d, h: d, fontSize: big ? 24 : 18, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[1], { x: cx - 0.5, y: 2.4, w: 1.6, h: 0.3, fontSize: big ? 15 : 13, bold: true, color: p[4], align: "center" });
      T(s, p[2], { x: cx - 0.5, y: 2.7, w: 1.6, h: 0.22, fontSize: 10, italic: true, color: C.muted, align: "center" });
      T(s, p[3], { x: cx - 0.5, y: 2.94, w: 1.6, h: 0.6, fontSize: 11, bold: big, color: C.text, align: "center" });
      if (i < 4) s.addImage({ data: J.arrow, x: cx + 0.98 + (i === 2 ? 0.12 : 0), y: 1.84, w: 0.2, h: 0.2 });
    });
    LN(s, 2.9, 3.78, 9.1, 3.78, C.magenta, 1.25, "dash");
    LN(s, 2.9, 3.58, 2.9, 3.78, C.magenta, 1.25, "dash");
    LN(s, 9.1, 3.68, 9.1, 3.78, C.magenta, 1.25, "dash");
    T(s, "Ciclo semanal: o ajuste volta para o trato", { x: 4.2, y: 3.67, w: 3.6, h: 0.22, fontSize: 10, bold: true, italic: true, color: C.magenta, align: "center", fill: { color: C.white } });
    const ph = [["Piloto", "Rodrigo e técnico de campo, até 5 fazendas", C.tealDark], ["Meses 3 a 6", "Rodrigo analisa até 20 fazendas", C.navy], ["A partir do mês 7", "Analistas online: 1 para 20 fazendas", C.magenta]];
    ph.forEach((p, i) => card(s, 0.5 + i * 3.05, 4.0, 2.9, 0.78, p[2], p[0], p[1], { ts: 11.5, th: 0.24, fs: 10.5 }));
    s.addNotes("Fala: 'Os passos 1 a 3 são registro, e qualquer sistema faz. Os passos 4 e 5 são o serviço da TN: o analista olha os dashboards, compara com o benchmark e o técnico ajusta dieta, trato e manejo. No piloto, eu e o técnico de campo acompanhamos de perto; depois eu atendo até 20 fazendas; a partir do mês 7 entram os analistas online.' Ponte: 'E no dia a dia: papel ou celular? Precisa de internet?'");
  }

  // N10 Na prática
  {
    const s = content("Na prática: o dia a dia e a arquitetura", "Do curral à recomendação, com ou sem internet no curral.");
    chap(s, 3);
    T(s, "Um dia na fazenda", { x: 0.5, y: 1.18, w: 3.6, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    [["Manhã", "Leitura de cocho antes do 1º trato"], ["Durante o dia", "Trato fornecido por baia"], ["Quando ocorre", "Entradas, pesagens e saídas"], ["Fim do dia", "Escritório confere e envia: 10 minutos"], ["Toda semana", "Relatório e análise do analista TN"]].forEach((d, i) => {
      const y = 1.55 + i * 0.6;
      OV(s, { x: 0.52, y: y + 0.06, w: 0.18, h: 0.18, fill: { color: i === 4 ? C.magenta : C.navy } });
      if (i < 4) LN(s, 0.61, y + 0.24, 0.61, y + 0.66, C.line, 1);
      T(s, d[0], { x: 0.82, y, w: 3.2, h: 0.24, fontSize: 12, bold: true, color: i === 4 ? C.magenta : C.navy });
      T(s, d[1], { x: 0.82, y: y + 0.25, w: 3.25, h: 0.3, fontSize: 11 });
    });
    T(s, "Arquitetura proposta", { x: 4.4, y: 1.18, w: 5, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    const box = (x, y, w, h, ic, t1, t2, col) => {
      R(s, { x, y, w, h, fill: { color: col } });
      s.addImage({ data: ic, x: x + 0.1, y: y + 0.12, w: 0.3, h: 0.3 });
      T(s, t1, { x: x + 0.48, y: y + 0.07, w: w - 0.55, h: 0.24, fontSize: 11, bold: true, color: C.white });
      T(s, t2, { x: x + 0.48, y: y + 0.32, w: w - 0.55, h: h - 0.36, fontSize: 9.5, color: C.white });
    };
    box(4.4, 1.55, 1.85, 0.72, J.mobile, "Celular", "No curral; sem sinal, envia depois", C.tealDark);
    box(4.4, 2.37, 1.85, 0.72, J.paper, "Ficha de papel", "Digitada na sede", C.tealDark);
    box(6.55, 1.55, 1.3, 1.54, J.cloud, "Nuvem", "Dados e cálculos: consumo, ganho, custo", C.navy);
    box(8.15, 1.55, 1.4, 0.72, J.chart, "Dashboards", "Relatório semanal", C.navy);
    box(8.15, 2.37, 1.4, 0.72, J.user, "Analista TN", "Interpreta e recomenda", C.magenta);
    s.addImage({ data: J.arrow, x: 6.3, y: 1.81, w: 0.2, h: 0.2 });
    s.addImage({ data: J.arrow, x: 6.3, y: 2.63, w: 0.2, h: 0.2 });
    s.addImage({ data: J.arrow, x: 7.9, y: 1.81, w: 0.2, h: 0.2 });
    LN(s, 8.85, 2.27, 8.85, 2.37, C.muted, 1.25);
    LN(s, 8.85, 3.09, 8.85, 3.24, C.magenta, 1.25);
    LN(s, 5.3, 3.24, 8.85, 3.24, C.magenta, 1.25);
    LN(s, 5.3, 3.09, 5.3, 3.24, C.magenta, 1.25);
    T(s, "Recomendação volta ao gestor", { x: 5.6, y: 3.27, w: 3.0, h: 0.22, fontSize: 10, bold: true, italic: true, color: C.magenta, align: "center" });
    R(s, { x: 4.4, y: 3.6, w: 5.15, h: 1.15, fill: { color: C.bg } });
    s.addImage({ data: J.wifi, x: 4.52, y: 3.7, w: 0.3, h: 0.3 });
    T(s, "Precisa de internet?", { x: 4.92, y: 3.68, w: 4.5, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    T(s, bullets(["No curral, não: papel ou celular que envia depois", "Na sede, sim: uma conexão por dia basta", "Uso sem sinal no celular: validar com a Dreamer (16 a 27/11)"]), { x: 4.55, y: 4.0, w: 4.95, h: 0.75, fontSize: 10.5 });
    s.addNotes("Fala: 'No dia a dia: leitura de cocho de manhã, trato registrado por baia, eventos quando acontecem e conferência no fim do dia, em 10 minutos. Lança no celular ou na ficha de papel. No curral não precisa de internet; na sede, uma conexão por dia basta.' A Embrapa, o Sebrae e o Inpe (2020) mostram que 61% dos produtores apontam a conexão como principal barreira: por isso o produto funciona sem sinal no curral. O uso sem sinal no celular será validado com a Dreamer na imersão técnica. Ponte: 'E a concorrência?'");
  }

  // N11 Concorrência
  {
    const s = content("Concorrência: falham em custo e suporte", "Todos vendem sistema. A pesquisa com 30 fazendas vai medir custo e satisfação com o suporte de cada um.", "Fontes: Cargill e Integra Software (Feed Manager); AgFeed e dsm-firmenich (FarmTell); Ponta Agro (TGC); AgFeed (Tecnobeef); Baguete (JetBov). Valor Cargill: informação de mercado levantada pela TN. Licença TGC: valores pagos pela TN. Suporte: avaliação interna e relato de clientes.");
    chap(s, 4);
    const cols = [["", 1.3, C.navy], ["Ponto forte", 2.05, C.navy], ["Ponto fraco", 2.05, C.navy], ["Implantação", 1.65, C.magenta], ["Suporte", 1.95, C.magenta]];
    let x = 0.5; const xs = [];
    cols.forEach((c) => { xs.push(x); if (c[0]) head(s, c[0], x, 1.18, c[1] - 0.04, c[2], 0.34, 11.5); x += c[1]; });
    const rows = [
      ["Cargill", "Solução completa e automatizada (Feed Manager, CattleView)", "Feita para quem já tem automação e estrutura", "Até R$ 240 mil", "Falho; sistema mantido por terceira (Integra)"],
      ["DSM (FarmTell)", "Tecnologia com consultoria; versão simples para pequenos", "Marca em transição; muitos produtos separados", "Não divulgado", "Poucos consultores de alto nível"],
      ["Ponta (TGC)", "Referência nos grandes confinamentos", "Complexo e caro para o pequeno", "R$ 10 mil a 40 mil por ano", "Pós-venda decepcionou clientes TN"],
      ["Tecnobeef", "Nasceu de confinador (grupo Campanelli)", "Sem sistema para o cliente", "Não se aplica", "Técnico de nutrição, sem dados"],
      ["JetBov e outros", "Barato e simples, para pequenos", "Só sistema: sem nutrição e sem técnico", "Mensalidade baixa", "Remoto, sem técnico no campo"],
    ];
    rows.forEach((r, i) => {
      const y = 1.56 + i * 0.62;
      R(s, { x: 0.5, y, w: 9.0, h: 0.58, fill: { color: i % 2 ? C.white : C.bg } });
      T(s, r[0], { x: 0.58, y, w: 1.2, h: 0.58, fontSize: 12, bold: true, color: C.navy, valign: "middle" });
      for (let k = 1; k < 5; k++) T(s, r[k], { x: xs[k] + 0.06, y, w: cols[k][1] - 0.14, h: 0.58, fontSize: k >= 3 ? 10.5 : 10, bold: k >= 3, color: k >= 3 ? C.magenta : C.text, valign: "middle" });
    });
    s.addNotes("Fala: 'Todos vendem sistema, e alguns são muito bons, como o Feed Manager da Cargill e o FarmTell da DSM. Mas falham em dois pontos que o confinador sente no bolso e no dia a dia: o custo da implantação, que na Cargill chega a R$ 240 mil, e o suporte depois da venda. Atenção ao FarmTell Beef Smart, a versão simples da DSM para fazendas menores: é o concorrente mais parecido com a nossa versão Web.' Fontes e cuidados: o Feed Manager cobre manejo, nutrição, trato, fábrica de ração e SISBOV; o histórico de versões é publicado pela Integra Software. A DSM reuniu as soluções da Prodap na marca FarmTell (Beef, Views e Beef Smart). A Ponta afirma atender 12 dos 20 maiores confinamentos. A Tecnobeef atende cerca de 170 propriedades (AgFeed). O valor de até R$ 240 mil da Cargill é informação de mercado levantada pela TN; a licença TGC de R$ 10 mil a 40 mil por ano é o que a TN paga; as avaliações de suporte vêm da avaliação interna e de relatos de clientes. A pesquisa com 30 fazendas vai perguntar custo e satisfação com o suporte. A PecCode fica fora porque é parceira. Ponte: 'É aí que vamos vencer.'");
  }

  // N12 Onde vamos vencer
  {
    const s = content("Onde vamos vencer: suporte e inteligência de dados", "O sistema é o meio. O que fazemos com os dados da fazenda é o diferencial.", "Fontes: avaliação interna TN e relatos de clientes sobre o suporte dos fornecedores atuais, a confirmar na pesquisa com 30 fazendas; licenças pagas pela TN.");
    chap(s, 4);
    head(s, "Onde os concorrentes falham", 0.5, 1.18, 3.3, "7A858F", 0.34, 12);
    head(s, "Onde a TN vai vencer", 4.2, 1.18, 5.3, C.magenta, 0.34, 12);
    [["Suporte fraco", "Depois da venda, o cliente fica sozinho", "Suporte excelente", "Analista online que responde rápido e técnico TN no campo"],
     ["Implantação cara e longa", "Até R$ 240 mil e meses para começar", "Implantação simples", "Começa em dias, sem licença para o cliente"],
     ["Dado parado no sistema", "Ninguém transforma o número em decisão", "Dashboards e equipe técnica", "Painéis criativos e ajustes baseados nos dados"]].forEach((r, i) => {
      const y = 1.6 + i * 0.76;
      R(s, { x: 0.5, y, w: 3.3, h: 0.68, fill: { color: "E3E7EC" } });
      T(s, r[0], { x: 0.62, y: y + 0.07, w: 3.1, h: 0.26, fontSize: 12.5, bold: true, color: C.muted });
      T(s, r[1], { x: 0.62, y: y + 0.35, w: 3.1, h: 0.3, fontSize: 10.5, color: C.muted });
      s.addImage({ data: J.arrow, x: 3.9, y: y + 0.24, w: 0.2, h: 0.2 });
      R(s, { x: 4.2, y, w: 5.3, h: 0.68, fill: { color: C.magLight } });
      R(s, { x: 4.2, y, w: 0.06, h: 0.68, fill: { color: C.magenta } });
      T(s, r[2], { x: 4.38, y: y + 0.07, w: 5.0, h: 0.26, fontSize: 13, bold: true, color: C.magenta });
      T(s, r[3], { x: 4.38, y: y + 0.35, w: 5.0, h: 0.3, fontSize: 11, bold: true, color: C.text });
    });
    T(s, "GRUPO DE DESENVOLVIMENTO: DO SISTEMA AO RESULTADO", { x: 0.5, y: 3.94, w: 9, h: 0.2, fontSize: 9.5, bold: true, color: C.navy, charSpacing: 2 });
    const st = [["Sistema", "o meio", "9AA3AC"], ["Dados da fazenda", "lançados no dia a dia", C.tealDark], ["Analistas", "auditam, validam e interpretam", C.navy], ["Relatório de safra", "e benchmark: o resultado", C.magenta]];
    const w = 9.0 / 4;
    st.forEach((p, i) => {
      s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: 0.5 + i * w, y: 4.18, w: w + 0.12, h: 0.62, fill: { color: p[2] }, line: { color: C.white, width: 2 } });
      T(s, [{ text: p[0], options: { bold: true, breakLine: true } }, { text: p[1], options: { fontSize: 9.5 } }], { x: 0.5 + i * w + (i === 0 ? 0.12 : 0.36), y: 4.18, w: w - 0.5, h: 0.62, fontSize: 11.5, color: C.white, align: "center", valign: "middle" });
    });
    s.addNotes("Fala: 'Onde os concorrentes falham, nós vamos vencer: suporte excelente, implantação simples e dashboards com uma equipe técnica que ajusta a fazenda com base nos dados. No longo prazo, o Grupo de Desenvolvimento fecha o ciclo: analistas auditam, validam e interpretam os dados e entregam o relatório de safra e o benchmark. O sistema é o meio; o que fazemos com os dados é o diferencial.' Ponte: 'Com quem começamos?'");
  }

  // N13 Perfil piloto
  {
    const s = content("Perfil do cliente-alvo para o piloto", "Critérios para escolher as primeiras fazendas: novos clientes pequenos e médios, sem automação.");
    chap(s, 5);
    const cols = [
      { t: "Obrigatórios", col: C.navy, items: ["Confinamento ativo, com ciclo regular", "Registro hoje em caderno, planilha ou sistema simples", "Uma pessoa para lançar os dados todo dia", "Técnico com visita ou contato frequente"] },
      { t: "Desejáveis", col: C.teal, items: ["Pesagem na entrada e na saída", "Leitura de cocho diária", "Internet na sede da fazenda", "Gestor que dá retorno a cada 15 dias"] },
      { t: "Fora do perfil inicial", col: C.magenta, items: ["Operações com automação (seguem na PecCode)", "Sem rotina mínima de registro", "Recria a pasto, leite e outras espécies"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      R(s, { x, y: 1.22, w, h: 0.48, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.15, y: 1.22, w: w - 0.3, h: 0.48, fontSize: 14, bold: true, color: C.white, valign: "middle" });
      R(s, { x, y: 1.7, w, h: 2.6, fill: { color: C.bg } });
      T(s, bullets(c.items, { paraSpaceAfter: 8 }), { x: x + 0.1, y: 1.84, w: w - 0.22, h: 2.4, fontSize: 12 });
    });
    R(s, { x: 0.5, y: 4.38, w: 9.0, h: 0.42, fill: { color: C.magLight } });
    T(s, [{ text: "Porte: ", options: { bold: true, color: C.magenta } }, { text: "a v1.0 mirava 2.000 cabeças ou mais. A versão Web mira de 200 a 2.000 cabeças." }], { x: 0.65, y: 4.38, w: 8.7, h: 0.42, fontSize: 12, valign: "middle" });
    s.addNotes("Fala: 'Para o piloto, de até 5 fazendas, escolhemos novos clientes pequenos e médios, com rotina mínima de registro e alguém para lançar todo dia. Quem tem automação continua na PecCode.' Ponte: 'Quanto custa e como cobramos?'");
  }

  // N14 Custo-benefício
  {
    const s = content("Custo-benefício e cobrança", "Pagamos 20 licenças hoje. O Web atende novos clientes sem automação; a PecCode segue com os maiores.");
    chap(s, 5);
    T(s, "O que pagamos hoje", { x: 0.5, y: 1.18, w: 4, h: 0.28, fontSize: 13, bold: true, color: C.navy });
    const cw = [1.3, 0.8, 1.3, 1.2]; const hx = [0.5]; cw.forEach((w, i) => hx.push(hx[i] + w));
    ["Fornecedor", "Licenças", "Por ano", "Total por ano"].forEach((h, i) => head(s, h, hx[i], 1.5, cw[i] - 0.03, C.navy, 0.32, 10.5));
    [["TGC (Ponta)", "6", "R$ 10 a 40 mil", "R$ 60 a 240 mil"], ["PecCode", "14", "R$ 8 a 12,8 mil", "R$ 112 a 179 mil"], ["Total", "20", "", "R$ 172 a 419 mil"]].forEach((r, i) => {
      const y = 1.85 + i * 0.36;
      R(s, { x: 0.5, y, w: 4.57, h: 0.32, fill: { color: i === 2 ? C.magLight : C.bg } });
      r.forEach((c, k) => T(s, c, { x: hx[k] + 0.08, y, w: cw[k] - 0.1, h: 0.32, fontSize: 11, bold: i === 2 || k === 0, color: i === 2 ? C.magenta : C.text, valign: "middle" }));
    });
    [["R$ 364 mil", "orçados", C.navy], ["R$ 64 mil", "pagos", C.tealDark], ["R$ 300 mil", "a pagar", C.magenta], ["R$ 18,2 mil", "por fazenda/ano", C.navy]].forEach((k, i) => {
      const x = 0.5 + i * 1.155;
      R(s, { x, y: 3.02, w: 1.08, h: 0.8, fill: { color: C.bg } });
      R(s, { x, y: 3.02, w: 1.08, h: 0.05, fill: { color: k[2] } });
      T(s, k[0], { x: x + 0.06, y: 3.12, w: 1.0, h: 0.34, fontSize: 14, bold: true, color: k[2] });
      T(s, k[1], { x: x + 0.06, y: 3.46, w: 1.0, h: 0.32, fontSize: 10, bold: true });
    });
    T(s, "A licença média supera o resultado que muitos confinamentos pequenos geram para a TN.", { x: 0.5, y: 3.94, w: 4.57, h: 0.5, fontSize: 11, bold: true, italic: true, color: C.navy });
    head(s, "O que muda com o TN Manager Web", 5.4, 1.18, 4.1, C.navy, 0.34, 11.5);
    R(s, { x: 5.4, y: 1.52, w: 4.1, h: 1.12, fill: { color: C.bg } });
    T(s, bullets(["Novos clientes sem automação: TN Manager Web", "Fazendas atuais sem automação podem migrar e deixar de pagar de R$ 8 mil a 40 mil por ano", "Se paga quando o custo anual do Web ficar abaixo das licenças"]), { x: 5.5, y: 1.58, w: 3.92, h: 1.04, fontSize: 10.5 });
    R(s, { x: 5.4, y: 2.72, w: 4.1, h: 0.6, fill: { color: "E3F3F7" } });
    R(s, { x: 5.4, y: 2.72, w: 0.06, h: 0.6, fill: { color: C.teal } });
    T(s, [{ text: "PecCode, parceria consolidada: ", options: { bold: true, color: C.tealDark } }, { text: "maiores com automação seguem com ela. TN Manager Automação volta no futuro." }], { x: 5.55, y: 2.72, w: 3.9, h: 0.6, fontSize: 10.5, color: C.text, valign: "middle" });
    head(s, "Cobrança do cliente (proposta)", 5.4, 3.4, 4.1, C.magenta, 0.34, 11.5);
    R(s, { x: 5.4, y: 3.74, w: 4.1, h: 1.0, fill: { color: C.magLight } });
    T(s, bullets(["Sem licença para o cliente", "Custo no pacote de nutrição e serviço técnico", "Sem nutrição TN: mensalidade de serviço", "Modelo final: 18 a 29/01/2027 (item 6.3)"]), { x: 5.5, y: 3.8, w: 3.92, h: 0.92, fontSize: 10.5 });
    s.addNotes("Dados: 20 licenças pagas a parceiros. TGC (Ponta): 6 licenças de R$ 10 mil a 40 mil. PecCode: 14 licenças de R$ 8 mil a 12,8 mil. Orçado R$ 364 mil, pago R$ 64 mil. Média: R$ 18,2 mil por fazenda. Premissa: valores anuais. A economia depende de quantas das 20 fazendas não usam automação: levantar a lista. O modelo de monetização é o item 6.3 do cronograma (18 a 29/01/2027). Fala: 'Hoje pagamos caro por licenças que não escalam para o pequeno. O Web atende os novos clientes sem automação e sem licença para o cliente; o custo vai no pacote de nutrição e serviço. Os maiores com automação seguem com a PecCode.' Ponte: 'Pondo tudo junto, onde somos fortes e onde estamos expostos?'");
  }

  // N15 SWOT
  {
    const s = content("Análise SWOT: TN Manager Web", "Itens em magenta são novos em relação à v1.0.");
    chap(s, 5);
    const q = [
      ["F", "Forças", C.navy, 0.5, 1.15, [["Suporte excelente, com analista online", 1], ["Dashboards criativos para decidir", 1], ["Ganho de tempo: relatório pronto toda semana", 1], ["Benchmarking e comunidade (Grupo)", 1], ["Dados próprios, sem custo de licença", 1], ["Marca forte e equipe técnica de campo", 0]]],
      ["F", "Fraquezas", C.muted, 5.05, 1.15, [["Depende da disciplina de lançamento manual", 1], ["Time de análise pequeno no início", 1], ["Ainda sem casos de sucesso próprios", 0], ["Modelo de cobrança a definir", 0]]],
      ["O", "Oportunidades", C.tealDark, 0.5, 3.0, [["11.194 confinamentos pequenos (Bússola Farm)", 1], ["Concorrentes com suporte fraco e implantação cara", 1], ["Nova geração quer dados e gestão de empresa", 1], ["Só 22% dos produtores usam app de gestão", 1], ["Recuperar ex-clientes e abrir não clientes", 0]]],
      ["A", "Ameaças", C.magenta, 5.05, 3.0, [["Cargill e DSM vendem sistema com a nutrição", 1], ["DSM já tem versão simples (Beef Smart)", 1], ["Baixa adesão ao lançamento manual", 1], ["Dependência da software house", 0], ["Não embutir o custo no preço do produto", 0]]],
    ];
    q.forEach((c) => {
      const [L, t, col, x, y, items] = c, w = 4.45, h = 1.78;
      R(s, { x, y, w, h, fill: { color: C.bg } });
      R(s, { x, y, w: 0.07, h, fill: { color: col } });
      OV(s, { x: x + 0.15, y: y + 0.1, w: 0.38, h: 0.38, fill: { color: col } });
      T(s, L, { x: x + 0.15, y: y + 0.1, w: 0.38, h: 0.38, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, t, { x: x + 0.62, y: y + 0.1, w: 3, h: 0.38, fontSize: 14, bold: true, color: col, valign: "middle" });
      T(s, items.map((it, j) => ({ text: it[0], options: { bullet: { indent: 10 }, breakLine: j < items.length - 1, paraSpaceAfter: 2, color: it[1] ? C.magenta : C.text, bold: true } })), { x: x + 0.18, y: y + 0.55, w: w - 0.3, h: h - 0.6, fontSize: 10.5 });
    });
    s.addNotes("Fala: 'As forças novas são justamente o nosso diferencial: suporte, dashboards e comunidade. A fraqueza principal é depender do lançamento manual, e por isso o produto precisa ser muito simples. A ameaça que mais preocupa é a DSM, que já tem uma versão simples para pequenos.' Fonte dos 22%: Embrapa, Sebrae e Inpe (2020). Ponte: 'Como levamos isso ao mercado?'");
  }

  // N16 GTM
  {
    const s = content("Go-to-market: piloto, 20 fazendas e analistas online", "O Rodrigo valida o modelo antes de contratar. O gatilho para crescer é resultado, não data.");
    chap(s, 5);
    [["Piloto (item 8.0)", "Até 5 fazendas", "Rodrigo e técnico de campo de perto; início definido no Roadmap (6.1)", C.tealDark],
     ["Meses 3 a 6", "Até 20 fazendas", "Rodrigo analisa toda semana, com relatório comentado e auditoria mensal", C.navy],
     ["A partir do mês 7", "1 analista para 20 fazendas", "Analistas online cobram, auditam, validam e interpretam", C.magenta]].forEach((p, i) => {
      const x = 0.5 + i * 3.08, w = 2.85;
      R(s, { x, y: 1.2, w, h: 1.62, fill: { color: C.bg } });
      R(s, { x, y: 1.2, w, h: 0.06, fill: { color: p[3] } });
      T(s, p[0].toUpperCase(), { x: x + 0.14, y: 1.32, w: w - 0.2, h: 0.22, fontSize: 9.5, bold: true, color: p[3], charSpacing: 1.5 });
      T(s, p[1], { x: x + 0.14, y: 1.55, w: w - 0.2, h: 0.32, fontSize: 15, bold: true, color: C.navy });
      T(s, p[2], { x: x + 0.14, y: 1.92, w: w - 0.26, h: 0.85, fontSize: 11 });
      if (i < 2) s.addImage({ data: J.arrow, x: x + w + 0.03, y: 1.9, w: 0.18, h: 0.18 });
    });
    [["Primeiras fazendas", "Novos clientes de 200 a 2.000 cabeças, sem automação, indicados por técnicos e RTVs"],
     ["O que recebem", "Sistema, relatório semanal comentado, visita do técnico e auditoria mensal"],
     ["Gatilho para analistas", "20 fazendas ativas e G2 aprovado: adesão, uso e recomendação"],
     ["Plano e preço", "GTM e monetização de 18 a 29/01/2027 (6.2 e 6.3); aprovação de 01 a 05/02 (7.0)"],
     ["Início do piloto", "Na abertura da safra de confinamento, alinhado à ocupação das fazendas"]].forEach((r, i) => {
      const y = 2.98 + i * 0.35;
      R(s, { x: 0.5, y, w: 9.0, h: 0.32, fill: { color: i % 2 ? C.white : C.bg } });
      T(s, r[0], { x: 0.6, y, w: 2.1, h: 0.32, fontSize: 11, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.7, y, w: 6.75, h: 0.32, fontSize: 10.5, valign: "middle" });
    });
    s.addNotes("Datas conferidas com o cronograma em Excel (Cronograma_TN_Manager_Web_v2.xlsx): estruturação do GTM (6.2) e monetização (6.3) de 18 a 29/01/2027; apresentação final de BP e GTM (7.0) de 01 a 05/02/2027; piloto (8.0) com sugestão de 5 fazendas, tamanho e data definidos no Roadmap (6.1, 05 a 15/01/2027), alinhados à ocupação dos confinamentos. Os meses do go-to-market contam a partir do início do piloto. Capacidade estimada do Rodrigo: cerca de 1 hora por fazenda por semana. Ponte: 'E o que mantém a fazenda com a gente depois do piloto?'");
  }

  // N17 Grupo
  {
    const s = content("Grupo de Desenvolvimento de Confinadores", "Comunidade, dados e reconhecimento: o modelo do Fazenda Nota 10 (2020), agora com sistema.");
    chap(s, 5);
    [["Lançar", "Rotina no TN Manager Web", C.tealDark], ["Analista online", "Cobra, audita e valida", C.tealDark], ["Indicadores", "Dashboards da semana", C.navy], ["Benchmarking", "Fazenda x grupo", C.navy], ["Safra", "Relatório e evolução", C.magenta], ["Evento anual", "Análise e premiação", C.magenta]].forEach((p, i) => {
      const x = 0.5 + i * 1.52, w = 1.42;
      R(s, { x, y: 1.2, w, h: 1.3, fill: { color: C.bg } });
      R(s, { x, y: 1.2, w, h: 0.06, fill: { color: p[2] } });
      OV(s, { x: x + 0.1, y: 1.34, w: 0.36, h: 0.36, fill: { color: p[2] } });
      T(s, String(i + 1), { x: x + 0.1, y: 1.34, w: 0.36, h: 0.36, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[0], { x: x + 0.1, y: 1.76, w: w - 0.14, h: 0.26, fontSize: 12, bold: true, color: p[2] });
      T(s, p[1], { x: x + 0.1, y: 2.04, w: w - 0.14, h: 0.42, fontSize: 10.5, bold: true });
      if (i < 5) s.addImage({ data: J.arrow, x: x + w - 0.04, y: 1.44, w: 0.14, h: 0.14 });
    });
    LN(s, 9.21, 2.5, 9.21, 2.72, C.magenta, 1.25, "dash");
    LN(s, 1.21, 2.72, 9.21, 2.72, C.magenta, 1.25, "dash");
    LN(s, 1.21, 2.5, 1.21, 2.72, C.magenta, 1.25, "dash");
    T(s, "O evento atrai novas fazendas e o ciclo recomeça", { x: 3.0, y: 2.61, w: 4.0, h: 0.24, fontSize: 10.5, bold: true, italic: true, color: C.magenta, align: "center", fill: { color: C.white } });
    [["Analistas online", C.tealDark, ["Cobram lançamentos", "Auditam e validam os dados", "Interpretam com o gestor", "Tiram dúvidas"]],
     ["A fazenda ganha", C.navy, ["Referência para decidir", "Evolução safra a safra", "Reconhecimento no evento", "Rede com outros confinadores"]],
     ["A TN ganha", C.magenta, ["Fidelização pelo grupo", "Base de dados própria", "Novas fazendas querendo entrar"]]].forEach((b, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      head(s, b[0], x, 3.0, w, b[1], 0.34, 12);
      R(s, { x, y: 3.34, w, h: 1.4, fill: { color: C.bg } });
      T(s, bullets(b[2], { paraSpaceAfter: 3 }), { x: x + 0.1, y: 3.42, w: w - 0.2, h: 1.3, fontSize: 11 });
    });
    s.addNotes("Fala: 'Em 2020 criei o Fazenda Nota 10, um grupo de fazendas que se desenvolviam juntas olhando os próprios números. A chave do longo prazo é repetir esse modelo com sistema: comunidade, indicadores, benchmarking, relatório de safra e um evento anual com premiação. Os analistas online cobram e auditam os lançamentos, validam, interpretam e tiram dúvidas. A fazenda fica pelo grupo, não pelo sistema.' Se houver números do Fazenda Nota 10, cite aqui. Ponte: 'Como isso entra no calendário?'");
  }

  // N18 Roadmap alinhado ao Excel
  {
    const s = content("Roadmap: alinhado ao cronograma do projeto", "Até fev/2027, datas do cronograma (itens 2.1 a 7.0). Depois, datas indicativas, fechadas no Roadmap (6.1).");
    chap(s, 6);
    const months = ["set", "out", "nov", "dez", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
    const x0 = 3.55, mw = 5.95 / 16, y0 = 1.15;
    T(s, "2026", { x: x0, y: y0, w: mw * 4, h: 0.2, fontSize: 9.5, bold: true, color: C.navy, align: "center" });
    T(s, "2027", { x: x0 + mw * 4, y: y0, w: mw * 12, h: 0.2, fontSize: 9.5, bold: true, color: C.navy, align: "center" });
    months.forEach((m, i) => {
      T(s, m, { x: x0 + i * mw, y: y0 + 0.2, w: mw, h: 0.18, fontSize: 8.5, color: C.muted, align: "center" });
      if (i) LN(s, x0 + i * mw, y0 + 0.4, x0 + i * mw, 4.32, "ECEFF2", 0.5);
    });
    LN(s, x0 + 4 * mw, y0, x0 + 4 * mw, 4.32, C.line, 0.75);
    LN(s, x0 + 5.17 * mw, y0 + 0.4, x0 + 5.17 * mw, 4.32, C.magenta, 0.75, "dash");
    const rows = [
      ["2.1 Business case e benchmark", "Rodrigo", 0.2, 1.74, C.tealDark, "07/09 a 23/10"],
      ["2.2 Formulário com Kátia Sato", "Rodrigo e Kátia", 1.6, 2.2, C.tealDark, "19/10 a 06/11"],
      ["2.2 Pesquisa com 30 fazendas", "Rodrigo e RTVs", 2.27, 3.13, C.magenta, "09/11 a 04/12"],
      ["3.1 Imersão técnica Dreamer", "Rodrigo e Dreamer", 2.5, 2.9, C.tealDark, "16 a 27/11"],
      ["4.1 e 5.1 Diagnóstico e PRDs", "Rodrigo e Matheus", 3.2, 3.58, C.navy, "07 a 18/12"],
      ["6.1 a 6.3 Roadmap, GTM e preço", "Rodrigo e liderança", 4.15, 4.94, C.navy, "05 a 29/01"],
      ["7.0 Apresentação final", "Liderança", 5.0, 5.17, C.navy, "01 a 05/02"],
      ["8.0 Piloto, até 5 fazendas", "Rodrigo e técnicos", 8, 10, C.magenta, "mai a jun"],
      ["Rodrigo atende até 20 fazendas", "Rodrigo", 10, 14, C.magenta, "jul a out"],
      ["Analistas online e safra", "Diretoria aprova", 14, 16, C.lime, "nov e dez"],
    ];
    rows.forEach((r, i) => {
      const y = y0 + 0.45 + i * 0.27;
      T(s, r[0], { x: 0.5, y, w: 2.2, h: 0.24, fontSize: 9.5, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.65, y, w: 0.92, h: 0.24, fontSize: 8, color: C.muted, valign: "middle" });
      const bx = x0 + r[2] * mw, bw = Math.max(0.08, (r[3] - r[2]) * mw);
      RR(s, { x: bx, y: y + 0.03, w: bw, h: 0.19, fill: { color: r[4] }, rectRadius: 0.04 });
      const inside = bw > 0.7;
      T(s, r[5], { x: inside ? bx : bx + bw + 0.04, y: y + 0.02, w: inside ? bw : 1.0, h: 0.21, fontSize: 8.5, bold: true, color: inside ? (r[4] === C.lime ? C.navy : C.white) : r[4], align: inside ? "center" : "left", valign: "middle" });
    });
    const gy = 4.35;
    T(s, "Decisões", { x: 0.5, y: gy - 0.02, w: 2, h: 0.24, fontSize: 10, bold: true, color: C.magenta, valign: "middle" });
    [["G0: escopo", 3.36, "right"], ["Aprovação", 5.17, "left"], ["G1: piloto", 10, "center"], ["G2: escala", 14, "center"]].forEach((g) => {
      const gx = x0 + g[1] * mw;
      s.addShape(pres.shapes.DIAMOND, { x: gx - 0.1, y: gy, w: 0.2, h: 0.2, fill: { color: C.lime }, line: { type: "none" } });
      const lx = g[2] === "right" ? gx - 1.1 : g[2] === "left" ? gx - 0.1 : gx - 0.6;
      T(s, g[0], { x: lx, y: gy + 0.21, w: 1.2, h: 0.2, fontSize: 9, bold: true, color: C.navy, align: g[2] });
    });
    s.addNotes("Linhas 2.1 a 7.0 seguem o cronograma em Excel (Cronograma_TN_Manager_Web_v2.xlsx, versão de 09/10/2026): 2.1 business case de 07/09 a 16/10 e 2.1.1 entrevistas e benchmark até 23/10; 2.2.1 e 2.2.2 formulário (rascunho até 23/10 e revisão com Kátia Sato de 26/10 a 06/11); 2.2.3 e 2.2.4 pesquisa com 30 fazendas de 09 a 27/11 e consolidação até 04/12; 3.1 imersão técnica com a Dreamer de 16 a 27/11; 4.1 diagnóstico de 07 a 11/12 e 5.1 PRDs de 14 a 18/12; 6.1 roadmap de 05 a 15/01, 6.2 GTM e 6.3 monetização de 18 a 29/01; 7.0 apresentação final de 01 a 05/02/2027. O piloto (8.0) tem data definida no Roadmap (6.1), com sugestão de 5 fazendas, alinhado à ocupação dos confinamentos; maio é indicativo (início da safra de confinamento). Ponto de atenção do cronograma: a entrega da Dreamer está prevista para o fim de novembro, antes do resultado da pesquisa; até lá, focar em correções. A linha tracejada magenta marca onde as datas passam a ser indicativas. Ponte: 'E a visão até 2030?'");
  }

  // N19 Cronograma 4 anos
  {
    const s = content("Cronograma de 4 anos: 2027 a 2030", "Do piloto à comunidade de confinadores TN. Metas propostas, revistas a cada gate.");
    chap(s, 6);
    const x0 = 2.05, cw = 1.86;
    [["2027", "Piloto e 20 fazendas"], ["2028", "Grupo e 1º evento"], ["2029", "Benchmark e fidelidade"], ["2030", "Programa de dados"]].forEach((y, i) => {
      const col = [C.tealDark, C.navy, C.magenta, C.lime][i];
      R(s, { x: x0 + i * cw, y: 1.15, w: cw - 0.06, h: 0.52, fill: { color: col } });
      T(s, y[0], { x: x0 + i * cw, y: 1.15, w: cw - 0.06, h: 0.28, fontSize: 14, bold: true, color: i === 3 ? C.navy : C.white, align: "center", valign: "middle" });
      T(s, y[1], { x: x0 + i * cw, y: 1.42, w: cw - 0.06, h: 0.22, fontSize: 9.5, bold: true, color: i === 3 ? C.navy : C.white, align: "center" });
    });
    const lanes = [
      ["Produto", ["Web estável e piloto", "Uso sem sinal e integração NOA", "Painel de benchmark", "Pegada de carbono e recomendação"]],
      ["Atendimento", ["Rodrigo analisa", "Analistas online (1 para 20)", "Núcleo de análise", "Análise na oferta comercial"]],
      ["Comunidade", ["Grupo piloto; 1ª safra fechada", "Grupo lançado; 1º evento", "Evento anual e benchmark", "Comunidade nacional TN"]],
      ["Fazendas", ["20", "60", "120", "200"]],
      ["Analistas", ["0 (Rodrigo)", "3", "6", "10"]],
      ["Automação", ["Maiores com a PecCode", "Preparar a retomada", "TN Manager Automação retomado", "Web e Automação lado a lado"]],
    ];
    lanes.forEach((l, i) => {
      const num = i === 3 || i === 4, h = num ? 0.4 : 0.5;
      const y = 1.75 + [0, 0.55, 1.1, 1.65, 2.1, 2.55][i];
      R(s, { x: 0.5, y, w: 1.5, h, fill: { color: C.navy } });
      T(s, l[0], { x: 0.6, y, w: 1.36, h, fontSize: 11.5, bold: true, color: C.white, valign: "middle" });
      l[1].forEach((c, k) => {
        R(s, { x: x0 + k * cw, y, w: cw - 0.06, h, fill: { color: C.bg } });
        T(s, c, { x: x0 + k * cw + 0.06, y, w: cw - 0.18, h, fontSize: num ? 15 : 10.5, bold: true, color: num ? C.navy : C.text, align: "center", valign: "middle" });
      });
    });
    s.addNotes("Metas propostas: 20 fazendas em 2027 (capacidade do Rodrigo), 60 em 2028, 120 em 2029 e 200 em 2030. Analistas online na proporção de 1 para cada 20 fazendas. Comunidade: Grupo de Desenvolvimento inspirado no Fazenda Nota 10, com fechamento de safra anual e evento com premiação a partir de 2028. Automação: os clientes maiores seguem com a PecCode; o TN Manager Automação é retomado em 2029, em paralelo com a versão Web. Ponte: 'Para começar, precisamos de algumas decisões hoje.'");
  }

  // N20 Decisões
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.lime } });
    chap(s, 6, true);
    T(s, "Decisões e próximos passos", { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true, color: C.white });
    T(s, "O que precisamos aprovar hoje para seguir com a pesquisa e a imersão técnica.", { x: 0.5, y: 0.8, w: 8.6, h: 0.3, fontSize: 13, italic: true, color: "D6DCEB" });
    T(s, "DECISÕES PARA HOJE", { x: 0.5, y: 1.3, w: 4, h: 0.22, fontSize: 10, bold: true, color: C.lime, charSpacing: 2 });
    ["Foco em novos clientes de 200 a 2.000 cabeças, sem automação; maiores seguem na PecCode", "Roadmap alinhado ao cronograma, com gates G0 a G2", "Modelo: Rodrigo atende 20 fazendas; analistas online após o G2", "Imersão técnica e orçamento com a Dreamer", "Grupo de Desenvolvimento e evento anual como pilares", "Indicar um patrocinador na diretoria"].forEach((d, i) => {
      const y = 1.6 + i * 0.52;
      OV(s, { x: 0.5, y, w: 0.36, h: 0.36, fill: { color: C.magenta } });
      s.addImage({ data: I.check, x: 0.59, y: y + 0.09, w: 0.18, h: 0.18 });
      T(s, d, { x: 1.0, y: y - 0.04, w: 4.0, h: 0.46, fontSize: 11.5, bold: true, color: C.white, valign: "middle" });
    });
    R(s, { x: 5.3, y: 1.25, w: 4.25, h: 3.55, fill: { color: C.navyDark } });
    T(s, "PRÓXIMAS SEMANAS", { x: 5.5, y: 1.35, w: 3.8, h: 0.22, fontSize: 10, bold: true, color: C.lime, charSpacing: 2 });
    [["até 23/10", "Entrevistas e benchmark"], ["até 23/10", "Rascunho do formulário"], ["26/10 a 06/11", "Revisão com Kátia Sato"], ["09 a 27/11", "Pesquisa com 30 fazendas"], ["16 a 27/11", "Imersão técnica Dreamer"], ["07 a 18/12", "Diagnóstico e PRDs (G0)"], ["01 a 05/02", "Apresentação final"]].forEach((r, i) => {
      const y = 1.65 + i * 0.44;
      T(s, r[0], { x: 5.5, y, w: 1.3, h: 0.38, fontSize: 11, bold: true, color: C.lime, valign: "middle" });
      T(s, r[1], { x: 6.8, y, w: 2.65, h: 0.38, fontSize: 11, bold: true, color: C.white, valign: "middle" });
    });
    s.addNotes("Fechamento sugerido: 'Começamos pelo começo: rotina, dado, análise e resultado. O sistema é o meio; suporte, dashboards e equipe técnica são o diferencial. Com o Grupo de Desenvolvimento e o evento anual, a fazenda fica pela comunidade, e a TN ganha dados e fidelidade.'");
  }


  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
