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

  const bullets = (items, o = {}) => items.map((it, j) => ({ text: it, options: Object.assign({ bullet: { indent: 10 }, breakLine: j < items.length - 1, paraSpaceAfter: 4 }, o) }));

  // 1. Cover
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.lime } });
    T(s, "BUSINESS CASE v2.0", { x: 0.7, y: 0.9, w: 6, h: 0.3, fontSize: 13, bold: true, color: C.lime, charSpacing: 4 });
    T(s, "TN Manager and TN Manager Web", { x: 0.7, y: 1.3, w: 8.6, h: 0.75, fontSize: 36, bold: true, color: C.white });
    T(s, "Where the product creates value and for whom: target customer, value proposition, MVP, target roadmap and go-to-market", { x: 0.7, y: 2.15, w: 7.6, h: 0.8, fontSize: 16, italic: true, color: "D6DEF0" });
    T(s, [{ text: "Trouw Nutrition Brazil", options: { bold: true, breakLine: true } }, { text: "Rodrigo Gennari  |  October 2026" }], { x: 0.7, y: 3.45, w: 7, h: 0.6, fontSize: 13, color: C.white });
    s.addImage({ data: WAVE, x: 3.5, y: 4.5, w: 6.5, h: 1.125 });
    s.addNotes("Abertura (30 s). Objetivo da reunião: revisar o Business Case do TN Manager à luz do que mudou desde 2023 e sair com 3 decisões. Fale que a apresentação dura cerca de 18 minutos e o restante é discussão.");
  }

  // 2. Executive summary
  {
    const s = content("Executive Summary", "Recommendation: take a step back, prove value with a few farms, then scale.");
    R(s, { x: 0.5, y: 1.3, w: 3.6, h: 3.35, fill: { color: C.navy } });
    T(s, "RECOMMENDATION", { x: 0.75, y: 1.5, w: 3.1, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 3 });
    T(s, "Launch TN Manager Web first: a simple, manual-entry version, piloted with 5 feedlots. Each new version moves forward only after the previous one proves use and value.", { x: 0.75, y: 1.85, w: 3.1, h: 1.8, fontSize: 15, color: C.white });
    T(s, "Before scaling: survey with at least 30 feedlots, internal interviews and competitor benchmark.", { x: 0.75, y: 3.75, w: 3.1, h: 0.7, fontSize: 11, italic: true, color: "C9D3EA" });
    const pts = [
      ["The business problem is still there", "Feedlot accounts for 19% of Beef Sales Profit, and v1.0 recorded 51,000 heads lost to competitors with stronger digital support."],
      ["The target customer must change", "v1.0 targeted large, automated feedlots, where Cargill and DSM/Prodap are strongest. Our hypothesis: the bigger gap is in feedlots that still record data on paper or spreadsheets."],
      ["Validate before investing more", "The final MVP scope comes from customers, former customers and non-customers, not from internal assumptions."],
    ];
    pts.forEach((p, i) => {
      const y = 1.3 + i * 1.15;
      OV(s, { x: 4.45, y, w: 0.5, h: 0.5, fill: { color: C.magenta } });
      T(s, String(i + 1), { x: 4.45, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[0], { x: 5.15, y: y - 0.02, w: 4.35, h: 0.3, fontSize: 14, bold: true, color: C.navy });
      T(s, p[1], { x: 5.15, y: y + 0.3, w: 4.35, h: 0.75, fontSize: 11 });
    });
    s.addNotes("Comece pela conclusão (2 min). Mensagem: o problema de 2023 continua, mas a forma de resolver muda. Em vez de um sistema completo para grandes confinamentos, começamos com uma versão simples, testada em 5 fazendas, e só crescemos quando houver prova de uso. Se o chefe só ouvir este slide, ele já sabe o que você está pedindo.");
  }

  // 3. v1 x v2
  {
    const s = content("What changes from v1.0 to v2.0", "The question moves from how to defend our base to where the product creates value and for whom.");
    const H = (t, fill) => ({ text: t, options: { bold: true, color: C.white, fill: { color: fill }, fontSize: 11 } });
    const L = (t) => ({ text: t, options: { bold: true, color: C.navy, fontSize: 11 } });
    const rows = [
      [H("", C.navy), H("v1.0 (October 2023)", C.tealDark), H("v2.0 (October 2026)", C.magenta)],
      [L("Target customer"), "Large feedlots (A+ and B+), minimum 2,000 heads, with weighing and structured control", "Small and medium feedlots with manual or spreadsheet records, customers and non-customers (hypothesis)"],
      [L("Product"), "Full cloud app, hardware integration and offline terminal", "TN Manager Web: manual entry, few indicators, released in versions"],
      [L("Decision basis"), "Internal assumptions and list of lost customers", "Internal interviews, competitor benchmark and survey with 30+ feedlots"],
      [L("Rollout"), "50 customers in year 1 and 200 in year 5", "5 pilot farms, then waves with go/no-go criteria"],
      [L("Success metric"), "Sales Profit and market share", "Recurring use and perceived value first; Sales Profit next"],
    ];
    s.addTable(rows, { x: 0.5, y: 1.3, w: 9, colW: [1.7, 3.65, 3.65], fontFace: FONT, fontSize: 11, color: C.text, border: { type: "solid", pt: 0.75, color: C.line }, rowH: [0.4, 0.68, 0.58, 0.58, 0.5, 0.5], valign: "middle", margin: [0.04, 0.08, 0.04, 0.08] });
    s.addNotes("Ponte com o Business Case de 2023 (1 min). Leia só a coluna da direita. Ponto-chave: a v1.0 decidiu com base em premissas internas; a v2.0 decide com base em pesquisa com fazendas.");
  }

  // 4. Business problem
  {
    const s = content("Feedlot is still a key revenue stream at risk", "The 2023 numbers still justify the project. v2.0 changes how we solve it.", "Source: Feedlot App Business Case v1.0 (Oct/2023) and OGP Brazil 2023. To be refreshed in the diagnosis.");
    T(s, "Beef Sales Profit 2022 (BRL million)", { x: 0.5, y: 1.25, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: C.navy });
    s.addChart(pres.charts.BAR, [{ name: "SP", labels: ["Top 100 feedlots", "Pareto 80% (265 accounts)", "Feedlot", "Total Beef"], values: [6.5, 10.5, 13, 68] }], {
      x: 0.4, y: 1.55, w: 5.4, h: 3.05, barDir: "bar", chartColors: [C.magenta], barGapWidthPct: 45,
      catAxisLabelColor: C.text, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, valAxisHidden: true,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showValue: true, dataLabelPosition: "outEnd",
      dataLabelColor: C.text, dataLabelFontSize: 10, dataLabelFormatCode: "0.0", showLegend: false, valAxisMaxVal: 80,
    });
    const stats = [["19%", "of Beef Sales Profit comes from feedlots", C.magenta], ["51k", "heads lost in 7 accounts, mostly to Cargill", C.navy], ["12.5k", "heads at risk of loss in 2023", C.tealDark], ["9.6%", "market share in premix, minerals and concentrates for feedlots", C.teal]];
    stats.forEach((st, i) => {
      const x = 6.1 + (i % 2) * 1.75, y = 1.3 + Math.floor(i / 2) * 1.62;
      R(s, { x, y, w: 1.6, h: 1.45, fill: { color: C.bg } });
      R(s, { x, y, w: 0.06, h: 1.45, fill: { color: st[2] } });
      T(s, st[0], { x: x + 0.17, y: y + 0.12, w: 1.38, h: 0.5, fontSize: 24, bold: true, color: st[2] });
      T(s, st[1], { x: x + 0.17, y: y + 0.65, w: 1.38, h: 0.75, fontSize: 9.5 });
    });
    s.addNotes("Contexto em números (1 min). São dados da própria v1.0, de 2023. Deixe claro que serão atualizados no diagnóstico de dezembro. Não se aprofunde: o chefe já conhece esses números.");
  }

  // 5. Target customer review (2x2)
  {
    const s = content("Target customer review: where TN Manager Web plays", "v1.0 aimed at the most contested quadrant. v2.0 starts where pain is high and simple tools are scarce.");
    const gx = 1.15, gy = 1.25, gw = 6.0, gh = 3.4, hw = gw / 2, hh = gh / 2;
    const q = [
      { x: gx, y: gy, t: "Large, manual records", d: "High pain and high product volume. Second wave for TN Manager Web.", fill: C.magLight, col: C.magenta, tg: "PRIORITY 2" },
      { x: gx + hw, y: gy, t: "Large, automated", d: "v1.0 focus. Contested by Cargill and DSM/Prodap. Full TN Manager territory.", fill: C.bg, col: C.tealDark, tg: "TN MANAGER" },
      { x: gx, y: gy + hh, t: "Medium, manual records", d: "Most farms and fewest simple tools available. Suggested entry point for the pilot.", fill: C.magenta, col: C.white, tg: "PILOT", strong: true },
      { x: gx + hw, y: gy + hh, t: "Medium, automated", d: "Already uses some tool. Assess switch or integration in a later version.", fill: C.bg, col: C.muted, tg: "ASSESS" },
    ];
    q.forEach((k) => {
      R(s, { x: k.x + 0.04, y: k.y + 0.04, w: hw - 0.08, h: hh - 0.08, fill: { color: k.fill } });
      tag(s, k.tg, k.x + 0.2, k.y + 0.2, 1.2, k.strong ? C.navy : k.col);
      T(s, k.t, { x: k.x + 0.2, y: k.y + 0.55, w: hw - 0.4, h: 0.3, fontSize: 13, bold: true, color: k.strong ? C.white : C.navy });
      T(s, k.d, { x: k.x + 0.2, y: k.y + 0.88, w: hw - 0.4, h: 0.75, fontSize: 10.5, color: k.strong ? C.white : C.text });
    });
    T(s, "Size", { x: 0.35, y: gy, w: 0.7, h: 0.3, fontSize: 10, bold: true, color: C.navy });
    T(s, "Large", { x: 0.35, y: gy + hh / 2, w: 0.75, h: 0.3, fontSize: 10, color: C.muted });
    T(s, "Medium", { x: 0.35, y: gy + hh * 1.5, w: 0.75, h: 0.3, fontSize: 10, color: C.muted });
    T(s, "Paper or spreadsheet", { x: gx, y: gy + gh + 0.06, w: hw, h: 0.25, fontSize: 10, color: C.muted, align: "center" });
    T(s, "Scale, feed mixer and software", { x: gx + hw, y: gy + gh + 0.06, w: hw, h: 0.25, fontSize: 10, color: C.muted, align: "center" });
    T(s, "Data management maturity", { x: gx, y: gy + gh + 0.3, w: gw, h: 0.25, fontSize: 10, bold: true, color: C.navy, align: "center" });
    R(s, { x: 7.45, y: 1.25, w: 2.1, h: 3.4, fill: { color: C.navy } });
    T(s, "Why shift focus", { x: 7.62, y: 1.4, w: 1.8, h: 0.3, fontSize: 12, bold: true, color: C.lime });
    T(s, [
      { text: "Automated feedlots compare TN Manager with mature systems and larger field teams.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "Manual feedlots compare it with a notebook. A simple product already shows a visible gain.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "The survey confirms minimum size and how big each quadrant is." },
    ], { x: 7.62, y: 1.78, w: 1.8, h: 2.8, fontSize: 10, color: C.white });
    s.addNotes("Revisão do público-alvo (2 min). Tema 1 da pauta. Explique os dois eixos: tamanho da fazenda e nível de controle de dados. A v1.0 mirou o canto superior direito, onde a Cargill é mais forte. A proposta é começar no canto inferior esquerdo: fazendas médias que anotam em papel ou planilha. Deixe claro que é hipótese e que a pesquisa de novembro confirma.");
  }

  // 6. Pilot customer profile
  {
    const s = content("Target customer profile for the pilot", "Proposed criteria to select the first farms. Minimum herd size comes from the survey.");
    const cols = [
      { t: "Must have", col: C.navy, items: ["Active feedlot with regular in and out cycles", "TN customer for nucleus, premix or minerals", "Records today on paper, notebook or spreadsheet", "One person responsible for daily data entry", "TN technician with regular visits or contact"] },
      { t: "Nice to have", col: C.teal, items: ["Weighing at entry and exit", "Daily bunk reading", "Internet access at the farm office", "Manager willing to give feedback every 2 weeks"] },
      { t: "Out of initial scope", col: C.magenta, items: ["Operations already integrated with scales, mixers and software", "Feedlots with no minimum recording routine", "Pasture backgrounding, dairy and other species (kept from v1.0)"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.85;
      R(s, { x, y: 1.3, w, h: 0.45, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.15, y: 1.3, w: w - 0.3, h: 0.45, fontSize: 13, bold: true, color: C.white, valign: "middle" });
      R(s, { x, y: 1.75, w, h: 2.55, fill: { color: C.bg } });
      T(s, bullets(c.items, { paraSpaceAfter: 5 }), { x: x + 0.1, y: 1.88, w: w - 0.22, h: 2.35, fontSize: 10.5 });
    });
    R(s, { x: 0.5, y: 4.42, w: 9.0, h: 0.4, fill: { color: C.magLight } });
    T(s, [{ text: "Minimum size: ", options: { bold: true, color: C.magenta } }, { text: "v1.0 used 2,000 heads. For the manual version, the ideal range will be set with survey data." }], { x: 0.65, y: 4.42, w: 8.7, h: 0.4, fontSize: 10.5, valign: "middle" });
    s.addNotes("Perfil da fazenda piloto (1 min). Mostre que existem critérios objetivos para escolher as 5 primeiras fazendas. Destaque: no piloto, só clientes TN, porque o técnico já está lá e o risco é menor.");
  }

  // 7. Value proposition canvas
  {
    const s = content("Value proposition: TN Manager Web canvas", "Left: what we deliver. Right: what the feedlot needs. Hypotheses to validate in the survey.");
    const box = (x, y, w, h, title, items, col) => {
      R(s, { x, y, w, h, fill: { color: C.white }, line: { color: col, width: 1 } });
      R(s, { x, y, w, h: 0.3, fill: { color: col } });
      T(s, title, { x: x + 0.1, y, w: w - 0.2, h: 0.3, fontSize: 10.5, bold: true, color: C.white, valign: "middle" });
      T(s, bullets(items, { paraSpaceAfter: 3 }), { x: x + 0.06, y: y + 0.36, w: w - 0.12, h: h - 0.4, fontSize: 9 });
    };
    T(s, "VALUE MAP", { x: 0.5, y: 1.2, w: 4.2, h: 0.25, fontSize: 10, bold: true, color: C.navy, charSpacing: 2 });
    T(s, "CUSTOMER PROFILE", { x: 5.3, y: 1.2, w: 4.2, h: 0.25, fontSize: 10, bold: true, color: C.magenta, charSpacing: 2 });
    box(0.5, 1.5, 1.95, 3.2, "Products and services", ["Web app with manual data entry", "Dashboard with a few indicators", "Automatic weekly report", "Follow-up by the TN technician"], C.navy);
    box(2.55, 1.5, 2.15, 1.55, "Pain relievers", ["Short, standard entry forms", "No license fee for the farmer", "Support from the technician who already visits"], C.tealDark);
    box(2.55, 3.15, 2.15, 1.55, "Gain creators", ["Report ready for the owner", "Diet adjustment advice", "Comparison between lots"], C.teal);
    box(5.3, 1.5, 2.15, 1.55, "Pains", ["Data scattered in notebooks and spreadsheets", "Typing errors and rework", "Software that is costly, complex or unsupported"], C.magenta);
    box(5.3, 3.15, 2.15, 1.55, "Gains", ["See lot performance every week", "Compare with similar farms", "A technician reviewing the numbers together"], "A30C5A");
    box(7.55, 1.5, 1.95, 3.2, "Customer jobs", ["Adjust daily feeding", "Close cost per lot and per arroba", "Control entries, exits and deaths", "Report to owner or partner"], C.navyDark);
    OV(s, { x: 4.73, y: 2.82, w: 0.55, h: 0.55, fill: { color: C.lime } });
    T(s, "FIT", { x: 4.73, y: 2.82, w: 0.55, h: 0.55, fontSize: 10, bold: true, color: C.navy, align: "center", valign: "middle" });
    s.addNotes("Proposta de valor (2 min). Tema 2 da pauta. Leia da direita para a esquerda: primeiro o que o pecuarista precisa fazer (tarefas), o que dói e o que ele quer ganhar. Depois mostre que cada item da esquerda responde a um item da direita. Frase de efeito: o cliente não compra software, ele compra saber o custo do lote sem montar planilha.");
  }

  // 8. Two versions
  {
    const s = content("Two versions, two roles in the strategy", "TN Manager Web opens the door and prepares the customer to move to the full version.");
    const rows = [
      ["For whom", "Large feedlots with scales, mixers and a structured team", "Feedlots with manual records (hypothesis)"],
      ["Data entry", "Hardware integration and automatic capture", "Manual entry in simple forms"],
      ["Access", "Cloud and offline terminal", "Browser, on computer or phone"],
      ["Scope", "Full management, integrations and add-ons", "Five essential modules"],
      ["Rollout and support", "Implementation project with a dedicated team", "TN technician who already serves the farm"],
      ["Role in strategy", "Defend A+ and B+ accounts", "Win, teach and prepare for migration"],
    ];
    T(s, "TN Manager (v1.0 scope)", { x: 2.35, y: 1.25, w: 3.5, h: 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", fill: { color: C.tealDark } });
    T(s, "TN Manager Web (v2.0 proposal)", { x: 6.0, y: 1.25, w: 3.5, h: 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", fill: { color: C.magenta } });
    rows.forEach((r, i) => {
      const y = 1.75 + i * 0.5;
      if (i % 2 === 0) R(s, { x: 0.5, y, w: 9.0, h: 0.46, fill: { color: C.bg } });
      T(s, r[0], { x: 0.6, y, w: 1.7, h: 0.46, fontSize: 10.5, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.45, y, w: 3.3, h: 0.46, fontSize: 10, valign: "middle" });
      T(s, r[2], { x: 6.1, y, w: 3.3, h: 0.46, fontSize: 10, valign: "middle", bold: true, color: C.navy });
    });
    s.addNotes("Posicionamento (1 min). Antecipe a objeção: a versão web não é uma versão pior. Ela é a porta de entrada. Quem amadurecer migra para o TN Manager completo.");
  }

  // 9. MVP scope
  {
    const s = content("MVP: what goes into TN Manager Web", "MoSCoW prioritization. Only the first column is required for the pilot.");
    const cols = [
      { t: "Must have", st: "version 0.1, pilot", col: C.magenta, items: ["Farm, pen and lot registration", "Daily feed delivery per pen", "Bunk reading", "Entries, weighing, exits, deaths and sick pen", "Dashboard: dry matter intake, ADG, days on feed and mortality", "Weekly PDF report"] },
      { t: "Should have", st: "version 0.2", col: C.navy, items: ["Diet registration with dry matter correction", "Cost per lot and per arroba produced", "Simple alerts for intake off target"] },
      { t: "Could have", st: "version 0.3 onward", col: C.teal, items: ["Import existing spreadsheet", "Comparison between lots", "Anonymous benchmark between farms"] },
      { t: "Won't have now", st: "kept for the full version", col: C.muted, items: ["Integration with scales, mixers and hardware", "Native offline app", "Loyalty and sustainability", "Backgrounding, dairy and other species"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 2.27, w = 2.15;
      R(s, { x, y: 1.3, w, h: 0.62, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.12, y: 1.34, w: w - 0.24, h: 0.3, fontSize: 13, bold: true, color: C.white });
      T(s, c.st, { x: x + 0.12, y: 1.62, w: w - 0.24, h: 0.25, fontSize: 9.5, italic: true, color: C.white });
      R(s, { x, y: 1.92, w, h: 2.78, fill: { color: i === 0 ? C.magLight : C.bg } });
      T(s, bullets(c.items), { x: x + 0.06, y: 2.02, w: w - 0.14, h: 2.6, fontSize: 9.5, color: i === 3 ? C.muted : C.text });
    });
    s.addNotes("Descrição do MVP (2 min). Tema 3 da pauta. MVP é o produto mínimo: o menor sistema que já resolve a dor principal. Mostre só a coluna rosa. Diga o que NÃO entra (coluna cinza): isso mostra disciplina e evita o erro da v1.0 de querer tudo de uma vez. Lembrete: a Dreamer já tem 78 de 131 tarefas prontas, com entrega prevista para o fim de novembro.");
  }

  // 10. How MVP works
  {
    const s = content("How the MVP fits the farm routine", "Five steps. If it does not fit the pen routine, it does not go into the MVP.");
    const st = [["Register", "once", "Farm, pens, lots and diet"], ["Record", "every day", "Feed delivered and bunk reading"], ["Move", "when it happens", "Entries, weighing, exits and deaths"], ["Analyze", "every week", "Dashboard and report generated automatically"], ["Adjust", "with the TN technician", "Review numbers and adjust the diet"]];
    const cols = [C.navy, C.tealDark, C.teal, C.aqua, C.magenta];
    st.forEach((p, i) => {
      const x = 0.55 + i * 1.84;
      OV(s, { x: x + 0.33, y: 1.35, w: 0.95, h: 0.95, fill: { color: cols[i] } });
      T(s, String(i + 1), { x: x + 0.33, y: 1.35, w: 0.95, h: 0.95, fontSize: 26, bold: true, color: C.white, align: "center", valign: "middle" });
      if (i < 4) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 1.38, y: 1.7, w: 0.38, h: 0.25, fill: { color: C.line }, line: { type: "none" } });
      T(s, p[0], { x, y: 2.42, w: 1.6, h: 0.3, fontSize: 13, bold: true, color: cols[i], align: "center" });
      T(s, p[1], { x, y: 2.72, w: 1.6, h: 0.25, fontSize: 10, italic: true, color: C.muted, align: "center" });
      T(s, p[2], { x: x + 0.05, y: 3.0, w: 1.5, h: 0.55, fontSize: 10, align: "center" });
    });
    const pr = [["Few fields per screen", "Target: daily entry in up to 10 minutes."], ["Feedlot language", "Terms the pen crew already uses, no system jargon."], ["Light for weak internet", "Simple screens that load on farm connections."]];
    pr.forEach((p, i) => {
      const x = 0.5 + i * 3.05;
      R(s, { x, y: 3.85, w: 2.85, h: 0.85, fill: { color: C.bg } });
      R(s, { x, y: 3.85, w: 0.06, h: 0.85, fill: { color: C.lime } });
      T(s, p[0], { x: x + 0.18, y: 3.93, w: 2.6, h: 0.28, fontSize: 11.5, bold: true, color: C.navy });
      T(s, p[1], { x: x + 0.18, y: 4.22, w: 2.6, h: 0.45, fontSize: 9.5 });
    });
    s.addNotes("Como o MVP funciona (1 min). Conte como uma história: cadastra uma vez, lança todo dia em poucos minutos, e toda semana o relatório sai sozinho para o dono e para o técnico. O técnico TN fecha o ciclo ajustando a dieta. É aqui que o produto vende nutrição.");
  }

  // 11. Target roadmap
  {
    const s = content("Target roadmap: grow by versions, with a decision at each step", "Indicative dates. The pilot is timed with feedlot occupancy in the dry season.");
    const x0 = 3.05, x1 = 9.5, months = 24, mw = (x1 - x0) / months;
    const qs = ["Q4 26", "Q1 27", "Q2 27", "Q3 27", "Q4 27", "Q1 28", "Q2 28", "Q3 28"];
    qs.forEach((q, i) => {
      const x = x0 + i * 3 * mw;
      if (i > 0) LN(s, x, 1.45, x, 4.5, C.line, 0.75, "dash");
      T(s, q, { x: x + 0.02, y: 1.2, w: 3 * mw - 0.04, h: 0.22, fontSize: 8.5, bold: true, color: C.muted, align: "center" });
    });
    const ph = [
      ["Phase 0  Discovery and plan", "Survey, diagnosis, PRDs and business plan", 0, 5, C.navy, "Oct to Feb"],
      ["Phase 1  MVP pilot", "Version 0.1 with 5 TN customer farms", 5, 11, C.magenta, "Mar to Aug 27"],
      ["Phase 2  First wave", "Version 0.2 with about 20 farms", 11, 15, C.tealDark, "Sep to Dec 27"],
      ["Phase 3  Expansion", "Version 0.3 with 50+ farms and trained technicians", 15, 21, C.teal, "Jan to Jun 28"],
      ["Phase 4  Convergence", "Mature farms move to full TN Manager", 21, 24, C.aqua, "Jul 28 on"],
    ];
    ph.forEach((p, i) => {
      const y = 1.55 + i * 0.58;
      T(s, p[0], { x: 0.5, y: y - 0.02, w: 2.45, h: 0.22, fontSize: 9.5, bold: true, color: p[4] });
      T(s, p[1], { x: 0.5, y: y + 0.2, w: 2.45, h: 0.3, fontSize: 8.5 });
      const bx = x0 + p[2] * mw, bw = (p[3] - p[2]) * mw - 0.04;
      RR(s, { x: bx, y, w: bw, h: 0.4, fill: { color: p[4] }, rectRadius: 0.06 });
      T(s, p[5], { x: bx + 0.03, y, w: bw - 0.06, h: 0.4, fontSize: 8.5, bold: true, color: C.white, valign: "middle", align: "center" });
    });
    [["G0", 5], ["G1", 11], ["G2", 15], ["G3", 21]].forEach((g) => {
      const x = x0 + g[1] * mw;
      s.addShape(pres.shapes.DIAMOND, { x: x - 0.14, y: 4.52, w: 0.28, h: 0.28, fill: { color: C.lime }, line: { type: "none" } });
      T(s, g[0], { x: x - 0.25, y: 4.81, w: 0.5, h: 0.2, fontSize: 9, bold: true, color: C.navy, align: "center" });
    });
    T(s, "Go/no-go decisions", { x: 0.5, y: 4.55, w: 2.45, h: 0.25, fontSize: 9.5, bold: true, color: C.navy });
    s.addNotes("Roadmap alvo (3 min). Tema 4 da pauta. Explique assim: o cronograma mostra as tarefas do projeto semana a semana; o roadmap mostra aonde o PRODUTO vai chegar e em que ordem, em meses e anos. Cada barra é uma fase com um tamanho de fazenda: 5, depois 20, depois 50 ou mais. Os losangos verdes (G0 a G3) são pontos de decisão: só avançamos se a fase anterior bater as metas (detalhe no anexo). O piloto começa em março porque muitos confinamentos pequenos e médios só enchem no período seco.");
  }

  // 12. Go-to-market
  {
    const s = content("Go-to-market: technician-led, inside-out", "Start with who already trusts us, then win back, then attract new customers.");
    const steps = [["1", "Land", "Existing TN customers", "5 pilot farms, onboarded by their own TN technician"], ["2", "Win back", "Former customers and detractors", "Reopen the conversation with proof from the pilot"], ["3", "Attract", "Non-customers", "App plus technical service as the entry door"]];
    steps.forEach((p, i) => {
      const x = 0.5 + i * 3.1, cols = [C.navy, C.magenta, C.teal];
      R(s, { x, y: 1.3, w: 2.75, h: 1.25, fill: { color: cols[i] } });
      T(s, p[0], { x: x + 0.15, y: 1.38, w: 0.5, h: 0.55, fontSize: 28, bold: true, color: C.lime });
      T(s, p[1], { x: x + 0.65, y: 1.42, w: 2.0, h: 0.3, fontSize: 15, bold: true, color: C.white });
      T(s, p[2], { x: x + 0.65, y: 1.72, w: 2.0, h: 0.25, fontSize: 10, italic: true, color: C.white });
      T(s, p[3], { x: x + 0.15, y: 2.02, w: 2.5, h: 0.5, fontSize: 9.5, color: C.white });
      if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.78, y: 1.8, w: 0.28, h: 0.25, fill: { color: C.line }, line: { type: "none" } });
    });
    const cards = [
      ["Channel", "TN technicians and RTVs onboard and follow up each farm. v1.0 premise: 1 technician for 5 farms."],
      ["Offer", "TN Manager Web, technical follow-up and a weekly report, bundled with TN nutrition."],
      ["Pricing", "No license fee for the farmer. Cost covered in product price (v1.0 modeled a 3.8% markup). Final model in the monetization step."],
      ["Proof", "Pilot results become case studies, shown in technical events and referrals between farmers."],
    ];
    cards.forEach((c, i) => {
      const x = 0.5 + i * 2.27, w = 2.15;
      R(s, { x, y: 2.8, w, h: 1.9, fill: { color: C.bg } });
      R(s, { x, y: 2.8, w, h: 0.06, fill: { color: C.lime } });
      T(s, c[0], { x: x + 0.12, y: 2.95, w: w - 0.24, h: 0.3, fontSize: 13, bold: true, color: C.navy });
      T(s, c[1], { x: x + 0.12, y: 3.3, w: w - 0.24, h: 1.35, fontSize: 10 });
    });
    s.addNotes("Go-to-market (3 min). Tema 5 da pauta. GTM é como o produto chega ao cliente: para quem vendemos primeiro, por qual canal, com qual oferta e como cobramos. Ordem: primeiro clientes atuais (menor risco), depois recuperar quem saiu, depois atrair novos. O canal é o técnico TN, nosso diferencial frente à Cargill. Preço: a proposta é não cobrar licença e embutir o custo no produto, como a v1.0 previa. Deixe claro que o modelo final de cobrança será fechado na etapa de monetização, em janeiro.");
  }

  // 13. Validate before scaling
  {
    const s = content("Validate before scaling: research plan and timeline", "Three research fronts feed the diagnosis on December 7 to 11.");
    const fr = [
      { t: "Internal alignment", sub: "Stakeholder interviews", ic: I.hand, col: C.navy, rows: [["Goal", "Measure understanding and support among those who can speed up or block the project"], ["Who", "Beef leadership, sales management, technicians, RTVs, IT and marketing"], ["When", "Until October 23"]] },
      { t: "External benchmark", sub: "Competing systems", ic: I.search, col: C.tealDark, rows: [["Goal", "Strengths and weaknesses of systems in the market"], ["Who", "Audit firms using several systems and consultants with their own system"], ["When", "Until October 23"]] },
      { t: "Customer discovery", sub: "Including detractors", ic: I.clip, col: C.magenta, rows: [["Goal", "Real perceived value and pains before building more"], ["Who", "Customers, former customers and non-customers. Minimum 30 feedlots"], ["When", "Form review with a partner consultant until Nov 6. Answers Nov 9 to 27. Results Dec 4"]] },
    ];
    fr.forEach((f, i) => {
      const x = 0.5 + i * 3.05, w = 2.85;
      R(s, { x, y: 1.25, w, h: 0.72, fill: { color: f.col } });
      OV(s, { x: x + 0.12, y: 1.36, w: 0.5, h: 0.5, fill: { color: C.white, transparency: 80 } });
      s.addImage({ data: f.ic, x: x + 0.24, y: 1.48, w: 0.26, h: 0.26 });
      T(s, f.t, { x: x + 0.72, y: 1.33, w: w - 0.8, h: 0.3, fontSize: 13, bold: true, color: C.white });
      T(s, f.sub, { x: x + 0.72, y: 1.62, w: w - 0.8, h: 0.25, fontSize: 9.5, italic: true, color: C.white });
      R(s, { x, y: 1.97, w, h: 2.78, fill: { color: C.bg } });
      f.rows.forEach((r, j) => {
        const y = 2.07 + j * 0.88;
        T(s, r[0].toUpperCase(), { x: x + 0.14, y, w: w - 0.28, h: 0.2, fontSize: 8, bold: true, color: f.col, charSpacing: 1 });
        T(s, r[1], { x: x + 0.14, y: y + 0.21, w: w - 0.28, h: 0.62, fontSize: 9.5 });
      });
    });
    s.addNotes("Plano de validação (1 min). Mostre que tudo que foi dito até aqui é hipótese e será testado. Datas do cronograma revisado: formulário revisado com a consultora até 06/11, respostas de 09/11 a 27/11, resultado em 04/12 e diagnóstico de 07 a 11/12.");
  }

  // 14. Decisions
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.lime } });
    T(s, "Decisions and next steps", { x: 0.6, y: 0.35, w: 9, h: 0.55, fontSize: 26, bold: true, color: C.white });
    T(s, "What we need to agree today to keep the schedule.", { x: 0.6, y: 0.9, w: 9, h: 0.3, fontSize: 13, italic: true, color: "C9D3EA" });
    T(s, "DECISIONS REQUESTED", { x: 0.6, y: 1.5, w: 4.3, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 2 });
    ["Approve TN Manager Web as the first v2.0 delivery", "Approve the research plan and the 30-feedlot minimum", "Name an internal sponsor and focal points in sales and technical"].forEach((d, i) => {
      const y = 1.9 + i * 0.85;
      OV(s, { x: 0.6, y, w: 0.42, h: 0.42, fill: { color: C.magenta } });
      s.addImage({ data: I.check, x: 0.71, y: y + 0.11, w: 0.2, h: 0.2 });
      T(s, d, { x: 1.15, y: y - 0.02, w: 3.7, h: 0.6, fontSize: 13, color: C.white, valign: "middle" });
    });
    R(s, { x: 5.3, y: 1.45, w: 4.2, h: 3.0, fill: { color: C.navyDark } });
    T(s, "NEXT MILESTONES", { x: 5.5, y: 1.6, w: 3.8, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 2 });
    [["Oct 23", "Interviews and benchmark done"], ["Nov 6", "Survey form validated"], ["Nov 27", "30+ feedlots answered"], ["Dec 11", "Diagnosis and go/no-go (G0)"], ["Feb 5", "Business plan and GTM"]].forEach((x, i) => {
      const y = 1.98 + i * 0.47;
      T(s, x[0], { x: 5.5, y, w: 0.85, h: 0.4, fontSize: 12, bold: true, color: C.lime, valign: "middle" });
      T(s, x[1], { x: 6.4, y, w: 2.95, h: 0.4, fontSize: 11, color: C.white, valign: "middle" });
    });
    s.addImage({ data: WAVE, x: 5.0, y: 4.81, w: 5.0, h: 0.8125 });
    s.addNotes("Fechamento (1 min) e abertura da discussão. Peça as 3 decisões em voz alta. Depois abra: o que falta para você aprovar? Use os slides do anexo se perguntarem sobre metas ou riscos.");
  }

  // Appendix divider-less: A1 gates
  {
    const s = content("Appendix A. Go/no-go criteria between phases", "Proposed targets for discussion. No phase moves forward on the calendar alone.");
    const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: 11 } });
    const G = (t, c) => ({ text: t, options: { bold: true, color: c, fontSize: 11 } });
    const rows = [
      [H("Decision"), H("When"), H("Criteria to move forward (proposal)")],
      [G("G0  Go to pilot", C.navy), "February 2027", "30+ feedlots answered; customer profile and top 5 pains confirmed; internal sponsor named; business plan approved"],
      [G("G1  Pilot validated", C.magenta), "August 2027", "Data entered on at least 70% of days for 8 weeks; up to 10 minutes a day; manager uses the weekly report"],
      [G("G2  Scale approved", C.tealDark), "December 2027", "80% of farms active after 90 days; average recommendation score of 8 or more; each technician serves 5 farms"],
      [G("G3  Integration", C.teal), "June 2028", "At least 5 farms asking for scale or mixer integration; cost per customer below third-party license (BRL 13,541 per year in v1.0)"],
    ];
    s.addTable(rows, { x: 0.5, y: 1.3, w: 9, colW: [2.2, 1.45, 5.35], fontFace: FONT, fontSize: 10.5, color: C.text, border: { type: "solid", pt: 0.75, color: C.line }, rowH: [0.36, 0.7, 0.7, 0.7, 0.7], valign: "middle", margin: [0.05, 0.1, 0.05, 0.1] });
    s.addNotes("Anexo: use só se o chefe perguntar como vamos saber se o piloto deu certo. São metas propostas, abertas para discussão.");
  }

  // A2 risks
  {
    const s = content("Appendix B. Main risks and mitigation", "The biggest v2.0 risks sit in field adoption.");
    const rk = [
      ["Low adoption of manual entry", "Few fields per screen, 10-minute daily target and the TN technician supporting the first weeks."],
      ["Biased survey sample", "Minimum quota of former customers and non-customers. Results guide decisions only with 30 valid answers."],
      ["Seen as a lesser version", "Position TN Manager Web as a step, with a clear migration path to the full version."],
      ["TN technical team capacity", "Pilot limited to 5 farms and 1 technician per 5 farms, reviewed at G2."],
    ];
    rk.forEach((r, i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.75;
      R(s, { x, y, w: 4.4, h: 1.6, fill: { color: C.bg } });
      R(s, { x, y, w: 0.07, h: 1.6, fill: { color: C.magenta } });
      s.addImage({ data: I.warn, x: x + 0.25, y: y + 0.2, w: 0.32, h: 0.3 });
      T(s, r[0], { x: x + 0.7, y: y + 0.18, w: 3.55, h: 0.35, fontSize: 13, bold: true, color: C.navy, valign: "middle" });
      T(s, "MITIGATION", { x: x + 0.7, y: y + 0.62, w: 3.5, h: 0.2, fontSize: 8, bold: true, color: C.magenta, charSpacing: 1 });
      T(s, r[1], { x: x + 0.7, y: y + 0.84, w: 3.55, h: 0.7, fontSize: 11.5 });
    });
    s.addNotes("Anexo: riscos. Use se a discussão for para o lado do que pode dar errado.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
