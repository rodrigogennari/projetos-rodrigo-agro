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
  pres.title = "Business Case v2.0 TN Manager e TN Manager Web";

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

  // 1. Capa
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.lime } });
    T(s, "BUSINESS CASE v2.0", { x: 0.7, y: 0.9, w: 6, h: 0.3, fontSize: 13, bold: true, color: C.lime, charSpacing: 4 });
    T(s, "TN Manager e TN Manager Web", { x: 0.7, y: 1.3, w: 8.5, h: 0.75, fontSize: 38, bold: true, color: C.white });
    T(s, "Onde o produto gera valor e para quem: revisão do cliente-alvo, proposta de valor, produto mínimo e roadmap macro", { x: 0.7, y: 2.15, w: 7.4, h: 0.8, fontSize: 16, italic: true, color: "D6DEF0" });
    T(s, [{ text: "Trouw Nutrition Brasil", options: { bold: true, breakLine: true } }, { text: "Fase 2 do projeto TN Manager: Base Estratégica  |  Outubro de 2026" }], { x: 0.7, y: 3.45, w: 7, h: 0.6, fontSize: 13, color: C.white });
    s.addImage({ data: WAVE, x: 3.5, y: 4.5, w: 6.5, h: 1.125 });
  }

  // 2. Resumo executivo
  {
    const s = content("Resumo executivo", "Recomendação: dar um passo atrás, validar com poucas fazendas e só então ampliar.");
    R(s, { x: 0.5, y: 1.3, w: 3.6, h: 3.35, fill: { color: C.navy } });
    T(s, "RECOMENDAÇÃO", { x: 0.75, y: 1.5, w: 3.1, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 3 });
    T(s, "Construir primeiro o TN Manager Web: uma versão simples, de lançamento manual, testada com 3 a 5 fazendas piloto. Cada nova versão só avança quando a anterior provar uso e valor.", { x: 0.75, y: 1.85, w: 3.1, h: 1.8, fontSize: 15, color: C.white });
    T(s, "Antes de construir: pesquisa com no mínimo 30 fazendas, entrevistas internas e benchmark.", { x: 0.75, y: 3.75, w: 3.1, h: 0.7, fontSize: 11, italic: true, color: "C9D3EA" });
    const pts = [
      ["O problema de negócio continua", "Confinamento responde por 19% do Sales Profit Beef e a v1.0 registrou 51 mil cabeças perdidas para concorrentes com melhor suporte digital."],
      ["O cliente-alvo precisa ser revisto", "A v1.0 mirou grandes confinamentos já automatizados, onde Cargill e DSM/Prodap são mais fortes. A hipótese da v2.0 é que a lacuna maior está em quem ainda registra em papel ou planilha."],
      ["Validar antes de investir", "O escopo final do produto mínimo sai do discovery com clientes, ex-clientes e não clientes."],
    ];
    pts.forEach((p, i) => {
      const y = 1.3 + i * 1.15;
      OV(s, { x: 4.45, y, w: 0.5, h: 0.5, fill: { color: C.magenta } });
      T(s, String(i + 1), { x: 4.45, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[0], { x: 5.15, y: y - 0.02, w: 4.35, h: 0.3, fontSize: 14, bold: true, color: C.navy });
      T(s, p[1], { x: 5.15, y: y + 0.3, w: 4.35, h: 0.75, fontSize: 11 });
    });
  }

  // 3. v1.0 x v2.0
  {
    const s = content("O que muda da v1.0 para a v2.0", "A pergunta deixa de ser como defender a base e passa a ser onde o produto gera valor e para quem.");
    const H = (t, fill) => ({ text: t, options: { bold: true, color: C.white, fill: { color: fill }, fontSize: 11 } });
    const L = (t) => ({ text: t, options: { bold: true, color: C.navy, fontSize: 10.5 } });
    const rows = [
      [H("", C.navy), H("v1.0 (outubro de 2023)", C.tealDark), H("v2.0 (outubro de 2026)", C.magenta)],
      [L("Cliente-alvo"), "Confinamentos grandes (A+ e B+), mínimo de 2.000 cabeças, com pesagem e controle estruturado", "Confinamentos com gestão manual ou em planilha, clientes e não clientes (hipótese a validar)"],
      [L("Produto"), "App completo em nuvem, integração com hardware e terminal offline", "TN Manager Web: lançamento manual, poucos indicadores, evolução por versões"],
      [L("Base da decisão"), "Premissas internas e lista de clientes perdidos", "Entrevistas internas, benchmark externo e discovery com 30 ou mais fazendas"],
      [L("Implantação"), "50 clientes no ano 1 e 200 no ano 5", "3 a 5 fazendas piloto, depois ondas com critério de avanço"],
      [L("Métrica de sucesso"), "Sales Profit e market share", "Uso recorrente e valor percebido primeiro; Sales Profit em seguida"],
    ];
    s.addTable(rows, { x: 0.5, y: 1.3, w: 9, colW: [1.7, 3.65, 3.65], fontFace: FONT, fontSize: 11, color: C.text, border: { type: "solid", pt: 0.75, color: C.line }, rowH: [0.4, 0.68, 0.58, 0.58, 0.5, 0.5], valign: "middle", margin: [0.04, 0.08, 0.04, 0.08] });
  }

  // 4. Ponto de partida
  {
    const s = content("O problema de negócio da v1.0 continua válido", "Os números de 2023 justificam o projeto. A v2.0 revisa a forma de resolver o problema.", "Fonte: Business Case Feedlot App v1.0 (out/2023) e OGP Brazil 2023. Atualizar no diagnóstico.");
    T(s, "Sales Profit Beef 2022 (R$ milhões)", { x: 0.5, y: 1.25, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: C.navy });
    s.addChart(pres.charts.BAR, [{ name: "SP", labels: ["Top 100 confinamentos", "Pareto 80% (265 contas)", "Confinamento", "Total Beef"], values: [6.5, 10.5, 13, 68] }], {
      x: 0.4, y: 1.55, w: 5.4, h: 3.05, barDir: "bar", chartColors: [C.magenta], barGapWidthPct: 45,
      catAxisLabelColor: C.text, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, valAxisHidden: true,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showValue: true, dataLabelPosition: "outEnd",
      dataLabelColor: C.text, dataLabelFontSize: 10, dataLabelFormatCode: "0.0", showLegend: false, valAxisMaxVal: 80,
    });
    const stats = [["19%", "do Sales Profit Beef vem de confinamento", C.magenta], ["51 mil", "cabeças perdidas em 7 clientes, a maior parte para a Cargill", C.navy], ["12,5 mil", "cabeças em risco de perda em 2023", C.tealDark], ["9,6%", "de market share em premix, mineral e concentrado para confinamento", C.teal]];
    stats.forEach((st, i) => {
      const x = 6.1 + (i % 2) * 1.75, y = 1.3 + Math.floor(i / 2) * 1.62;
      R(s, { x, y, w: 1.6, h: 1.45, fill: { color: C.bg } });
      R(s, { x, y, w: 0.06, h: 1.45, fill: { color: st[2] } });
      T(s, st[0], { x: x + 0.17, y: y + 0.12, w: 1.38, h: 0.5, fontSize: 24, bold: true, color: st[2] });
      T(s, st[1], { x: x + 0.17, y: y + 0.65, w: 1.38, h: 0.75, fontSize: 9.5 });
    });
  }

  // 5. Mapa mental de valor
  {
    const s = content("Mapa de valor: quem ganha o quê com o TN Manager Web", "Quatro públicos, quatro tipos de valor. Todos são hipóteses a validar no discovery.");
    const cx = 5.0, cy = 2.85;
    const br = [
      { t: "Gestor do confinamento", ic: I.user, col: C.navy, bx: 2.1, by: 2.05, leaves: ["Custo e desempenho por lote sem montar planilha", "Decidir com o dado da semana, não do mês"], lx: 0.4, ly: 1.25 },
      { t: "Equipe do curral", ic: I.hat, col: C.tealDark, bx: 2.1, by: 3.65, leaves: ["Lançar leitura de cocho e trato em poucos minutos", "Tela simples, no celular ou no computador"], lx: 0.4, ly: 3.65 },
      { t: "Técnico TN", ic: I.tech, col: C.teal, bx: 7.9, by: 2.05, leaves: ["Acompanhar várias fazendas à distância", "Recomendar ajuste de dieta com dado real"], lx: 7.1, ly: 1.25, right: true },
      { t: "Trouw Nutrition", ic: I.bldg, col: C.magenta, bx: 7.9, by: 3.65, leaves: ["Fidelizar clientes de núcleo e premix", "Porta de entrada para não clientes"], lx: 7.1, ly: 3.65, right: true },
    ];
    br.forEach((b) => LN(s, cx, cy, b.bx, b.by, b.col, 2));
    OV(s, { x: cx - 1.0, y: cy - 0.55, w: 2.0, h: 1.1, fill: { color: C.navy } });
    T(s, "TN Manager Web", { x: cx - 1.0, y: cy - 0.55, w: 2.0, h: 1.1, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle" });
    br.forEach((b) => {
      OV(s, { x: b.bx - 0.3, y: b.by - 0.3, w: 0.6, h: 0.6, fill: { color: b.col } });
      s.addImage({ data: b.ic, x: b.bx - 0.15, y: b.by - 0.15, w: 0.3, h: 0.3 });
            T(s, b.t, { x: b.right ? b.bx - 2.1 : b.bx + 0.38, y: b.by - 0.13, w: 1.72, h: 0.26, fontSize: 11.5, bold: true, color: b.col, align: b.right ? "right" : "left" });
      const leafY = b.by < 3 ? 1.18 : 4.02;
      const leafX = b.right ? 6.9 : 0.45;
      b.leaves.forEach((lf, j) => {
        RR(s, { x: leafX + j * 1.35, y: leafY, w: 1.28, h: 0.55, fill: { color: C.bg }, line: { color: b.col, width: 0.75 } });
        T(s, lf, { x: leafX + j * 1.35 + 0.06, y: leafY + 0.02, w: 1.16, h: 0.51, fontSize: 8.5, valign: "middle", align: "center" });
      });
    });
  }

  // 6. Segmentação 2x2
  {
    const s = content("Revisão do cliente-alvo: onde o TN Manager Web entra", "A v1.0 mirou o quadrante mais disputado. A v2.0 propõe começar onde há dor e pouca oferta simples.");
    const gx = 1.15, gy = 1.25, gw = 6.0, gh = 3.4, hw = gw / 2, hh = gh / 2;
    const q = [
      { x: gx, y: gy, t: "Grande porte, registro manual", d: "Alta dor e alto volume de produto. Segunda onda do TN Manager Web.", fill: C.magLight, col: C.magenta, tg: "PRIORIDADE 2" },
      { x: gx + hw, y: gy, t: "Grande porte, automatizado", d: "Foco da v1.0. Disputado por Cargill e DSM/Prodap. Território do TN Manager completo.", fill: C.bg, col: C.tealDark, tg: "TN MANAGER" },
      { x: gx, y: gy + hh, t: "Médio porte, registro manual", d: "Mais fazendas e menos software simples disponível. Ponto de entrada sugerido para o piloto.", fill: C.magenta, col: C.white, tg: "PILOTO", strong: true },
      { x: gx + hw, y: gy + hh, t: "Médio porte, automatizado", d: "Já usa alguma ferramenta. Avaliar troca ou integração em versão futura.", fill: C.bg, col: C.muted, tg: "AVALIAR" },
    ];
    q.forEach((k) => {
      R(s, { x: k.x + 0.04, y: k.y + 0.04, w: hw - 0.08, h: hh - 0.08, fill: { color: k.fill } });
      tag(s, k.tg, k.x + 0.2, k.y + 0.2, 1.2, k.strong ? C.navy : k.col);
      T(s, k.t, { x: k.x + 0.2, y: k.y + 0.55, w: hw - 0.4, h: 0.3, fontSize: 13, bold: true, color: k.strong ? C.white : C.navy });
      T(s, k.d, { x: k.x + 0.2, y: k.y + 0.88, w: hw - 0.4, h: 0.75, fontSize: 10.5, color: k.strong ? C.white : C.text });
    });
    T(s, "Porte", { x: 0.35, y: gy, w: 0.7, h: 0.3, fontSize: 10, bold: true, color: C.navy });
    T(s, "Grande", { x: 0.35, y: gy + hh / 2, w: 0.75, h: 0.3, fontSize: 10, color: C.muted });
    T(s, "Médio", { x: 0.35, y: gy + hh * 1.5, w: 0.75, h: 0.3, fontSize: 10, color: C.muted });
    T(s, "Registro manual ou planilha", { x: gx, y: gy + gh + 0.06, w: hw, h: 0.25, fontSize: 10, color: C.muted, align: "center" });
    T(s, "Balança, vagão e software", { x: gx + hw, y: gy + gh + 0.06, w: hw, h: 0.25, fontSize: 10, color: C.muted, align: "center" });
    T(s, "Maturidade de gestão de dados", { x: gx, y: gy + gh + 0.3, w: gw, h: 0.25, fontSize: 10, bold: true, color: C.navy, align: "center" });
    R(s, { x: 7.45, y: 1.25, w: 2.1, h: 3.4, fill: { color: C.navy } });
    T(s, "Por que mudar o foco", { x: 7.62, y: 1.4, w: 1.8, h: 0.3, fontSize: 12, bold: true, color: C.lime });
    T(s, [
      { text: "Quem já tem balança e vagão compara o TN Manager com sistemas maduros e equipes de campo maiores.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "Quem registra no papel compara com o caderno. Um produto simples já entrega ganho visível.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "O discovery confirma porte mínimo e tamanho de cada quadrante." },
    ], { x: 7.62, y: 1.78, w: 1.8, h: 2.8, fontSize: 10, color: C.white });
  }

  // 7. Perfil do cliente-alvo
  {
    const s = content("Perfil do cliente-alvo para o piloto", "Critérios propostos para escolher as primeiras fazendas. A faixa de porte sai do discovery.");
    const cols = [
      { t: "Obrigatórios", col: C.navy, items: ["Confinamento ativo, com ciclo regular de entrada e saída", "Cliente TN de núcleo, premix ou mineral", "Registro atual em papel, caderno ou planilha", "Uma pessoa responsável por lançar os dados todo dia", "Técnico TN com visita ou contato recorrente"] },
      { t: "Desejáveis", col: C.teal, items: ["Pesagem na entrada e na saída dos animais", "Leitura de cocho diária", "Internet na sede da fazenda", "Gestor disposto a dar retorno a cada 15 dias"] },
      { t: "Fora do perfil inicial", col: C.magenta, items: ["Operações já integradas a balança, vagão e software", "Confinamentos sem rotina mínima de registro", "Recria a pasto, leite e outras espécies (mantido da v1.0)"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 3.05, w = 2.85;
      R(s, { x, y: 1.3, w, h: 0.45, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.15, y: 1.3, w: w - 0.3, h: 0.45, fontSize: 13, bold: true, color: C.white, valign: "middle" });
      R(s, { x, y: 1.75, w, h: 2.55, fill: { color: C.bg } });
      T(s, c.items.map((it, j) => ({ text: it, options: { bullet: { indent: 12 }, breakLine: j < c.items.length - 1, paraSpaceAfter: 5 } })), { x: x + 0.1, y: 1.88, w: w - 0.22, h: 2.35, fontSize: 10.5 });
    });
    R(s, { x: 0.5, y: 4.42, w: 9.0, h: 0.4, fill: { color: C.magLight } });
    T(s, [{ text: "Porte mínimo: ", options: { bold: true, color: C.magenta } }, { text: "a v1.0 usava 2.000 cabeças. Para a versão manual, a faixa ideal será definida com os dados do discovery." }], { x: 0.65, y: 4.42, w: 8.7, h: 0.4, fontSize: 10.5, valign: "middle" });
  }

  // 8. Canvas de proposta de valor
  {
    const s = content("Proposta de valor: canvas para o TN Manager Web", "Lado esquerdo: o que entregamos. Lado direito: o que o confinamento precisa. Hipóteses para o discovery.");
    const box = (x, y, w, h, title, items, col) => {
      R(s, { x, y, w, h, fill: { color: C.white }, line: { color: col, width: 1 } });
      R(s, { x, y, w, h: 0.3, fill: { color: col } });
      T(s, title, { x: x + 0.1, y, w: w - 0.2, h: 0.3, fontSize: 10.5, bold: true, color: C.white, valign: "middle" });
      T(s, items.map((it, j) => ({ text: it, options: { bullet: { indent: 10 }, breakLine: j < items.length - 1, paraSpaceAfter: 3 } })), { x: x + 0.06, y: y + 0.36, w: w - 0.12, h: h - 0.4, fontSize: 9 });
    };
    T(s, "MAPA DE VALOR", { x: 0.5, y: 1.2, w: 4.2, h: 0.25, fontSize: 10, bold: true, color: C.navy, charSpacing: 2 });
    T(s, "PERFIL DO CLIENTE", { x: 5.3, y: 1.2, w: 4.2, h: 0.25, fontSize: 10, bold: true, color: C.magenta, charSpacing: 2 });
    box(0.5, 1.5, 1.95, 3.2, "Produto e serviço", ["Aplicação web com lançamento manual", "Painel com poucos indicadores", "Relatório semanal automático", "Acompanhamento do técnico TN"], C.navy);
    box(2.55, 1.5, 2.15, 1.55, "Aliviadores de dor", ["Formulários padronizados e curtos", "Sem custo de licença para o cliente", "Suporte de quem já visita a fazenda"], C.tealDark);
    box(2.55, 3.15, 2.15, 1.55, "Criadores de ganho", ["Relatório pronto para o dono", "Recomendação de ajuste de dieta", "Comparativo entre lotes"], C.teal);
    box(5.3, 1.5, 2.15, 1.55, "Dores", ["Dados espalhados em caderno e planilha", "Erro de digitação e retrabalho", "Software caro, complexo ou sem suporte"], C.magenta);
    box(5.3, 3.15, 2.15, 1.55, "Ganhos esperados", ["Ver o desempenho do lote na semana", "Comparar com fazendas parecidas", "Um técnico olhando os números junto"], "A30C5A");
    box(7.55, 1.5, 1.95, 3.2, "Tarefas do cliente", ["Ajustar o trato diário", "Fechar custo por lote e por arroba", "Controlar entradas, saídas e mortes", "Prestar contas ao dono ou sócio"], C.navyDark);
    OV(s, { x: 4.73, y: 2.82, w: 0.55, h: 0.55, fill: { color: C.lime } });
    T(s, "FIT", { x: 4.73, y: 2.82, w: 0.55, h: 0.55, fontSize: 10, bold: true, color: C.navy, align: "center", valign: "middle" });
  }

  // 9. TN Manager x TN Manager Web
  {
    const s = content("Duas versões, dois papéis na estratégia", "O TN Manager Web abre a porta e prepara o cliente para migrar à versão completa.");
    const rows = [
      ["Para quem", "Grandes confinamentos com balança, vagão e equipe estruturada", "Confinamentos com registro manual (hipótese)"],
      ["Entrada de dados", "Integração com hardware e coleta automática", "Lançamento manual em formulários simples"],
      ["Acesso", "Nuvem e terminal offline", "Navegador, no computador ou no celular"],
      ["Escopo", "Gestão completa, integrações e módulos adicionais", "Cinco módulos essenciais"],
      ["Implantação e suporte", "Projeto de implantação com equipe dedicada", "Técnico TN que já atende a fazenda"],
      ["Papel na estratégia", "Defender contas A+ e B+", "Conquistar, ensinar e preparar a migração"],
    ];
    T(s, "TN Manager (escopo da v1.0)", { x: 2.35, y: 1.25, w: 3.5, h: 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", fill: { color: C.tealDark } });
    T(s, "TN Manager Web (proposta v2.0)", { x: 6.0, y: 1.25, w: 3.5, h: 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", fill: { color: C.magenta } });
    rows.forEach((r, i) => {
      const y = 1.75 + i * 0.5;
      if (i % 2 === 0) R(s, { x: 0.5, y, w: 9.0, h: 0.46, fill: { color: C.bg } });
      T(s, r[0], { x: 0.6, y, w: 1.7, h: 0.46, fontSize: 10.5, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.45, y, w: 3.3, h: 0.46, fontSize: 10, valign: "middle" });
      T(s, r[2], { x: 6.1, y, w: 3.3, h: 0.46, fontSize: 10, valign: "middle", bold: true, color: C.navy });
    });
  }

  // 10. MVP MoSCoW
  {
    const s = content("Produto mínimo: o que entra no TN Manager Web", "Priorização MoSCoW. Só a primeira coluna é obrigatória para o piloto.");
    const cols = [
      { t: "Precisa ter", st: "versão 0.1, piloto", col: C.magenta, items: ["Cadastro de fazenda, currais e lotes", "Lançamento diário de trato por curral", "Leitura de cocho", "Entradas, pesagens, saídas, mortes e enfermaria", "Painel: consumo de matéria seca, GMD, dias de cocho e mortalidade", "Relatório semanal em PDF"] },
      { t: "Deveria ter", st: "versão 0.2", col: C.navy, items: ["Cadastro de dietas com correção de matéria seca", "Custo por lote e por arroba produzida", "Alertas simples de consumo fora do previsto"] },
      { t: "Poderia ter", st: "versão 0.3 em diante", col: C.teal, items: ["Importação de planilha existente", "Comparativo entre lotes", "Benchmark anônimo entre fazendas"] },
      { t: "Não terá agora", st: "fica para a versão completa", col: C.muted, items: ["Integração com balança, vagão e hardware", "Aplicativo nativo offline", "Fidelidade e sustentabilidade", "Recria, leite e outras espécies"] },
    ];
    cols.forEach((c, i) => {
      const x = 0.5 + i * 2.27, w = 2.15;
      R(s, { x, y: 1.3, w, h: 0.62, fill: { color: c.col } });
      T(s, c.t, { x: x + 0.12, y: 1.34, w: w - 0.24, h: 0.3, fontSize: 13, bold: true, color: C.white });
      T(s, c.st, { x: x + 0.12, y: 1.62, w: w - 0.24, h: 0.25, fontSize: 9.5, italic: true, color: C.white });
      R(s, { x, y: 1.92, w, h: 2.78, fill: { color: i === 0 ? C.magLight : C.bg } });
      T(s, c.items.map((it, j) => ({ text: it, options: { bullet: { indent: 10 }, breakLine: j < c.items.length - 1, paraSpaceAfter: 4 } })), { x: x + 0.06, y: 2.02, w: w - 0.14, h: 2.6, fontSize: 9.5, color: i === 3 ? C.muted : C.text });
    });
  }

  // 11. Jornada semanal
  {
    const s = content("Como o produto mínimo funciona na rotina da fazenda", "Cinco passos. Se não cabe na rotina do curral, não entra no produto mínimo.");
    const st = [
      ["Cadastrar", "uma vez", "Fazenda, currais, lotes e dieta"],
      ["Registrar", "todo dia", "Trato fornecido e leitura de cocho"],
      ["Movimentar", "quando ocorre", "Entradas, pesagens, saídas e mortes"],
      ["Analisar", "toda semana", "Painel e relatório gerados sozinhos"],
      ["Ajustar", "com o técnico TN", "Revisão dos números e ajuste de dieta"],
    ];
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
    const pr = [["Poucos campos por tela", "Meta de lançamento diário em até 10 minutos."], ["Linguagem do curral", "Termos que o tratador já usa, sem jargão de sistema."], ["Leve para conexão fraca", "Telas simples que carregam com internet de fazenda."]];
    pr.forEach((p, i) => {
      const x = 0.5 + i * 3.05;
      R(s, { x, y: 3.85, w: 2.85, h: 0.85, fill: { color: C.bg } });
      R(s, { x, y: 3.85, w: 0.06, h: 0.85, fill: { color: C.lime } });
      T(s, p[0], { x: x + 0.18, y: 3.93, w: 2.6, h: 0.28, fontSize: 11.5, bold: true, color: C.navy });
      T(s, p[1], { x: x + 0.18, y: 4.22, w: 2.6, h: 0.45, fontSize: 9.5 });
    });
  }

  // 12. Plano de pesquisa
  {
    const s = content("Plano de pesquisa: três frentes antes de construir", "Itens 2.1, 2.1.1 e 2.2 do cronograma, nas semanas 2 e 3.");
    const fr = [
      { t: "Aderência interna", sub: "Entrevista com stakeholders", ic: I.hand, col: C.navy, rows: [["Objetivo", "Medir entendimento e apoio entre quem pode acelerar ou travar o projeto na Trouw"], ["Público", "Diretoria Beef, gerência comercial, técnicos, RTVs, TI e marketing"], ["Método", "Entrevista semiestruturada de 30 a 45 minutos"], ["Saída", "Mapa de aderência e lista de objeções"]] },
      { t: "Benchmark externo", sub: "Sistemas concorrentes", ic: I.search, col: C.tealDark, rows: [["Objetivo", "Pontos fortes e fracos dos sistemas do mercado"], ["Público", "Auditorias que usam vários sistemas e consultores com sistema próprio"], ["Método", "Entrevista com demonstração guiada"], ["Saída", "Matriz de funcionalidades, dores e soluções já testadas"]] },
      { t: "Discovery com clientes", sub: "Incluindo detratores", ic: I.clip, col: C.magenta, rows: [["Objetivo", "Percepção real de valor e dores antes de construir"], ["Público", "Clientes, ex-clientes e não clientes"], ["Método", "Formulário estruturado e entrevista curta"], ["Saída", "Perfil de cliente validado e ranking de dores"]] },
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
        const y = 2.05 + j * 0.67;
        T(s, r[0].toUpperCase(), { x: x + 0.14, y, w: w - 0.28, h: 0.2, fontSize: 8, bold: true, color: f.col, charSpacing: 1 });
        T(s, r[1], { x: x + 0.14, y: y + 0.2, w: w - 0.28, h: 0.45, fontSize: 9.5 });
      });
    });
  }

  // 13. Discovery: amostra e formulário
  {
    const s = content("Discovery: no mínimo 30 fazendas, com detratores na amostra", "Composição proposta para evitar uma amostra feita só de clientes satisfeitos.");
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Amostra", labels: ["Clientes ativos", "Não clientes", "Ex-clientes e detratores"], values: [12, 10, 8] }], {
      x: 0.4, y: 1.25, w: 3.4, h: 3.0, holeSize: 62, chartColors: [C.navy, C.teal, C.magenta],
      showLegend: false, showValue: true, showPercent: false, dataLabelColor: C.white, dataLabelFontSize: 12, dataLabelFontBold: true,
    });
    T(s, "30", { x: 1.6, y: 2.35, w: 1.0, h: 0.5, fontSize: 30, bold: true, color: C.navy, align: "center", valign: "middle" });
    T(s, "fazendas", { x: 1.6, y: 2.82, w: 1.0, h: 0.25, fontSize: 10, color: C.muted, align: "center" });
    [["Clientes ativos", "12", C.navy], ["Não clientes", "10", C.teal], ["Ex-clientes e detratores", "8", C.magenta]].forEach((l, i) => {
      const x = [0.45, 1.5, 2.5][i];
      R(s, { x, y: 4.42, w: 0.13, h: 0.13, fill: { color: l[2] } });
      T(s, l[0], { x: x + 0.18, y: 4.37, w: [1.0, 0.85, 1.4][i], h: 0.25, fontSize: 8.5, valign: "middle" });
    });
    T(s, "Blocos do formulário", { x: 4.2, y: 1.25, w: 5.3, h: 0.3, fontSize: 13, bold: true, color: C.navy });
    const bl = [["Perfil da operação", "Porte, ciclo, equipe e quem lança os dados"], ["Como registra hoje", "Papel, planilha ou software, e quanto tempo gasta"], ["Dores na rotina de dados", "Ranking das cinco maiores dificuldades"], ["Sistemas já usados", "O que funcionou e por que parou de usar"], ["Valor percebido", "Disposição a usar e a recomendar, de 0 a 10"], ["Relação com fornecedores", "Peso do suporte digital na escolha da nutrição"]];
    bl.forEach((b, i) => {
      const x = 4.2 + (i % 2) * 2.68, y = 1.65 + Math.floor(i / 2) * 0.98;
      R(s, { x, y, w: 2.55, h: 0.86, fill: { color: C.bg } });
      T(s, String(i + 1), { x: x + 0.1, y: y + 0.1, w: 0.35, h: 0.4, fontSize: 20, bold: true, color: C.magenta });
      T(s, b[0], { x: x + 0.48, y: y + 0.1, w: 1.98, h: 0.27, fontSize: 10.5, bold: true, color: C.navy });
      T(s, b[1], { x: x + 0.48, y: y + 0.38, w: 1.98, h: 0.45, fontSize: 9 });
    });
  }

  // 14. Mapa de stakeholders
  {
    const s = content("Aderência interna: mapa inicial de stakeholders", "Posições de partida (hipótese). As entrevistas confirmam ou movem cada grupo.");
    const gx = 0.9, gy = 1.25, gw = 5.4, gh = 3.4;
    const qd = [["Manter satisfeito", gx, gy, "EEF1F8"], ["Gerenciar de perto", gx + gw / 2, gy, C.magLight], ["Monitorar", gx, gy + gh / 2, C.bg], ["Manter informado", gx + gw / 2, gy + gh / 2, "E8F4F7"]];
    qd.forEach((q) => {
      R(s, { x: q[1] + 0.03, y: q[2] + 0.03, w: gw / 2 - 0.06, h: gh / 2 - 0.06, fill: { color: q[3] } });
      T(s, q[0].toUpperCase(), { x: q[1] + 0.12, y: q[2] + 0.1, w: 2.4, h: 0.22, fontSize: 8.5, bold: true, color: C.muted, charSpacing: 1 });
    });
    const pts = [["Diretoria Beef", 0.6, 0.84, C.navy], ["Gerência comercial", 0.74, 0.72, C.magenta], ["Finanças", 0.26, 0.8, C.navy], ["TI e Digital", 0.55, 0.62, C.navy], ["Técnicos e RTVs", 0.76, 0.38, C.magenta], ["Software house", 0.7, 0.3, C.teal], ["Marketing", 0.36, 0.22, C.teal]];
    pts.forEach((p) => {
      const x = gx + p[1] * gw, y = gy + (1 - p[2]) * gh;
      OV(s, { x: x - 0.1, y: y - 0.1, w: 0.2, h: 0.2, fill: { color: p[3] } });
      T(s, p[0], { x: x + 0.14, y: y - 0.12, w: 1.4, h: 0.24, fontSize: 9.5, bold: true, color: C.navy, valign: "middle" });
    });
    T(s, "Influência", { x: 0.2, y: gy + gh / 2 - 0.15, w: 0.65, h: 0.3, fontSize: 10, bold: true, color: C.navy });
    T(s, "Interesse e envolvimento no projeto", { x: gx, y: gy + gh + 0.06, w: gw, h: 0.25, fontSize: 10, bold: true, color: C.navy, align: "center" });
    R(s, { x: 6.6, y: 1.25, w: 2.9, h: 3.4, fill: { color: C.navy } });
    T(s, "O que perguntar", { x: 6.78, y: 1.38, w: 2.6, h: 0.3, fontSize: 12, bold: true, color: C.lime });
    const qs = ["O que você entende que é o TN Manager hoje?", "Que problema de negócio ele precisa resolver?", "O que faria você apoiar ou frear o projeto?", "Quem precisa estar convencido para ele avançar?", "Como medir sucesso em 12 meses?"];
    T(s, qs.map((q, j) => ({ text: q, options: { bullet: { type: "number" }, breakLine: j < qs.length - 1, paraSpaceAfter: 6 } })), { x: 6.75, y: 1.75, w: 2.65, h: 2.8, fontSize: 10, color: C.white });
  }

  // 15. Roadmap macro
  {
    const s = content("Roadmap macro: evoluir por versões, com critério para avançar", "Prazos indicativos, contados a partir do início do projeto. Cada fase fecha com uma decisão de avanço.");
    const x0 = 3.05, x1 = 9.5, months = 18, mw = (x1 - x0) / months;
    [3, 6, 9, 12, 15, 18].forEach((m) => {
      const x = x0 + m * mw;
      LN(s, x, 1.45, x, 4.55, C.line, 0.75, "dash");
      T(s, "Mês " + m, { x: x - 0.4, y: 1.2, w: 0.8, h: 0.22, fontSize: 8.5, color: C.muted, align: "center" });
    });
    const ph = [
      ["Fase 0  Diagnóstico e pesquisa", "Business case, discovery, SWOT, roadmap e GTM", 0, 3, C.navy],
      ["Fase 1  Produto mínimo e piloto", "Versão 0.1 com 3 a 5 fazendas", 3, 6, C.magenta],
      ["Fase 2  Ajuste e primeira onda", "Versão 0.2 com 10 a 15 fazendas", 6, 9, C.tealDark],
      ["Fase 3  Expansão", "Versão 0.3 com 30 ou mais fazendas e técnicos treinados", 9, 15, C.teal],
      ["Fase 4  Convergência", "Migração para o TN Manager completo de quem amadurecer", 15, 18, C.aqua],
    ];
    ph.forEach((p, i) => {
      const y = 1.55 + i * 0.6;
      T(s, p[0], { x: 0.5, y: y - 0.02, w: 2.45, h: 0.22, fontSize: 9.5, bold: true, color: p[4] });
      T(s, p[1], { x: 0.5, y: y + 0.2, w: 2.45, h: 0.3, fontSize: 8.5, color: C.text });
      const bx = x0 + p[2] * mw, bw = (p[3] - p[2]) * mw - 0.05;
      RR(s, { x: bx, y, w: bw, h: 0.42, fill: { color: p[4] }, rectRadius: 0.06 });
      T(s, p[2] === 15 ? "Mês 16 em diante" : "Meses " + (p[2] + 1) + " a " + p[3], { x: bx + 0.05, y, w: bw - 0.1, h: 0.42, fontSize: 9, bold: true, color: C.white, valign: "middle", align: "center" });
    });
    const gates = [["G0", 3], ["G1", 6], ["G2", 9], ["G3", 15]];
    gates.forEach((g) => {
      const x = x0 + g[1] * mw;
      s.addShape(pres.shapes.DIAMOND, { x: x - 0.15, y: 4.6, w: 0.3, h: 0.3, fill: { color: C.lime }, line: { type: "none" } });
      T(s, g[0], { x: x - 0.25, y: 4.9, w: 0.5, h: 0.2, fontSize: 9, bold: true, color: C.navy, align: "center" });
    });
    T(s, "Decisões de avanço", { x: 0.5, y: 4.62, w: 2.45, h: 0.25, fontSize: 9.5, bold: true, color: C.navy });
  }

  // 16. Gates
  {
    const s = content("Critérios de avanço entre fases", "Metas propostas para discussão. Nenhuma fase avança só pelo calendário.");
    const H = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy }, fontSize: 11 } });
    const G = (t, c) => ({ text: t, options: { bold: true, color: c, fontSize: 11 } });
    const rows = [
      [H("Decisão"), H("Quando"), H("Critério para avançar (proposta)")],
      [G("G0  Escopo do produto mínimo", C.navy), "Fim da fase 0", "30 ou mais fazendas respondidas; perfil de cliente e cinco dores principais confirmados; patrocinador interno definido"],
      [G("G1  Piloto validado", C.magenta), "Fim da fase 1", "Lançamento em pelo menos 70% dos dias, por 8 semanas; até 10 minutos por dia; gestor usa o relatório semanal"],
      [G("G2  Escala aprovada", C.tealDark), "Fim da fase 2", "80% das fazendas ativas após 90 dias; nota de recomendação média de 8 ou mais; cada técnico atende 5 fazendas (premissa da v1.0)"],
      [G("G3  Integração", C.teal), "Durante a fase 3", "Pelo menos 5 fazendas pedindo integração com balança ou vagão; custo por cliente abaixo da licença de terceiros (R$ 13.541 por ano na v1.0)"],
    ];
    s.addTable(rows, { x: 0.5, y: 1.3, w: 9, colW: [2.4, 1.35, 5.25], fontFace: FONT, fontSize: 10.5, color: C.text, border: { type: "solid", pt: 0.75, color: C.line }, rowH: [0.36, 0.7, 0.7, 0.7, 0.7], valign: "middle", margin: [0.05, 0.1, 0.05, 0.1] });
  }

  // 17. Riscos
  {
    const s = content("Riscos principais e como reduzir cada um", "Os maiores riscos da v2.0 estão na adoção pelo campo.");
    const rk = [
      ["Baixa adesão ao lançamento manual", "Poucos campos por tela, meta de 10 minutos por dia e o técnico TN como apoio nas primeiras semanas."],
      ["Amostra do discovery enviesada", "Cota mínima de ex-clientes e não clientes. Resultado só orienta decisão com 30 respostas válidas."],
      ["Percepção de versão inferior", "Posicionar o TN Manager Web como degrau, com caminho claro de migração para a versão completa."],
      ["Capacidade do time técnico TN", "Piloto limitado a 3 a 5 fazendas e premissa de 5 fazendas por técnico, revista no G2."],
    ];
    rk.forEach((r, i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.75;
      R(s, { x, y, w: 4.4, h: 1.6, fill: { color: C.bg } });
      R(s, { x, y, w: 0.07, h: 1.6, fill: { color: C.magenta } });
      s.addImage({ data: I.warn, x: x + 0.25, y: y + 0.2, w: 0.32, h: 0.3 });
      T(s, r[0], { x: x + 0.7, y: y + 0.18, w: 3.55, h: 0.35, fontSize: 13, bold: true, color: C.navy, valign: "middle" });
      T(s, "MITIGAÇÃO", { x: x + 0.7, y: y + 0.62, w: 3.5, h: 0.2, fontSize: 8, bold: true, color: C.magenta, charSpacing: 1 });
      T(s, r[1], { x: x + 0.7, y: y + 0.84, w: 3.55, h: 0.7, fontSize: 11.5 });
    });
  }

  // 18. Decisões e próximos passos
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: C.lime } });
    T(s, "Decisões e próximos passos", { x: 0.6, y: 0.35, w: 9, h: 0.55, fontSize: 26, bold: true, color: C.white });
    T(s, "O que precisamos aprovar hoje para iniciar o discovery na próxima semana.", { x: 0.6, y: 0.9, w: 9, h: 0.3, fontSize: 13, italic: true, color: "C9D3EA" });
    T(s, "DECISÕES SOLICITADAS", { x: 0.6, y: 1.5, w: 4.3, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 2 });
    const ds = ["Aprovar o TN Manager Web como primeira entrega da v2.0", "Validar o plano de pesquisa e a meta mínima de 30 fazendas", "Indicar patrocinador interno e pontos focais nas áreas comercial e técnica"];
    ds.forEach((d, i) => {
      const y = 1.9 + i * 0.85;
      OV(s, { x: 0.6, y, w: 0.42, h: 0.42, fill: { color: C.magenta } });
      s.addImage({ data: I.check, x: 0.71, y: y + 0.11, w: 0.2, h: 0.2 });
      T(s, d, { x: 1.15, y: y - 0.02, w: 3.7, h: 0.6, fontSize: 13, color: C.white, valign: "middle" });
    });
    R(s, { x: 5.3, y: 1.45, w: 4.2, h: 3.0, fill: { color: C.navyDark } });
    T(s, "PRÓXIMAS SEMANAS", { x: 5.5, y: 1.6, w: 3.8, h: 0.25, fontSize: 11, bold: true, color: C.lime, charSpacing: 2 });
    const nx = [["S2", "Entrevistas internas e benchmark externo"], ["S2 e S3", "Aplicação do formulário nas 30 fazendas"], ["S3", "Consolidação do discovery e revisão do perfil de cliente"], ["S4", "Imersão técnica com a software house"]];
    nx.forEach((x, i) => {
      const y = 2.0 + i * 0.58;
      T(s, x[0], { x: 5.5, y, w: 0.85, h: 0.45, fontSize: 13, bold: true, color: C.lime, valign: "middle" });
      T(s, x[1], { x: 6.4, y, w: 2.95, h: 0.45, fontSize: 11, color: C.white, valign: "middle" });
    });
    s.addImage({ data: WAVE, x: 5.0, y: 4.81, w: 5.0, h: 0.8125 });
  }

  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
