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

  // 1 Agenda
  {
    const s = content("Agenda", "O que vamos apresentar e decidir hoje");
    const items = [
      ["Resumo executivo e aderência estratégica", "A recomendação em uma página"],
      ["Cliente-alvo e o que muda da v1.0", "Confinamentos de 200 a 2.000 cabeças"],
      ["Mercado: estamos no caminho certo?", "Censo de Confinamento 2025"],
      ["Proposta de valor e objetivos", "Quem ganha o quê com o TN Manager Web"],
      ["Nosso diferencial: a entrega técnica", "O que fazemos com os dados do cliente"],
      ["Produto mínimo e como funciona na prática", "Rotina, escore de cocho, internet e arquitetura"],
      ["Pesquisa com 30 fazendas", "Plano e composição da amostra"],
      ["Concorrência", "Cargill, DSM e Tecnobeef"],
      ["Custo-benefício e cobrança", "20 licenças pagas hoje e o que muda"],
      ["SWOT e go-to-market", "Da v1.0 para a Web; 20 fazendas e analistas"],
      ["Roadmap de 12 meses e cronograma de 4 anos", "Datas, responsáveis e critérios de avanço"],
      ["Riscos, decisões e próximos passos", "O que precisamos aprovar hoje"],
    ];
    items.forEach((it, i) => {
      const col = Math.floor(i / 6), row = i % 6;
      const x = 0.55 + col * 4.6, y = 1.3 + row * 0.56;
      OV(s, { x, y, w: 0.4, h: 0.4, fill: { color: i === 4 ? C.navy : C.magenta } });
      T(s, String(i + 1), { x, y, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, it[0], { x: x + 0.52, y: y - 0.03, w: 3.9, h: 0.25, fontSize: 12, bold: true, color: C.navy });
      T(s, it[1], { x: x + 0.52, y: y + 0.21, w: 3.9, h: 0.2, fontSize: 9.5, color: C.text });
    });
    LN(s, 5.0, 1.3, 5.0, 4.5, C.line, 0.75);
    s.addNotes("Roteiro sugerido para a diretoria: itens 1 a 4 em 8 minutos; itens 5 e 6 são o coração da conversa (o diferencial é a entrega técnica e como funciona na prática); itens 7 a 11 em 12 minutos; reservar 10 minutos para o item 12.");
  }

  // 2 Diferencial
  {
    const s = content("Nosso diferencial é a entrega técnica, não a ferramenta", "Software de gestão virou item de prateleira. O cliente paga pela decisão que sai dos dados.");
    const lv = [
      ["1", "Ferramenta", "Lançamento, painel e relatório. Cargill, DSM, TGC e outros já oferecem.", "D5DAE0", C.text, "Item de prateleira"],
      ["2", "Dados confiáveis", "Lançamento diário conferido por auditoria mensal da rotina.", C.tealDark, C.white, ""],
      ["3", "Análise técnica TN", "O analista lê os números toda semana e compara com fazendas parecidas (benchmark).", C.navy, C.white, ""],
      ["4", "Decisão na fazenda", "Ajuste de dieta, trato, lotação, compra e venda de animais.", C.magenta, C.white, "Onde está o nosso valor"],
    ];
    lv.forEach((l, i) => {
      const y = 1.3 + i * 0.84, w = 3.6 + i * 0.3, x = 0.5;
      R(s, { x, y, w, h: 0.72, fill: { color: l[3] } });
      T(s, l[0], { x: x + 0.12, y, w: 0.35, h: 0.72, fontSize: 22, bold: true, color: l[4], valign: "middle" });
      T(s, l[1], { x: x + 0.55, y: y + 0.08, w: w - 0.7, h: 0.26, fontSize: 12.5, bold: true, color: l[4] });
      T(s, l[2], { x: x + 0.55, y: y + 0.35, w: w - 0.7, h: 0.34, fontSize: 9.5, color: l[4] });
      if (l[5]) tag(s, l[5], x + w + 0.1, y + 0.24, 1.55, i === 0 ? C.muted : C.magenta);
    });
    R(s, { x: 6.85, y: 1.3, w: 2.65, h: 3.24, fill: { color: C.bg } });
    R(s, { x: 6.85, y: 1.3, w: 0.06, h: 3.24, fill: { color: C.magenta } });
    T(s, "Como explicar", { x: 7.03, y: 1.4, w: 2.4, h: 0.3, fontSize: 12, bold: true, color: C.navy });
    T(s, bullets([
      "A ferramenta é o meio. O produto é o acompanhamento técnico",
      "Medimos sucesso pelo resultado do cliente (ganho de peso, custo da arroba), não por acessos ao sistema",
      "Benchmark entre fazendas só existe com a base de clientes da TN",
      "A ferramenta pode mudar de fornecedor; o método de análise fica com a TN",
    ]), { x: 7.0, y: 1.78, w: 2.42, h: 2.7, fontSize: 9.5 });
    s.addNotes("Termo sugerido: 'item de prateleira' (ou 'commodity', palavra que o agro conhece bem). Mensagem central: qualquer concorrente entrega um sistema; poucos entregam um analista que olha os números toda semana e transforma em ajuste de dieta e manejo. É por isso que o projeto não é de TI: é de serviço técnico apoiado por dados.");
  }

  // 3 Como funciona (destaque 4 e 5)
  {
    const s = content("Como o produto funciona na rotina da fazenda", "A fazenda registra (passos 1 a 3). A TN transforma em decisão (passos 4 e 5), toda semana.");
    R(s, { x: 5.48, y: 1.18, w: 4.07, h: 2.45, fill: { color: C.magLight } });
    T(s, "ENTREGA TÉCNICA TN", { x: 5.6, y: 1.24, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: C.magenta, charSpacing: 2 });
    T(s, "A FAZENDA REGISTRA", { x: 0.5, y: 1.24, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: C.muted, charSpacing: 2 });
    const st = [
      ["1", "Cadastrar", "uma vez", "Fazenda, currais, lotes e dieta", C.muted],
      ["2", "Registrar", "todo dia", "Trato fornecido e leitura de cocho", C.muted],
      ["3", "Movimentar", "quando ocorre", "Entradas, pesagens, saídas e mortes", C.muted],
      ["4", "Analisar", "toda semana", "Painel e relatório prontos. O analista compara com o benchmark", C.navy],
      ["5", "Ajustar", "com o técnico TN", "Recomendação de dieta, trato e manejo; revisão no mês", C.magenta],
    ];
    st.forEach((p, i) => {
      const big = i >= 3, cx = 0.95 + i * 1.88 + (big ? 0.12 : 0), d = big ? 0.78 : 0.6, y = big ? 1.55 : 1.64;
      OV(s, { x: cx - d / 2 + 0.3, y, w: d, h: d, fill: { color: p[4] } });
      T(s, p[0], { x: cx - d / 2 + 0.3, y, w: d, h: d, fontSize: big ? 24 : 18, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[1], { x: cx - 0.55, y: 2.42, w: 1.7, h: 0.26, fontSize: big ? 13 : 11.5, bold: true, color: p[4], align: "center" });
      T(s, p[2], { x: cx - 0.55, y: 2.67, w: 1.7, h: 0.2, fontSize: 9, italic: true, color: C.muted, align: "center" });
      T(s, p[3], { x: cx - 0.55, y: 2.88, w: 1.7, h: 0.6, fontSize: 9, color: C.text, align: "center" });
      if (i < 4) s.addImage({ data: J.arrow, x: cx + 0.98 + (i === 2 ? 0.12 : 0), y: 1.85, w: 0.2, h: 0.2 });
    });
    // conexão: ciclo semanal
    LN(s, 2.9, 3.72, 9.1, 3.72, C.magenta, 1.25, "dash");
    LN(s, 2.9, 3.52, 2.9, 3.72, C.magenta, 1.25, "dash");
    LN(s, 9.1, 3.52, 9.1, 3.72, C.magenta, 1.25, "dash");
    T(s, "Ciclo semanal: o ajuste volta para o trato do dia seguinte", { x: 4.1, y: 3.62, w: 3.8, h: 0.2, fontSize: 9, italic: true, color: C.magenta, align: "center", fill: { color: C.white } });
    // quem faz 4 e 5
    T(s, "Quem faz os passos 4 e 5", { x: 0.5, y: 3.95, w: 3, h: 0.25, fontSize: 11, bold: true, color: C.navy });
    const ph = [["Meses 1 e 2", "Piloto: Rodrigo e técnico de campo em 3 a 5 fazendas", C.tealDark], ["Meses 3 a 6", "Rodrigo analisa até 20 fazendas, com auditoria mensal", C.navy], ["A partir do mês 7", "Analistas dedicados: 1 para cada 20 fazendas", C.magenta]];
    ph.forEach((p, i) => {
      const x = 0.5 + i * 2.75;
      R(s, { x, y: 4.24, w: 2.6, h: 0.56, fill: { color: C.bg } });
      R(s, { x, y: 4.24, w: 0.06, h: 0.56, fill: { color: p[2] } });
      T(s, p[0], { x: x + 0.15, y: 4.27, w: 2.4, h: 0.2, fontSize: 9.5, bold: true, color: p[2] });
      T(s, p[1], { x: x + 0.15, y: 4.47, w: 2.4, h: 0.28, fontSize: 8.5 });
    });
    s.addNotes("Os passos 1 a 3 são registro, e qualquer sistema faz. Os passos 4 e 5 são o serviço da TN e o motivo do projeto. Plano de capacidade: nos dois primeiros meses (piloto), Rodrigo e o técnico de campo acompanham de perto. Do mês 3 ao 6, Rodrigo analisa até 20 fazendas (estimativa de 1 hora por fazenda por semana, cerca de 20 horas semanais). A partir do mês 7, condicionado ao G2, entram analistas dedicados na proporção proposta de 1 para cada 20 fazendas. A proporção é proposta e será revista com dados do piloto.");
  }

  // 4 Arquitetura e dia a dia
  {
    const s = content("Na prática: o dia a dia e a arquitetura do sistema", "Como o dado sai do curral, chega à TN e volta como recomendação.");
    T(s, "Um dia na fazenda", { x: 0.5, y: 1.2, w: 3.6, h: 0.26, fontSize: 12, bold: true, color: C.navy });
    const dia = [
      ["Início da manhã", "Leitura de cocho antes do 1º trato: o tratador dá a nota de cada baia (escala adotada pela TN)"],
      ["Durante o dia", "Trato fornecido por baia, anotado na hora"],
      ["Quando ocorre", "Entradas, pesagens, saídas, mortes e troca de baia"],
      ["Fim do dia", "Escritório confere e envia. Meta: até 10 minutos"],
      ["Toda semana", "Relatório automático e análise do analista TN"],
    ];
    dia.forEach((d, i) => {
      const y = 1.52 + i * 0.6;
      OV(s, { x: 0.52, y: y + 0.05, w: 0.16, h: 0.16, fill: { color: i === 4 ? C.magenta : C.navy } });
      if (i < 4) LN(s, 0.6, y + 0.21, 0.6, y + 0.65, C.line, 1);
      T(s, d[0], { x: 0.8, y, w: 3.2, h: 0.2, fontSize: 10, bold: true, color: i === 4 ? C.magenta : C.navy });
      T(s, d[1], { x: 0.8, y: y + 0.2, w: 3.25, h: 0.38, fontSize: 9 });
    });
    // arquitetura
    T(s, "Arquitetura proposta", { x: 4.4, y: 1.2, w: 5, h: 0.26, fontSize: 12, bold: true, color: C.navy });
    const box = (x, y, w, h, ic, t1, t2, col) => {
      R(s, { x, y, w, h, fill: { color: col } });
      s.addImage({ data: ic, x: x + 0.1, y: y + 0.12, w: 0.3, h: 0.3 });
      T(s, t1, { x: x + 0.48, y: y + 0.06, w: w - 0.55, h: 0.22, fontSize: 9.5, bold: true, color: C.white });
      T(s, t2, { x: x + 0.48, y: y + 0.27, w: w - 0.55, h: h - 0.3, fontSize: 8, color: C.white });
    };
    box(4.4, 1.55, 1.85, 0.66, J.mobile, "Celular no curral", "Navegador; sem sinal, guarda e envia depois", C.tealDark);
    box(4.4, 2.32, 1.85, 0.66, J.paper, "Ficha de papel", "Digitada no computador da sede", C.tealDark);
    box(6.55, 1.55, 1.3, 1.43, J.cloud, "Nuvem", "Banco de dados e cálculos: consumo, GMD, dias de cocho", C.navy);
    box(8.15, 1.55, 1.4, 0.66, J.chart, "Painel", "Relatório semanal automático", C.navy);
    box(8.15, 2.32, 1.4, 0.66, J.user, "Analista TN", "Lê, compara e recomenda", C.magenta);
    s.addImage({ data: J.arrow, x: 6.3, y: 1.78, w: 0.2, h: 0.2 });
    s.addImage({ data: J.arrow, x: 6.3, y: 2.55, w: 0.2, h: 0.2 });
    s.addImage({ data: J.arrow, x: 7.9, y: 1.78, w: 0.2, h: 0.2 });
    LN(s, 8.85, 2.21, 8.85, 2.32, C.muted, 1.25);
    LN(s, 8.85, 2.98, 8.85, 3.16, C.magenta, 1.25);
    LN(s, 5.3, 3.16, 8.85, 3.16, C.magenta, 1.25);
    LN(s, 5.3, 2.98, 5.3, 3.16, C.magenta, 1.25);
    T(s, "Recomendação volta ao gestor: WhatsApp, ligação ou visita", { x: 5.4, y: 3.19, w: 3.4, h: 0.2, fontSize: 8.5, italic: true, color: C.magenta, align: "center" });
    // internet
    R(s, { x: 4.4, y: 3.52, w: 5.15, h: 1.2, fill: { color: C.bg } });
    s.addImage({ data: J.wifi, x: 4.52, y: 3.62, w: 0.3, h: 0.3 });
    T(s, "Precisa de internet?", { x: 4.92, y: 3.6, w: 4.5, h: 0.24, fontSize: 11, bold: true, color: C.navy });
    T(s, bullets([
      "No curral, não: ficha de papel ou celular que guarda o lançamento e envia ao ter sinal",
      "Na sede, sim: uma conexão por dia basta para enviar dados e ver o painel",
      "Celular sem sinal e app nativo offline: validar com a Dreamer na imersão técnica (16 a 27/11)",
    ]), { x: 4.55, y: 3.9, w: 4.9, h: 0.8, fontSize: 8.5 });
    s.addNotes("Arquitetura proposta para discussão com a software house Dreamer, não definida ainda. Princípio: o produto mínimo é um sistema web (abre no navegador do computador ou do celular), com dados na nuvem. Duas formas de lançar: (1) ficha de papel padrão preenchida no curral e digitada no fim do dia no computador da sede; (2) celular no curral pelo navegador. Guardar no aparelho e enviar depois (modo offline) depende de viabilidade técnica e custo, e o app nativo offline continua fora do produto mínimo, como no slide de MoSCoW. A escala de escore de cocho segue o padrão técnico da TN (por exemplo, notas de 0 a 4). A meta de 10 minutos por dia é critério do G1.");
  }

  // 5 Competitiva
  {
    const s = content("Análise competitiva: Cargill, DSM e Tecnobeef", "Os concorrentes vendem ferramenta junto da nutrição. Nosso espaço é a análise para o confinamento médio.", "Fontes: Cargill e Nutron (CattleView); estudo UFV com 6 confinamentos; DSM-firmenich (aquisição da Prodap); site e LinkedIn da Tecnobeef; avaliação interna TN (2023).");
    const cols = [["", 1.3], ["Oferta digital", 2.25], ["O que é bom", 2.7], ["O que é ruim", 2.75]];
    let x = 0.5; const xs = [];
    cols.forEach((c) => { xs.push(x); if (c[0]) head(s, c[0], x, 1.2, c[1] - 0.04, C.navy, 0.32, 10); x += c[1]; });
    const rows = [
      ["Cargill", "CattleView: câmeras, drones e IA para contar animais, ler cocho e acompanhar bem-estar; atendimento da Nutron", "Tecnologia de ponta e marketing forte da ferramenta; presença física nos grandes confinamentos", "Pensado para o grande confinamento com estrutura e sinal; leitura de sobra ainda em ajuste (estudo UFV); consultoria não é o ponto forte"],
      ["DSM (Prodap)", "Prodap Views: gestão do confinamento e da fazenda, com consultoria remota ou presencial (mais de 5 mil fazendas)", "Já une tecnologia, consultoria e nutrição sob medida; maior equipe de campo no Brasil", "Reestruturação após a aquisição; poucos consultores de alto nível; presença nas grandes contas abaixo da Cargill"],
      ["Tecnobeef", "Nutrição e suplementação sob medida; ferramenta digital própria não identificada", "Nasceu dentro de um grupo confinador (Campanelli, cerca de 110 mil cabeças): credibilidade prática", "Atuação mais regional; sem oferta digital pública. Confirmar no benchmark"],
      ["Trouw Nutrition", "Hoje: 20 licenças pagas de terceiros (TGC/Ponta e PecCode). Amanhã: TN Manager Web", "Equipe de campo reconhecida e produtos premium", "Dependemos de fornecedor; o pós-venda da Ponta decepcionou clientes; pouca presença digital"],
    ];
    rows.forEach((r, i) => {
      const y = 1.56 + i * 0.73, tn = i === 3;
      R(s, { x: 0.5, y, w: 9.0, h: 0.68, fill: { color: tn ? C.magLight : C.bg } });
      T(s, r[0], { x: 0.6, y, w: 1.15, h: 0.68, fontSize: 11, bold: true, color: tn ? C.magenta : C.navy, valign: "middle" });
      for (let k = 1; k < 4; k++) T(s, r[k], { x: xs[k] + 0.06, y: y + 0.05, w: cols[k][1] - 0.16, h: 0.6, fontSize: 8.5, valign: "middle", color: k === 3 && !tn ? C.text : C.text });
    });
    R(s, { x: 0.5, y: 4.5, w: 4.4, h: 0.34, fill: { color: C.navy } });
    T(s, [{ text: "Leitura: ", options: { bold: true, color: C.lime } }, { text: "ferramenta todos têm; análise semanal para 200 a 2.000 cabeças, ninguém faz de perto." }], { x: 0.6, y: 4.5, w: 4.25, h: 0.34, fontSize: 8.5, color: C.white, valign: "middle" });
    s.addNotes("Cargill: CattleView usa visão computacional, drones e nuvem para contagem de animais, leitura de cocho, estruturas, comportamento e bem-estar, água e pesagem; a Nutron usa drones em cerca de 15 confinamentos. Estudo da UFV em 6 confinamentos (1.754 imagens) mostrou que a quantificação de sobras ainda precisa de ajustes. DSM-firmenich comprou a Prodap (Belo Horizonte), que une tecnologia, consultoria e nutrição personalizada e atende mais de 5 mil fazendas. Tecnobeef: empresa de nutrição criada em 2017 dentro do grupo Campanelli, que confina cerca de 110 mil animais (dado do LinkedIn da empresa); não encontramos ferramenta digital pública, item a confirmar nas entrevistas de benchmark (2.1.1). Pontos fracos de Cargill e DSM e da própria TN vêm da avaliação interna do business case de 2023 e devem ser atualizados no benchmark.");
  }

  // 6 Custo-benefício
  {
    const s = content("Custo-benefício e cobrança", "Hoje pagamos licença de terceiros para 20 fazendas. O TN Manager Web troca custo de licença por serviço.");
    T(s, "O que pagamos hoje", { x: 0.5, y: 1.18, w: 4, h: 0.26, fontSize: 12, bold: true, color: C.navy });
    const cw = [1.25, 0.8, 1.35, 1.2]; const hx = [0.5]; cw.forEach((w, i) => hx.push(hx[i] + w));
    ["Fornecedor", "Licenças", "Por licença/ano", "Total/ano"].forEach((h, i) => head(s, h, hx[i], 1.48, cw[i] - 0.03, C.navy, 0.3, 9));
    [["TGC (Ponta)", "6", "R$ 10 a 40 mil", "R$ 60 a 240 mil"], ["PecCode", "14", "R$ 8 a 12,8 mil", "R$ 112 a 179 mil"], ["Total", "20", "", "R$ 172 a 419 mil"]].forEach((r, i) => {
      const y = 1.8 + i * 0.34;
      R(s, { x: 0.5, y, w: 4.57, h: 0.3, fill: { color: i === 2 ? C.magLight : C.bg } });
      r.forEach((c, k) => T(s, c, { x: hx[k] + 0.08, y, w: cw[k] - 0.12, h: 0.3, fontSize: 9.5, bold: i === 2 || k === 0, color: i === 2 ? C.magenta : C.text, valign: "middle" }));
    });
    const kp = [["R$ 364 mil", "orçados", C.navy], ["R$ 64 mil", "pagos até agora", C.tealDark], ["R$ 300 mil", "ainda a pagar", C.magenta], ["R$ 18,2 mil", "média por fazenda/ano", C.navy]];
    kp.forEach((k, i) => {
      const x = 0.5 + i * 1.155;
      R(s, { x, y: 2.92, w: 1.08, h: 0.78, fill: { color: C.bg } });
      R(s, { x, y: 2.92, w: 1.08, h: 0.05, fill: { color: k[2] } });
      T(s, k[0], { x: x + 0.06, y: 3.02, w: 1.0, h: 0.32, fontSize: 13, bold: true, color: k[2] });
      T(s, k[1], { x: x + 0.06, y: 3.34, w: 1.0, h: 0.32, fontSize: 8.5 });
    });
    T(s, "A licença média por fazenda supera o resultado que muitos confinamentos pequenos geram para a TN. Por isso a licença de terceiros não escala para esse público.", { x: 0.5, y: 3.82, w: 4.57, h: 0.55, fontSize: 9, italic: true, color: C.navy });
    // direita
    head(s, "O que muda com o TN Manager Web", 5.4, 1.18, 4.1, C.navy, 0.34, 10.5);
    R(s, { x: 5.4, y: 1.52, w: 4.1, h: 1.55, fill: { color: C.bg } });
    T(s, bullets([
      "Migrar as 14 fazendas PecCode: economia de R$ 112 a 179 mil por ano",
      "Migrar as 20 fazendas: até R$ 364 mil por ano (orçamento atual)",
      "Ponto de equilíbrio: o projeto se paga quando o custo anual do TN Manager Web (desenvolvimento, nuvem e suporte) ficar abaixo da licença que substitui",
      "Custo de desenvolvimento: orçamento da Dreamer após a imersão técnica (nov/2026)",
    ]), { x: 5.5, y: 1.6, w: 3.92, h: 1.45, fontSize: 9 });
    head(s, "Cobrança do cliente (proposta)", 5.4, 3.15, 4.1, C.magenta, 0.34, 10.5);
    R(s, { x: 5.4, y: 3.49, w: 4.1, h: 1.22, fill: { color: C.magLight } });
    T(s, bullets([
      "Sem licença para o cliente",
      "Custo embutido no pacote de nutrição e serviço técnico, como previsto na v1.0",
      "Quem não compra nutrição TN: mensalidade de serviço técnico",
      "Modelo final na etapa de monetização (jan/2027)",
    ]), { x: 5.5, y: 3.56, w: 3.92, h: 1.12, fontSize: 9 });
    s.addNotes("Dados informados: 20 licenças pagas a empresas parceiras. TGC (Ponta): 6 licenças de R$ 10 mil a R$ 40 mil. PecCode: 14 licenças de R$ 8 mil a R$ 12,8 mil. Orçado R$ 364 mil, pago até agora R$ 64 mil. Faixas calculadas: TGC 6 x 10 a 40 mil = 60 a 240 mil; PecCode 14 x 8 a 12,8 mil = 112 a 179,2 mil; total 172 a 419 mil. Média: 364 mil / 20 = 18,2 mil por fazenda. Premissa: valores anuais; confirmar a periodicidade dos contratos. As fazendas TGC tendem a ser maiores e podem continuar no TGC até existir a versão com integração. Referência da v1.0 (2023): custo estimado de R$ 2.500 por cliente no app interno contra R$ 13.541 no TGC.");
  }

  // 7 SWOT
  {
    const s = content("Análise SWOT: TN Manager Web", "Atualizada da v1.0 (2023) para a v2.0. Itens marcados em magenta são novos na v2.0.");
    const q = [
      ["F", "Forças", C.navy, 0.5, 1.2, [["Ganho de tempo: relatório pronto toda semana, sem planilha", 1], ["Benchmarking entre fazendas da carteira TN", 1], ["Dados próprios, sem depender de fornecedor", 1], ["Sem custo de licença para o cliente", 1], ["Marca forte, produtos premium e equipe técnica", 0]]],
      ["F", "Fraquezas", C.muted, 5.05, 1.2, [["Depende da disciplina de lançamento manual", 1], ["Time de análise pequeno no início (só Rodrigo)", 1], ["Ainda sem casos de sucesso próprios", 0], ["Modelo de cobrança ainda a definir", 0]]],
      ["O", "Oportunidades", C.tealDark, 0.5, 3.0, [["1.131 confinamentos com menos de 1.000 cabeças, pouco atendidos", 1], ["Cortar até R$ 364 mil por ano em licenças", 1], ["Recuperar ex-clientes e abrir não clientes", 0], ["Recomendação nutricional baseada em dados", 0], ["Clube de membros e programa de sustentabilidade", 0]]],
      ["A", "Ameaças", C.magenta, 5.05, 3.0, [["Cargill e DSM já vendem ferramenta junto da nutrição", 1], ["Baixa adesão ao lançamento manual", 1], ["Dependência da software house", 0], ["Não conseguir embutir o custo no preço do produto", 0]]],
    ];
    q.forEach((c) => {
      const [L, t, col, x, y, items] = c, w = 4.45, h = 1.7;
      R(s, { x, y, w, h, fill: { color: C.bg } });
      R(s, { x, y, w: 0.06, h, fill: { color: col } });
      OV(s, { x: x + 0.15, y: y + 0.1, w: 0.36, h: 0.36, fill: { color: col } });
      T(s, L, { x: x + 0.15, y: y + 0.1, w: 0.36, h: 0.36, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, t, { x: x + 0.6, y: y + 0.12, w: 3, h: 0.32, fontSize: 12.5, bold: true, color: col, valign: "middle" });
      T(s, items.map((it, j) => ({ text: it[0], options: { bullet: { indent: 9 }, breakLine: j < items.length - 1, paraSpaceAfter: 2, color: it[1] ? C.magenta : C.text, bold: !!it[1] } })), { x: x + 0.18, y: y + 0.52, w: w - 0.3, h: h - 0.56, fontSize: 9 });
    });
    s.addNotes("Mudanças em relação à SWOT da v1.0: entram como forças o ganho de tempo (relatório automático), o benchmarking entre fazendas, a posse dos dados e a ausência de licença. Sai a premissa de integração com hardware no início. Entram como fraquezas a dependência do lançamento manual e a capacidade de análise concentrada no Rodrigo nos primeiros meses. Oportunidade nova: o público de menos de 1.000 cabeças do Censo de Confinamento 2025 e a economia com licenças de terceiros.");
  }

  // 8 GTM
  {
    const s = content("Go-to-market: começar com 20 fazendas e escalar com analistas", "O Rodrigo valida o modelo de atendimento antes de contratar. O gatilho para crescer é resultado, não data.");
    const st = [
      ["Meses 1 e 2", "Piloto", "3 a 5 fazendas", "Rodrigo e técnico de campo acompanham de perto; ajustes rápidos no produto com a Dreamer", C.tealDark],
      ["Meses 3 a 6", "Rodrigo analisa", "até 20 fazendas", "Análise semanal remota, relatório comentado e auditoria mensal de dados e rotina", C.navy],
      ["A partir do mês 7", "Analistas dedicados", "1 para cada 20 fazendas", "Rodrigo coordena, padroniza a análise e cuida do benchmark", C.magenta],
    ];
    st.forEach((p, i) => {
      const x = 0.5 + i * 3.08, w = 2.85;
      R(s, { x, y: 1.22, w, h: 1.6, fill: { color: C.bg } });
      R(s, { x, y: 1.22, w, h: 0.06, fill: { color: p[4] } });
      T(s, p[0].toUpperCase(), { x: x + 0.14, y: 1.36, w: w - 0.2, h: 0.2, fontSize: 8.5, bold: true, color: p[4], charSpacing: 1.5 });
      T(s, p[1], { x: x + 0.14, y: 1.56, w: w - 0.2, h: 0.3, fontSize: 14, bold: true, color: C.navy });
      tag(s, p[2], x + 0.14, 1.9, 1.75, p[4]);
      T(s, p[3], { x: x + 0.14, y: 2.22, w: w - 0.26, h: 0.56, fontSize: 9 });
      if (i < 2) s.addImage({ data: J.arrow, x: x + w + 0.03, y: 1.92, w: 0.18, h: 0.18 });
    });
    const rows = [
      ["Primeiras fazendas", "Carteira com licença paga que cabe na faixa de 200 a 2.000 cabeças, depois indicação dos técnicos e RTVs"],
      ["O que o cliente recebe", "Ferramenta, relatório semanal comentado, visita do técnico e auditoria mensal da rotina"],
      ["Auditoria mensal", "Confere se o lançado bate com a realidade: estoque de ração, contagem de animais e leitura de cocho"],
      ["Gatilho para analistas", "20 fazendas ativas e G2 aprovado (adesão, uso e nota de recomendação)"],
      ["Preço", "Sem licença; custo no pacote de nutrição e serviço (modelo final em jan/2027)"],
    ];
    rows.forEach((r, i) => {
      const y = 2.98 + i * 0.35;
      R(s, { x: 0.5, y, w: 9.0, h: 0.31, fill: { color: i % 2 ? C.white : C.bg } });
      T(s, r[0], { x: 0.6, y, w: 1.9, h: 0.31, fontSize: 9.5, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.5, y, w: 6.9, h: 0.31, fontSize: 9, valign: "middle" });
    });
    s.addNotes("Capacidade estimada do Rodrigo: cerca de 1 hora por fazenda por semana, ou 20 horas semanais com 20 fazendas, mais a auditoria mensal. A proporção de 1 analista para cada 20 fazendas é proposta e será calibrada com o tempo real medido no piloto. Contar os meses a partir do início do piloto (item 8.0 do cronograma). Começar pelas fazendas que já têm licença paga pela TN gera economia imediata e usa clientes que já registram dados.");
  }

  // 9 Roadmap 12+ meses
  {
    const s = content("Roadmap: do papel para a fazenda em 15 meses", "Datas do cronograma aprovado até fev/2027; depois, datas indicativas que dependem de cada gate.");
    const months = ["out", "nov", "dez", "jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
    const x0 = 3.3, mw = 0.4, y0 = 1.22;
    T(s, "2026", { x: x0, y: y0 - 0.02, w: mw * 3, h: 0.18, fontSize: 8.5, bold: true, color: C.navy, align: "center" });
    T(s, "2027", { x: x0 + mw * 3, y: y0 - 0.02, w: mw * 12, h: 0.18, fontSize: 8.5, bold: true, color: C.navy, align: "center" });
    months.forEach((m, i) => {
      T(s, m, { x: x0 + i * mw, y: y0 + 0.17, w: mw, h: 0.18, fontSize: 8, color: C.muted, align: "center" });
      if (i) LN(s, x0 + i * mw, y0 + 0.36, x0 + i * mw, 4.4, "ECEFF2", 0.5);
    });
    LN(s, x0 + 3 * mw, y0, x0 + 3 * mw, 4.4, C.line, 0.75);
    const rows = [
      ["Business case e benchmark", "Rodrigo", 0, 1, C.tealDark, "até 23/10"],
      ["Pesquisa com 30 fazendas", "Rodrigo e Kátia Sato", 0.6, 2.2, C.tealDark, "19/10 a 04/12"],
      ["Imersão técnica Dreamer", "Rodrigo e Dreamer", 1.5, 0.5, C.tealDark, "16 a 27/11"],
      ["Diagnóstico e PRDs", "Rodrigo", 2.2, 0.6, C.navy, "07 a 18/12"],
      ["Roadmap, GTM e monetização", "Rodrigo e diretoria", 3.1, 1.0, C.navy, "jan"],
      ["Desenvolver o produto mínimo", "Dreamer", 4.2, 3.0, C.navy, "fev a abr"],
      ["Piloto em 3 a 5 fazendas", "Rodrigo e técnicos", 7, 2, C.magenta, "mai a jun"],
      ["Rodrigo atende até 20 fazendas", "Rodrigo", 9, 4, C.magenta, "jul a out"],
      ["Analistas dedicados", "Diretoria aprova", 13, 2, C.magenta, "nov em diante"],
    ];
    rows.forEach((r, i) => {
      const y = y0 + 0.42 + i * 0.3;
      T(s, r[0], { x: 0.5, y, w: 2.0, h: 0.26, fontSize: 9, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.5, y, w: 0.85, h: 0.26, fontSize: 7.5, color: C.muted, valign: "middle" });
      RR(s, { x: x0 + r[2] * mw, y: y + 0.03, w: r[3] * mw, h: 0.2, fill: { color: r[4] }, rectRadius: 0.05 });
      const lw = r[3] * mw;
      T(s, r[5], { x: lw > 0.75 ? x0 + r[2] * mw : x0 + (r[2] + r[3]) * mw + 0.04, y: y + 0.03, w: lw > 0.75 ? lw : 1.0, h: 0.2, fontSize: 7.5, bold: true, color: lw > 0.75 ? C.white : r[4], align: lw > 0.75 ? "center" : "left", valign: "middle" });
    });
    const gy = 4.4;
    T(s, "Decisões", { x: 0.5, y: gy - 0.04, w: 2, h: 0.26, fontSize: 9, bold: true, color: C.magenta, valign: "middle" });
    [["G0", 2.8, "escopo"], ["Aprovação", 4.1, "investimento"], ["G1", 9, "piloto"], ["G2", 13, "escala"]].forEach((g) => {
      const gx = x0 + g[1] * mw;
      s.addShape(pres.shapes.DIAMOND, { x: gx - 0.1, y: gy, w: 0.2, h: 0.2, fill: { color: C.lime }, line: { type: "none" } });
      const al = g[0] === "G0" ? "right" : g[0] === "Aprovação" ? "left" : "center";
      const lx = al === "right" ? gx - 1.1 : al === "left" ? gx - 0.1 : gx - 0.6;
      T(s, g[0] + ": " + g[2], { x: lx, y: gy + 0.21, w: 1.2, h: 0.18, fontSize: 7.5, bold: true, color: C.navy, align: al });
    });
    s.addNotes("Até fevereiro de 2027 as datas são as do cronograma revisado (itens 2.1 a 7.0). G0 em dezembro, com o diagnóstico; a apresentação final de 01 a 05/02/2027 é o momento de aprovar o investimento. Desenvolvimento de cerca de 12 semanas é estimativa a confirmar com a Dreamer após a imersão técnica. O piloto dura 8 semanas, prazo do critério G1. G2 depois de 90 dias com as fazendas ativas. Se um gate não for atingido, as barras seguintes andam para a direita: as datas a partir de março são indicativas.");
  }

  // 10 Cronograma 4 anos
  {
    const s = content("Cronograma de 4 anos: 2027 a 2030", "Do piloto ao programa de dados da TN. Metas propostas para discussão, revistas a cada gate.");
    const yrs = [["2027", "Piloto e 20 fazendas"], ["2028", "Escala com analistas"], ["2029", "Benchmark e fidelidade"], ["2030", "Programa de dados"]];
    const x0 = 2.05, cw = 1.86;
    yrs.forEach((y, i) => {
      const col = [C.tealDark, C.navy, C.magenta, C.lime][i];
      R(s, { x: x0 + i * cw, y: 1.2, w: cw - 0.06, h: 0.5, fill: { color: col } });
      T(s, y[0], { x: x0 + i * cw, y: 1.2, w: cw - 0.06, h: 0.27, fontSize: 13, bold: true, color: i === 3 ? C.navy : C.white, align: "center", valign: "middle" });
      T(s, y[1], { x: x0 + i * cw, y: 1.45, w: cw - 0.06, h: 0.22, fontSize: 8.5, color: i === 3 ? C.navy : C.white, align: "center" });
    });
    const lanes = [
      ["Produto", ["Produto mínimo e piloto; versão estável", "Integração com balança, vagão e NOA; uso sem sinal", "Painel de benchmark entre fazendas", "Pegada de carbono e recomendação de produto"]],
      ["Atendimento técnico", ["Rodrigo analisa; técnicos de campo apoiam", "Analistas dedicados (1 para 20 fazendas)", "Núcleo de análise com método padrão", "Análise integrada à oferta comercial"]],
      ["Fazendas atendidas", ["20", "60", "120", "200"]],
      ["Analistas dedicados", ["0 (Rodrigo)", "3", "6", "10"]],
      ["Licenças de terceiros", ["Migrar fazendas na faixa do produto", "Encerrar PecCode", "TGC só onde houver integração", "Sem licença de terceiros"]],
    ];
    lanes.forEach((l, i) => {
      const num = i === 2 || i === 3, h = num ? 0.42 : 0.56;
      const y = 1.8 + [0, 0.62, 1.24, 1.72, 2.2][i];
      R(s, { x: 0.5, y, w: 1.5, h, fill: { color: C.navy } });
      T(s, l[0], { x: 0.58, y, w: 1.38, h, fontSize: 9.5, bold: true, color: C.white, valign: "middle" });
      l[1].forEach((c, k) => {
        R(s, { x: x0 + k * cw, y, w: cw - 0.06, h, fill: { color: C.bg } });
        T(s, c, { x: x0 + k * cw + 0.06, y, w: cw - 0.18, h, fontSize: num ? 13 : 8.5, bold: num, color: num ? C.navy : C.text, align: "center", valign: "middle" });
      });
    });
    T(s, "Antes, em 2026: business case, pesquisa com 30 fazendas, imersão técnica e PRDs (out a dez).", { x: 0.5, y: 4.62, w: 4.4, h: 0.3, fontSize: 8.5, italic: true, color: C.navy });
    s.addNotes("Metas de fazendas e analistas são propostas: 20 em 2027 (capacidade do Rodrigo), 60 em 2028, 120 em 2029 e 200 em 2030, alinhado à meta de 200 clientes da v1.0. Analistas na proporção de 1 para cada 20 fazendas. As funcionalidades por ano vêm da visão estratégica original: 2027 versão estável e integração com hardware; 2028 versão madura com NOA e complementos da AgTech Garage; 2029 fidelidade, clube de membros e benchmark; 2030 sustentabilidade, pegada de carbono e recomendação de produto. Licenças: priorizar a migração das 14 fazendas PecCode; manter TGC onde a integração com balança e vagão for indispensável até a versão que a ofereça.");
  }

  // 11 Decisões
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.navy };
    R(s, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.lime } });
    T(s, "Decisões e próximos passos", { x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true, color: C.white });
    T(s, "O que precisamos aprovar hoje para seguir com a pesquisa e a imersão técnica.", { x: 0.5, y: 0.8, w: 8.6, h: 0.3, fontSize: 13, italic: true, color: "D6DCEB" });
    T(s, "DECISÕES PARA HOJE", { x: 0.5, y: 1.3, w: 4, h: 0.22, fontSize: 9, bold: true, color: C.lime, charSpacing: 2 });
    const dec = [
      "Aprovar o foco em confinamentos de 200 a 2.000 cabeças",
      "Aprovar o roadmap de 15 meses com os gates G0 a G2",
      "Aprovar o modelo: Rodrigo atende 20 fazendas; analistas após o G2",
      "Autorizar a imersão técnica e o orçamento com a Dreamer",
      "Indicar um patrocinador na diretoria",
    ];
    dec.forEach((d, i) => {
      const y = 1.62 + i * 0.52;
      OV(s, { x: 0.5, y, w: 0.34, h: 0.34, fill: { color: C.magenta } });
      s.addImage({ data: I.check, x: 0.58, y: y + 0.08, w: 0.18, h: 0.18 });
      T(s, d, { x: 0.98, y, w: 3.9, h: 0.34, fontSize: 11, color: C.white, valign: "middle" });
    });
    R(s, { x: 5.3, y: 1.25, w: 4.25, h: 3.5, fill: { color: C.navyDark } });
    T(s, "PRÓXIMAS SEMANAS", { x: 5.5, y: 1.35, w: 3.8, h: 0.22, fontSize: 9, bold: true, color: C.lime, charSpacing: 2 });
    [["até 23/10", "Entrevistas internas e benchmark (inclui Tecnobeef)"], ["19 a 23/10", "Rascunho do formulário"], ["26/10 a 06/11", "Revisão com Kátia Sato e teste"], ["09 a 27/11", "Respostas das fazendas (meta: 30)"], ["16 a 27/11", "Imersão técnica com a Dreamer"], ["07 a 18/12", "Diagnóstico e PRDs (G0)"], ["01 a 05/02", "Apresentação final e investimento"]].forEach((r, i) => {
      const y = 1.65 + i * 0.42;
      T(s, r[0], { x: 5.5, y, w: 1.2, h: 0.36, fontSize: 9.5, bold: true, color: C.lime, valign: "middle" });
      T(s, r[1], { x: 6.7, y, w: 2.75, h: 0.36, fontSize: 9.5, color: C.white, valign: "middle" });
    });
  }

  // 12-13 Mercado

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
