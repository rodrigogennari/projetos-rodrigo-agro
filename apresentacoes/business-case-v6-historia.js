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

  // 1 Agenda em capítulos
  {
    const s = content("Agenda", "Uma história em seis capítulos");
    const items = [
      ["De onde viemos", "A v1.0 começou pela automação. A versão Web começa pela rotina"],
      ["Para quem", "Mercado, cliente-alvo e a nova geração que assume as fazendas"],
      ["O que entregamos", "O diferencial é a entrega técnica: como funciona na prática"],
      ["Contra quem", "Cargill, DSM, Ponta, Tecnobeef e outros"],
      ["Como crescemos", "Custos, go-to-market e o Grupo de Desenvolvimento"],
      ["Até onde vamos", "Roadmap, visão de 4 anos e decisões de hoje"],
    ];
    items.forEach((it, i) => {
      const x = 0.6, y = 1.3 + i * 0.56;
      OV(s, { x, y, w: 0.42, h: 0.42, fill: { color: i === 2 ? C.navy : C.magenta } });
      T(s, String(i + 1), { x, y, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, it[0], { x: x + 0.6, y: y - 0.02, w: 6.5, h: 0.26, fontSize: 14, bold: true, color: C.navy });
      T(s, it[1], { x: x + 0.6, y: y + 0.24, w: 6.5, h: 0.2, fontSize: 10, color: C.text });
      if (i < 5) LN(s, x + 0.6, y + 0.5, 8.2, y + 0.5, "ECEFF2", 0.5);
    });
    s.addNotes("Abertura sugerida: 'Vou contar como saímos de um projeto que começou pelo fim para um produto que começa pela rotina do curral e termina no resultado do pecuarista.' O capítulo 3 é o coração da conversa. O anexo, depois do agradecimento, tem o material de apoio para perguntas.");
  }

  // 2 Nossa história
  {
    const s = content("Nossa história: da automação para o resultado", "A v1.0 começou pelo fim. A versão Web começa pela rotina do curral e termina no resultado do pecuarista.", "Fontes: Probeef (95% dos confinamentos analisados usam software de gestão); Embrapa (2018): 106 softwares para pecuária de corte e queixa de sistemas complexos e caros.");
    const chain = (y, label, sub, steps, cols, txtCols) => {
      T(s, label, { x: 0.5, y: y + 0.02, w: 1.75, h: 0.28, fontSize: 12, bold: true, color: cols[cols.length - 1] === C.muted ? C.muted : C.navy });
      T(s, sub, { x: 0.5, y: y + 0.3, w: 1.75, h: 0.4, fontSize: 9, italic: true, color: C.muted });
      const x0 = 2.3, w = 7.2 / steps.length;
      steps.forEach((st, i) => {
        s.addShape(i === 0 ? pres.shapes.PENTAGON : pres.shapes.CHEVRON, { x: x0 + i * w, y, w: w + 0.12, h: 0.7, fill: { color: cols[i] }, line: { color: C.white, width: 2 } });
        T(s, st, { x: x0 + i * w + (i === 0 ? 0.1 : 0.34), y, w: w - 0.42, h: 0.7, fontSize: 9, bold: true, color: txtCols ? txtCols[i] : C.white, align: "center", valign: "middle" });
      });
    };
    chain(1.25, "TN Manager v1.0", "2023: começou pelo fim", ["Automação e hardware", "Integrações", "Grandes confinamentos", "Implantação longa e cara"], ["9AA3AC", "9AA3AC", "9AA3AC", "9AA3AC"]);
    chain(2.15, "TN Manager Web", "2026: começa pelo começo", ["Rotina do curral", "Dados confiáveis", "Análise técnica TN", "Resultado do pecuarista", "Automação, se o cliente pedir"], [C.tealDark, C.navy, C.navy, C.magenta, "E3E7EC"], [C.white, C.white, C.white, C.white, C.muted]);
    const pr = [["Simples", "Poucos campos e 10 minutos por dia, no celular ou no papel"], ["Objetivo", "Só os indicadores que mudam a decisão: consumo, ganho, custo e dias de cocho"], ["Flexível", "Funciona sem internet no curral e cresce por versões"]];
    pr.forEach((p, i) => {
      const x = 0.5 + i * 1.78;
      R(s, { x, y: 3.12, w: 1.68, h: 1.5, fill: { color: C.bg } });
      R(s, { x, y: 3.12, w: 1.68, h: 0.06, fill: { color: [C.tealDark, C.navy, C.magenta][i] } });
      T(s, p[0], { x: x + 0.12, y: 3.26, w: 1.48, h: 0.3, fontSize: 14, bold: true, color: [C.tealDark, C.navy, C.magenta][i] });
      T(s, p[1], { x: x + 0.12, y: 3.6, w: 1.46, h: 0.95, fontSize: 9.5 });
    });
    R(s, { x: 5.95, y: 3.12, w: 3.55, h: 1.5, fill: { color: C.navy } });
    T(s, "MENOS É MAIS", { x: 6.15, y: 3.24, w: 3.2, h: 0.24, fontSize: 9, bold: true, color: C.lime, charSpacing: 3 });
    T(s, "Sistema virou commodity. O que manda é o resultado que ajudamos o pecuarista a alcançar.", { x: 6.15, y: 3.5, w: 3.2, h: 1.05, fontSize: 14, bold: true, color: C.white, valign: "top" });
    s.addNotes("Fala sugerida: 'A primeira versão do TN Manager foi desenhada do fim para o começo: automação, integração com balança e vagão, foco nos grandes. É um projeto pesado, caro e demorado para chegar ao campo. Entrei para inverter essa lógica: começar pela rotina do curral, com um sistema simples, objetivo e flexível para o confinamento pequeno e médio, e deixar a automação para quando o cliente pedir. Sistema hoje é commodity: a Probeef mostra que 95% dos confinamentos que ela analisa já usam software de gestão, e a Embrapa mapeou mais de 100 softwares para pecuária de corte, com reclamação de que são complexos e caros. O que diferencia é o que fazemos com o dado.' Ponte: 'E para quem estamos fazendo isso?'");
  }

  // 3 Nova geração
  {
    const s = content("A nova geração está assumindo o confinamento", "Sucessão familiar: quem chega quer praticidade, dados e decisão com cara de empresa.", "Fontes: pesquisa Unicesumar com 119 filhos de produtores de 17 a 35 anos (Canal Rural, 2026); Embrapa Pecuária Sudeste (estudos com 184 e 130 confinamentos).");
    R(s, { x: 0.5, y: 1.25, w: 2.45, h: 1.75, fill: { color: C.navy } });
    T(s, "90%", { x: 0.65, y: 1.32, w: 2.2, h: 0.62, fontSize: 36, bold: true, color: C.white });
    T(s, "dos filhos de produtores ouvidos pretendem assumir a gestão da propriedade da família", { x: 0.65, y: 1.95, w: 2.2, h: 0.95, fontSize: 10, color: C.white });
    R(s, { x: 0.5, y: 3.1, w: 2.45, h: 1.55, fill: { color: C.bg } });
    R(s, { x: 0.5, y: 3.1, w: 0.06, h: 1.55, fill: { color: C.tealDark } });
    T(s, "Embrapa", { x: 0.68, y: 3.2, w: 2.2, h: 0.26, fontSize: 11, bold: true, color: C.tealDark });
    T(s, "Escolaridade e consultoria aumentam a adoção de tecnologia digital nos confinamentos. Quem adota software tem mais eficiência técnica.", { x: 0.68, y: 3.48, w: 2.2, h: 1.1, fontSize: 9.5 });
    T(s, "O que muda com a sucessão", { x: 3.2, y: 1.2, w: 3.2, h: 0.26, fontSize: 12, bold: true, color: C.navy });
    const dp = [["Caderno e memória", "Celular e painel"], ["\"Sempre fiz assim\"", "Indicador e meta"], ["Decisão no olho", "Decisão com número"], ["Fazenda da família", "Empresa rural com fechamento de safra"]];
    dp.forEach((d, i) => {
      const y = 1.55 + i * 0.78;
      R(s, { x: 3.2, y, w: 1.35, h: 0.66, fill: { color: "E3E7EC" } });
      T(s, d[0], { x: 3.25, y, w: 1.25, h: 0.66, fontSize: 9.5, color: C.muted, align: "center", valign: "middle" });
      s.addImage({ data: J.arrow, x: 4.62, y: y + 0.23, w: 0.2, h: 0.2 });
      R(s, { x: 4.9, y, w: 1.6, h: 0.66, fill: { color: C.navy } });
      T(s, d[1], { x: 4.95, y, w: 1.5, h: 0.66, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle" });
    });
    R(s, { x: 6.75, y: 1.2, w: 2.75, h: 3.45, fill: { color: C.magLight } });
    R(s, { x: 6.75, y: 1.2, w: 0.06, h: 3.45, fill: { color: C.magenta } });
    T(s, "Como o TN Manager Web responde", { x: 6.92, y: 1.3, w: 2.5, h: 0.45, fontSize: 11.5, bold: true, color: C.magenta });
    const rsp = [["Praticidade", "10 minutos por dia, no celular"], ["Dados", "Painel semanal e benchmark com fazendas parecidas"], ["Decisão", "Analista TN junto, toda semana"], ["Gestão empresarial", "Fechamento de safra e evolução ano a ano"]];
    rsp.forEach((r, i) => {
      const y = 1.85 + i * 0.68;
      T(s, r[0], { x: 6.92, y, w: 2.5, h: 0.24, fontSize: 10.5, bold: true, color: C.navy });
      T(s, r[1], { x: 6.92, y: y + 0.24, w: 2.5, h: 0.4, fontSize: 9.5 });
    });
    s.addNotes("Fala sugerida: 'Quem está chegando na fazenda é outra geração. Nove em cada dez filhos de produtores ouvidos em uma pesquisa da Unicesumar querem assumir a gestão. Eles não querem planilha nem sistema complicado: querem praticidade, número e decisão. A Embrapa mostra que escolaridade e consultoria puxam a adoção digital no confinamento. O TN Manager Web foi desenhado para essa pessoa.' Dado: pesquisa com 119 jovens de 17 a 35 anos, 90,4% pretendem assumir a gestão (Canal Rural, 2026). A pesquisa não é só de pecuária de corte: usar como sinal, não como número do nosso público. Ponte: 'Então, o mercado é grande e o cliente está mudando. Estamos no caminho certo?'");
  }

  // 4 Grupo de Desenvolvimento
  {
    const s = content("Grupo de Desenvolvimento de Confinadores", "Comunidade, dados e reconhecimento: o modelo do Fazenda Nota 10 (2020), agora apoiado pelo TN Manager Web.");
    const st = [
      ["Lançar", "A fazenda registra a rotina no TN Manager Web", C.tealDark],
      ["Analista online", "Cobra, audita e valida os lançamentos", C.tealDark],
      ["Indicadores", "Dashboards de consumo, ganho, custo e dias de cocho", C.navy],
      ["Benchmarking", "Cada fazenda se compara com o grupo", C.navy],
      ["Fechamento de safra", "Relatório da safra e evolução ano a ano", C.magenta],
      ["Evento anual", "Análise dos dados e premiação das melhores fazendas", C.magenta],
    ];
    st.forEach((p, i) => {
      const x = 0.5 + i * 1.52, w = 1.42;
      R(s, { x, y: 1.25, w, h: 1.32, fill: { color: C.bg } });
      R(s, { x, y: 1.25, w, h: 0.06, fill: { color: p[2] } });
      OV(s, { x: x + 0.1, y: 1.4, w: 0.34, h: 0.34, fill: { color: p[2] } });
      T(s, String(i + 1), { x: x + 0.1, y: 1.4, w: 0.34, h: 0.34, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });
      T(s, p[0], { x: x + 0.1, y: 1.8, w: w - 0.16, h: 0.24, fontSize: 10.5, bold: true, color: p[2] });
      T(s, p[1], { x: x + 0.1, y: 2.05, w: w - 0.16, h: 0.5, fontSize: 8.5 });
      if (i < 5) s.addImage({ data: J.arrow, x: x + w - 0.04, y: 1.48, w: 0.14, h: 0.14 });
    });
    LN(s, 9.21, 2.57, 9.21, 2.8, C.magenta, 1.25, "dash");
    LN(s, 1.21, 2.8, 9.21, 2.8, C.magenta, 1.25, "dash");
    LN(s, 1.21, 2.57, 1.21, 2.8, C.magenta, 1.25, "dash");
    T(s, "O evento atrai novas fazendas, e o ciclo recomeça na safra seguinte", { x: 3.0, y: 2.69, w: 4.0, h: 0.22, fontSize: 9, italic: true, color: C.magenta, align: "center", fill: { color: C.white } });
    const box = [
      ["Analistas online", C.tealDark, ["Cobram lançamentos em atraso", "Auditam e validam os dados", "Interpretam os números com o gestor", "Tiram dúvidas e mantêm o engajamento"]],
      ["O que a fazenda ganha", C.navy, ["Referência para decidir: onde está em relação ao grupo", "Evolução medida safra a safra", "Reconhecimento no evento anual", "Rede com outros confinadores"]],
      ["O que a TN ganha", C.magenta, ["Fidelização: a fazenda fica pelo grupo, não pelo sistema", "Base de dados própria para benchmark e nutrição", "Novas fazendas querendo entrar"]],
    ];
    box.forEach((b, i) => {
      const x = 0.5 + i * 3.05, w = 2.88;
      head(s, b[0], x, 3.05, w, b[1], 0.32, 10.5);
      R(s, { x, y: 3.37, w, h: 1.3, fill: { color: C.bg } });
      T(s, bullets(b[2]), { x: x + 0.08, y: 3.44, w: w - 0.16, h: 1.2, fontSize: 9 });
    });
    s.addNotes("Fala sugerida: 'Em 2020 criei o Fazenda Nota 10, um grupo de fazendas que se desenvolviam juntas olhando os próprios números. A chave do longo prazo é repetir isso com o TN Manager Web: uma comunidade de confinadores com dados, indicadores e benchmarking para decidir, fechamento de safra para cada fazenda e um evento anual com análise dos dados e premiação. Analistas online cobram e auditam lançamentos, validam, interpretam e tiram dúvidas. É isso que mantém o engajamento: a fazenda fica pelo grupo, não pelo sistema.' Sugestão: se houver números do Fazenda Nota 10 (quantas fazendas, quanto evoluíram), coloque aqui; é a melhor prova de que o modelo funciona. Ponte: 'Como isso entra no calendário?'");
  }

  // 5 Concorrência
  {
    const s = content("Concorrência: software todos têm", "A diferença está em quem transforma o dado em resultado para o confinador médio.", "Fontes: Cargill; Integra Software (Feed Manager 1.3, 2026); AgFeed e dsm-firmenich (FarmTell, 2025); Ponta Agro; AgFeed (Tecnobeef); Baguete (JetBov); Embrapa (2018); UFV; avaliação interna TN.");
    const cols = [["", 1.2], ["Ferramenta", 2.25], ["Pontos positivos", 2.75], ["Pontos negativos", 2.8]];
    let x = 0.5; const xs = [];
    cols.forEach((c) => { xs.push(x); if (c[0]) head(s, c[0], x, 1.18, c[1] - 0.04, C.navy, 0.3, 10); x += c[1]; });
    const rows = [
      ["Cargill", "Feed Manager (manejo, nutrição, trato, fábrica de ração, SISBOV) com Feed Tracer e CattleView (câmeras, drones e IA)", "Solução completa e automatizada; previsão de abate; forte nos grandes confinamentos", "Pensado para quem já tem automação e estrutura; software mantido por terceira (Integra); leitura de sobra por imagem ainda em ajuste (UFV)"],
      ["DSM-firmenich", "FarmTell, ex-Prodap: Beef, Views e Beef Smart; quase 10 softwares e uma plataforma de IA", "Une tecnologia e consultoria; cresce de 20% a 30% ao ano; tem versão simples para fazendas menores (Beef Smart)", "Marca em transição de Prodap para FarmTell; portfólio grande e fragmentado; Beef e Views voltados aos grandes"],
      ["Ponta (TGC)", "TGC, gestão de confinamento; parceiro da Assocon desde 2009", "Referência nos grandes: diz atender 12 dos 20 maiores confinamentos", "Licença cara para o pequeno (R$ 10 mil a 40 mil por ano nas nossas licenças); pós-venda decepcionou clientes TN"],
      ["Tecnobeef", "Nutrição sob medida; ferramenta digital para o cliente não identificada", "Nasceu de confinador (grupo Campanelli, mais de 110 mil cabeças); atende cerca de 170 propriedades", "Sem oferta digital para o cliente; escala menor e regional"],
      ["JetBov e outros", "Software em nuvem para pecuária (Berrante, iRancho e outros)", "Barato e simples; foco em fazendas de até 3 mil animais", "Só software: sem nutrição, sem técnico no campo e sem análise"],
      ["Trouw Nutrition", "Hoje: licenças TGC e PecCode. Amanhã: TN Manager Web com analista online", "Equipe de campo, produtos premium, comunidade e benchmark", "Ainda dependemos de fornecedor; ferramenta própria em construção"],
    ];
    rows.forEach((r, i) => {
      const y = 1.52 + i * 0.54, tn = i === 5;
      R(s, { x: 0.5, y, w: 9.0, h: 0.5, fill: { color: tn ? C.magLight : C.bg } });
      T(s, r[0], { x: 0.58, y, w: 1.1, h: 0.5, fontSize: 10, bold: true, color: tn ? C.magenta : C.navy, valign: "middle" });
      for (let k = 1; k < 4; k++) T(s, r[k], { x: xs[k] + 0.05, y, w: cols[k][1] - 0.13, h: 0.5, fontSize: 7.8, valign: "middle" });
    });
    s.addNotes("Fontes e cuidados. Cargill: Feed Manager cobre manejo, nutrição, controle de trato, fábrica de ração, rastreabilidade SISBOV e previsão de abate; o Feed Tracer coleta o trato automaticamente; o histórico da versão 1.3 é publicado pela Integra Software (última versão listada 1.3.101, maio de 2026). CattleView: visão computacional, drones e IA; estudo da UFV em 6 confinamentos mostrou que a quantificação de sobras ainda precisa de ajuste. DSM-firmenich: comprou a Prodap em 2022 e reuniu as soluções na marca FarmTell (Beef para grandes, Views para rotina e indicadores, Beef Smart mais simples para fazendas menores); cerca de dez softwares e uma plataforma de IA; crescimento de 20% a 30% ao ano (AgFeed). Atenção: o Beef Smart é o concorrente mais parecido com o TN Manager Web. Ponta: a afirmação dos 12 dos 20 maiores é da própria Ponta. Tecnobeef: nutrição criada em 2017 para o confinamento do grupo Campanelli; cerca de 170 propriedades atendidas (AgFeed); não encontramos app próprio para clientes. JetBov: foco em até 3 mil animais, modelo em nuvem com plano gratuito limitado (preços de reportagem antiga). Probeef: 95% dos confinamentos analisados usam software de gestão. Embrapa (2018): 106 softwares para pecuária de corte e queixa de sistemas complexos e caros. PecCode fica fora da tabela porque é parceira para automação. Ponte: 'Se software todos têm, quanto custa o nosso e como cobramos?'");
  }

  // 7 SWOT
  {
    const s = content("Análise SWOT: TN Manager Web", "Atualizada da v1.0 (2023) para a v2.0. Itens marcados em magenta são novos na v2.0.");
    const q = [
      ["F", "Forças", C.navy, 0.5, 1.2, [["Ganho de tempo: relatório pronto toda semana, sem planilha", 1], ["Benchmarking e comunidade (Grupo de Desenvolvimento)", 1], ["Dados próprios, sem depender de fornecedor", 1], ["Sem custo de licença para o cliente", 1], ["Marca forte, produtos premium e equipe técnica", 0]]],
      ["F", "Fraquezas", C.muted, 5.05, 1.2, [["Depende da disciplina de lançamento manual", 1], ["Time de análise pequeno no início (só Rodrigo)", 1], ["Ainda sem casos de sucesso próprios", 0], ["Modelo de cobrança ainda a definir", 0]]],
      ["O", "Oportunidades", C.tealDark, 0.5, 3.0, [["11.194 confinamentos de até 1.500 cabeças (Bússola Farm), pouco atendidos", 1], ["Cortar até R$ 364 mil por ano em licenças", 1], ["Recuperar ex-clientes e abrir não clientes", 0], ["Recomendação nutricional baseada em dados", 0], ["Nova geração na sucessão quer dados e gestão empresarial", 1]]],
      ["A", "Ameaças", C.magenta, 5.05, 3.0, [["Cargill e DSM já vendem ferramenta junto da nutrição", 1], ["DSM já tem versão simples para pequenos (FarmTell Beef Smart)", 1], ["Baixa adesão ao lançamento manual", 1], ["Dependência da software house", 0], ["Não conseguir embutir o custo no preço do produto", 0]]],
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
      ["A partir do mês 7", "Analistas online", "1 para cada 20 fazendas", "Cobram, auditam e validam lançamentos, interpretam e tiram dúvidas. Rodrigo coordena", C.magenta],
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
      ["Analistas online", "Diretoria aprova", 13, 2, C.magenta, "nov em diante"],
      ["Grupo piloto e fechamento de safra", "Rodrigo e analistas", 12, 3, C.lime, "out a dez"],
    ];
    rows.forEach((r, i) => {
      const y = y0 + 0.4 + i * 0.27;
      T(s, r[0], { x: 0.5, y, w: 2.0, h: 0.26, fontSize: 9, bold: true, color: C.navy, valign: "middle" });
      T(s, r[1], { x: 2.5, y, w: 0.85, h: 0.26, fontSize: 7.5, color: C.muted, valign: "middle" });
      RR(s, { x: x0 + r[2] * mw, y: y + 0.03, w: r[3] * mw, h: 0.2, fill: { color: r[4] }, rectRadius: 0.05 });
      const lw = r[3] * mw;
      T(s, r[5], { x: lw > 0.75 ? x0 + r[2] * mw : x0 + (r[2] + r[3]) * mw + 0.04, y: y + 0.03, w: lw > 0.75 ? lw : 1.0, h: 0.2, fontSize: 7.5, bold: true, color: lw > 0.75 ? (r[4] === C.lime ? C.navy : C.white) : r[4], align: lw > 0.75 ? "center" : "left", valign: "middle" });
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
    const s = content("Cronograma de 4 anos: 2027 a 2030", "Do piloto à comunidade de confinadores TN. Metas propostas, revistas a cada gate.");
    const yrs = [["2027", "Piloto e 20 fazendas"], ["2028", "Grupo e 1º evento"], ["2029", "Benchmark e fidelidade"], ["2030", "Programa de dados"]];
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
      ["Comunidade", ["Grupo piloto com as 20 fazendas; 1º fechamento de safra", "Grupo de Desenvolvimento lançado; 1º evento anual com premiação", "Evento anual e benchmark entre os membros", "Comunidade nacional de confinadores TN"]],
      ["Fazendas atendidas", ["20", "60", "120", "200"]],
      ["Analistas online", ["0 (Rodrigo)", "3", "6", "10"]],
      ["Licenças de terceiros", ["Migrar o que for web para o TN Manager", "PecCode só para automação", "TGC só onde houver integração", "Licença só para automação"]],
    ];
    lanes.forEach((l, i) => {
      const num = i === 3 || i === 4, h = num ? 0.38 : 0.5;
      const y = 1.78 + [0, 0.55, 1.1, 1.65, 2.08, 2.51][i];
      R(s, { x: 0.5, y, w: 1.5, h, fill: { color: C.navy } });
      T(s, l[0], { x: 0.58, y, w: 1.38, h, fontSize: 9.5, bold: true, color: C.white, valign: "middle" });
      l[1].forEach((c, k) => {
        R(s, { x: x0 + k * cw, y, w: cw - 0.06, h, fill: { color: C.bg } });
        T(s, c, { x: x0 + k * cw + 0.06, y, w: cw - 0.18, h, fontSize: num ? 13 : 8.5, bold: num, color: num ? C.navy : C.text, align: "center", valign: "middle" });
      });
    });
    s.addNotes("Metas de fazendas e analistas são propostas: 20 em 2027 (capacidade do Rodrigo), 60 em 2028, 120 em 2029 e 200 em 2030, alinhado à meta de 200 clientes da v1.0. Analistas na proporção de 1 para cada 20 fazendas. As funcionalidades por ano vêm da visão estratégica original: 2027 versão estável e integração com hardware; 2028 versão madura com NOA e complementos da AgTech Garage; 2029 fidelidade, clube de membros e benchmark; 2030 sustentabilidade, pegada de carbono e recomendação de produto. Comunidade: Grupo de Desenvolvimento inspirado no Fazenda Nota 10 (2020), com fechamento de safra anual e evento com premiação a partir de 2028. Antes, em 2026: business case, pesquisa com 30 fazendas, imersão técnica e PRDs. Licenças: tudo que for web migra para o TN Manager Web; a PecCode segue como parceira para fazendas que quiserem automação, que ainda não temos; manter TGC onde a integração com balança e vagão for indispensável até a versão que a ofereça.");
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
      "Aprovar o Grupo de Desenvolvimento e o evento anual como pilares",
      "Indicar um patrocinador na diretoria",
    ];
    dec.forEach((d, i) => {
      const y = 1.58 + i * 0.5;
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


  await pres.writeFile({ fileName: OUT });
  console.log("ok", OUT);
})();
