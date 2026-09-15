# Projetos Rodrigo Agro: instruções do projeto

Repositório do dashboard de BI Comercial (`index.html`, `data.json`), de sites, landing pages, apps e relatórios relacionados ao negócio, e de apoio à consultoria em agro e zootecnia do Rodrigo.

## Regra permanente para qualquer site, landing page, app, dashboard ou relatório

Sempre que a tarefa envolver criar, corrigir, ajustar ou melhorar um site, landing page, aplicativo, dashboard ou relatório neste repositório, use a skill `rodrigo-agro-build-standards` (`.claude/skills/rodrigo-agro-build-standards/SKILL.md`) antes de entregar o trabalho. Ela define o stack de componentes, o stack de motion design/cinematográfico e o checklist de UX e conversão por tipo de projeto.

## Regra permanente para estratégia, decisão e pesquisa em agro/zootecnia

Sempre que a tarefa envolver análise SWOT, benchmarking, business plan, estudo de caso, roadmap, comparativo entre alternativas, ideia de negócio, ou pesquisa de práticas atualizadas para produtores e empreendedores do agro, use a skill `rodrigo-agro-strategy` (`.claude/skills/rodrigo-agro-strategy/SKILL.md`). Ela encadeia a skill `consulting` (frameworks de McKinsey, BCG, Bain e Accenture, vendorizada de `yoichiojima-2/consultant`, licença MIT) para a parte de estratégia, e um conjunto de repositórios de referência em Python, genética, sensoriamento remoto e pesquisa acadêmica para a parte técnica.

Duas regras não têm exceção, mesmo em edições rápidas:

1. **Nenhum emoji.** Em nenhum lugar: títulos, botões, mensagens, textos de relatório, comentários de código.
2. **Nenhum travessão (—) nem substituto com hífen duplo (--).** Usar ponto, vírgula ou reescrever em duas frases.

Antes de marcar qualquer entrega como pronta, rode a skill `avoid-ai-writing` em modo `edit` sobre todo o texto voltado ao usuário final.

## Stack atual

Site estático em HTML/CSS/JS puro, sem bundler. Chart.js, SheetJS (xlsx) e jsPDF carregados via CDN diretamente no `index.html`. Design tokens em CSS custom properties (`--bg`, `--text`, `--border` etc.), com suporte a tema claro/escuro via `[data-theme]`.

Não introduzir React, Tailwind com build step, ou outra stack sem pedido explícito do usuário. Ver a matriz de decisão em `rodrigo-agro-build-standards` para quando cada alternativa (shadcn/ui, Tremor, Mantine, Aceternity UI, Magic UI, Tamagui, NativeWind, Gluestack-UI) se aplica, e a tabela de motion design (GSAP, Lenis, Motion, SplitType, CountUp.js, Three.js/React Three Fiber, Rive, Lottie, Locomotive Scroll) para quando cada biblioteca de animação entra em jogo.
