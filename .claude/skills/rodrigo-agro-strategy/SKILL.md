---
name: rodrigo-agro-strategy
description: Strategy and research standard for agro and zootecnia work. Use whenever the user asks for business ideas, case studies, decision-making support, SWOT analysis, benchmarking, a business plan, comparisons, frameworks, or a roadmap for producers and entrepreneurs in agro or zootecnia, or asks to research new practices, updated references, Python-based data analysis, or specific modeling for the livestock and agro field. Triggers on "análise SWOT", "benchmarking", "business plan", "estudo de caso", "roadmap", "tomada de decisão", "ideia de negócio no agro", "pesquisa em zootecnia", as well as the English equivalents. Chains into the vendored consulting skill for the business frameworks and into a curated toolkit of Python, genomics, remote sensing, and academic research repositories for the technical grounding.
---

# Rodrigo Agro strategy standards

This is the entry point for business strategy, decision support, and applied research in agro and zootecnia. Read this before answering, and pick the framework or resource that actually fits the question instead of listing everything available.

## 1. Business and decision frameworks

Use the vendored `consulting` skill (`.claude/skills/consulting/SKILL.md`, sourced from `yoichiojima-2/consultant`, MIT licensed) as the framework library. It already maps each situation to the right tool. The most relevant mappings for agro work:

- **Diagnóstico de propriedade ou negócio:** PESTEL for the macro environment, Porter's 5 Forces for the sector, SWOT for the specific operation. Template ready at `consulting/templates/swot.md`.
- **Ideia de negócio nova (confinamento, integração lavoura-pecuária, novo produto):** Business Model Canvas and Business Case, both in `consulting/templates/`, backed by `consulting/finance.md` for NPV, IRR, payback, and break-even on the investment.
- **Benchmarking entre propriedades ou operações:** 3C Framework and Value Chain from `consulting/strategy.md`, plus the Profitability Framework from `consulting/cases.md` for revenue and cost comparisons.
- **Estudo de caso:** the case frameworks in `consulting/cases.md` (profitability, market sizing, market entry) give the structure; close with the Pyramid Principle from `consulting/problem-solving.md` so the answer leads with the recommendation, not the analysis.
- **Roadmap de implementação:** WBS, RACI, and the McKinsey Horizons model from `consulting/projects.md`. Template ready at `consulting/templates/project-charter.md`.
- **Comparativo entre alternativas (raças, sistemas de produção, fornecedores):** GE-McKinsey Matrix or BCG Matrix from `consulting/strategy.md` when the comparison has a volume/attractiveness axis, otherwise a plain weighted decision matrix.

Match effort to the ask, per the `consulting` skill's own rule: a quick question gets one well-chosen framework, not the full engagement flow.

## 2. Data and research toolkit for agro and zootecnia

Reach for these when the answer needs technical grounding rather than a business framework. None of these are installed as skills; they are reference points to search, cite, or pull code patterns from.

**Manipulação e análise de dados em Python**
- pandas: base para qualquer análise tabular (rebanho, safra, financeiro)
- Polars: mesma tarefa em bases grandes, mais rápido que pandas em agregações pesadas
- awesome-python-data-science (krzjoa): lista curada para escolher a ferramenta certa por tipo de problema

**IA aplicada a lavoura e produtividade**
- awesome-agricultural-data-analytics (brandonhimpfen) and awesome-agriculture (brycejohnston): curated entry points for tools and datasets
- crop_yield_prediction (JiaxuanYou): CNN/LSTM and Gaussian Process approach to yield prediction from satellite imagery, a solid reference architecture
- CropHarvest: labeled remote-sensing dataset for crop-type mapping, useful when a client needs a defensible data source

**Pecuária de precisão e zootecnia**
- livestock-vision-ai (diegoreynac): computer vision applied to precision livestock farming
- GitHub topics `cattle`, `dairy-cow`, and `livestock` (filtered to Python): live index of current projects, check before assuming a technique is novel

**Genética e melhoramento animal**
- AlphaSimR: stochastic breeding program simulation down to the DNA sequence level, in R
- genomicMateSelectR and Genomic-Selection (MarcooLopez): genomic prediction and mate selection models
- Breeding-Insight: R Shiny apps for genomic data analysis without a command line

**Geoprocessamento e sensoriamento remoto para pastagem**
- Awesome-GEE (opengeos): curated Google Earth Engine resources in Python
- Global Pasture Watch: annual pasture and grassland maps at 30-meter resolution, 2000 to 2022
- remote-sensing-textbook (calekochenour): full course in Python and JavaScript for reproducible remote-sensing workflows

**Pesquisa acadêmica e referências atualizadas** (direct support for the Unesp master's research)
- Pyzotero: Python client for the Zotero API, for automating reference organization
- awesome-zotero: curated plugin and integration list
- Awesome-Auto-Research-Tools: literature search, paper reading, and experiment management tooling

**Gestão de propriedade**
- GitHub topic `farmmanagementsystem`: live index of open farm management platforms covering crop, herd, inventory, and financial tracking

## 3. Zero AI-writing tells

Every answer, report, or document produced from this skill follows the same rule as the rest of this project: no emojis, no em dashes (—) or double-hyphen substitutes, no filler openers, no cliché vocabulary. Run `avoid-ai-writing` in edit mode on the final text before delivering it.

The vendored `consulting` reference files (section 1) are third-party source material read for their frameworks, not text that ships to the client. Leave them as vendored. The writing rule applies to what you produce from them, not to the library itself.

## 4. Orchestration

1. Identify whether the ask is a business/strategy question, a technical/research question, or both.
2. For strategy: load `consulting`, pick one to three frameworks from section 1, apply them to the user's actual numbers and context. Never present a framework menu instead of an answer.
3. For technical grounding: pull the relevant entry from section 2, note it as a reference, and adapt the approach to the user's data rather than copying code blindly.
4. When the deliverable is a written report, case study, or business plan document, run `avoid-ai-writing` (edit mode) before calling it done, same as any other delivered copy in this project.
5. When the deliverable is a page, dashboard, or app presenting this analysis, hand off to `rodrigo-agro-build-standards` for the UI and UX layer.
6. Lead with the recommendation, then the supporting analysis. State assumptions explicitly when data is missing instead of guessing silently.
