---
name: rodrigo-agro-build-standards
description: Master build standard for this project. Use whenever the user asks to create, build, fix, redesign, adjust, or improve a website, landing page, dashboard, report, or app in this repository, including "criar site", "montar landing page", "corrigir dashboard", "ajustar relatório", "melhorar conversão", "revisar UX", "animação cinematográfica", "motion design", "build app", "fix the site". Selects the component, motion, and cinematic-scroll stack for the project type, strips every AI-writing tell from delivered copy (emojis, em dashes, AI clichés), and runs a UX and conversion checklist before anything is marked done. Chains into avoid-ai-writing, web-design-guidelines, accessibility, core-web-vitals, dataviz, scroll-experience, 3d-web-experience, and the relevant CRO skill (form-cro, page-cro, popup-cro, signup-flow-cro, onboarding-cro, paywall-upgrade-cro) as needed.
---

# Rodrigo Agro build standards

This is the entry point for any site, landing page, app, or report built in this repository. Read this before writing code or copy, and run the delivery gate at the end before calling the work done.

## 1. Pick the stack for the project type

This repo currently ships as static HTML/CSS/JS (`index.html`, `data.json`) with Chart.js, SheetJS, and jsPDF loaded from CDN. No bundler, no React. Match new work to that reality unless the user asks for a framework migration.

**Landing pages and marketing sites**
- Structure and styling: Tailwind CSS (CDN build if no bundler exists yet)
- Visual hero blocks (3D cards, spotlight, glow effects): Aceternity UI patterns, reimplemented in plain CSS/JS when the project has no React
- Reusable animated sections (marquees, bento grids, animated lists): Magic UI patterns, same vanilla-adaptation rule
- Motion and cinematic scroll: see section 2 below for the full toolkit
- If the user requests a React/Next.js rebuild: shadcn/ui as the component base, Tailwind CSS as the styling layer

**Dashboards and interactive reports** (this covers the current `index.html`)
- Data visualization: keep Chart.js already in use; apply the `dataviz` skill for chart-type selection, color palette, and legend accessibility on every new chart
- Layout tokens: keep the existing CSS custom-property system (`--bg`, `--text`, `--border`, etc.) rather than introducing a second design-token source
- Motion: functional only, from the dashboard-safe list in section 2. Hover states, number count-ups, panel transitions. No decorative 3D or particle motion on data screens
- If the user requests a React rebuild: Tremor for the chart layer, Mantine for the surrounding admin shell

**Apps (mobile or cross-platform)**
- Styling: Tamagui or NativeWind, whichever the user's existing app already uses
- Components: Gluestack-UI
- Motion: Reanimated

## 2. Motion design and cinematic layer

Pick from this toolkit for scroll storytelling, reveals, and cinematic feel. Every entry below is framework-agnostic unless marked otherwise, so it drops into this repo's plain `<script>` setup without a bundler.

| Library | Role | Where it belongs |
|---|---|---|
| GSAP + ScrollTrigger | Sequencing, timelines, scroll-driven reveals | Any project type. The default engine for section 1 |
| Lenis | Momentum smooth scroll, syncs with GSAP's ticker | Landing pages and institutional sites, optional on dashboards |
| Motion (Motion One / Framer Motion) | Lightweight animation on the Web Animations API | Vanilla or React, wherever GSAP is heavier than needed |
| SplitType | Splits text into lines, words, and characters for reveal animation | Landing page headlines, dashboard KPI labels |
| CountUp.js | Animates a number counting up to its final value, zero dependencies | Dashboard and report metrics, the dashboard-safe default |
| Three.js / React Three Fiber (`3d-web-experience` skill) | 3D scenes, product configurators, immersive heroes | Institutional landing pages only, never the BI dashboard |
| Rive | State-driven vector animation that can react to live data | Icons or illustrations tied to a data value, used sparingly |
| Lottie / dotLottie | Lightweight After Effects exports | Loaders and micro-interactions, landing pages and apps |
| Locomotive Scroll | Smooth scroll with built-in parallax | Alternative to Lenis when parallax layers are needed |

**Default combo for this repo:** GSAP, Lenis, CountUp.js, and SplitType. All four load from CDN, none require a bundler or React, and together they cover scroll reveal, smooth navigation, animated metrics, and headline reveal without adding weight to the BI dashboard. Reserve Three.js, React Three Fiber, Rive, and Lottie for a future institutional site, not the BI dashboard itself: dashboards need fast, clear reading over visual spectacle.

For any scroll-driven storytelling section (landing pages, case studies, annual report style pages), load the `scroll-experience` skill. For anything involving 3D scenes or WebGL, load the `3d-web-experience` skill instead of hand-rolling Three.js setup.

## 3. Zero AI-writing tells

Apply to every string that ships: headings, buttons, labels, chart titles, empty states, error messages, emails, report text.

- No emojis anywhere. Not in headers, buttons, toasts, chat bubbles, commit messages, or code comments.
- No em dashes (—) and no double-hyphen substitute (--). Use a period, a comma, or rewrite as two sentences.
- Cut cliché vocabulary on sight: robusto, abrangente, perfeito, revolucionário, sinergia, elevar, capacitar, transformador, de ponta, and their English equivalents (delve, seamless, leverage, unlock, elevate, empower, robust, comprehensive, cutting-edge, game-changer, best-in-class).
- No filler openers ("Bem-vindo ao futuro de...", "Neste mundo em constante evolução..."). Start with the specific value or number.
- No rule-of-three padding, no bolded inline list headers repeating themselves, no uniform paragraph lengths.
- Before marking any copy final, run the `avoid-ai-writing` skill in `edit` mode over the exact file or text block. This is not optional. It is the last pass, every time.

## 4. UX and conversion checklist

Apply inline while building, not as an afterthought:

- One primary call to action per screen. On landing pages it sits above the fold.
- Mobile-first layout, verified at 375px width minimum before desktop.
- WCAG 2.1 AA contrast and full keyboard navigation. Run the `accessibility` skill on any new interface.
- Core Web Vitals in mind from the first draft: lazy-load below-the-fold media, avoid layout shift. Run `core-web-vitals` before delivery on anything performance-sensitive.
- Forms: minimum required fields, inline validation, specific error copy. Route to `form-cro` (general forms), `signup-flow-cro` (account creation), or `popup-cro` (modals/overlays) depending on what the form actually is.
- Real data over placeholders. This project already uses real client names and figures; keep that standard, never ship lorem ipsum in delivered work.
- Microcopy is specific: "Enviar proposta" beats "Clique aqui", a number beats a vague claim.
- Dark/light parity wherever the project already themes via `[data-theme]`. Check both before delivery.

## 5. Orchestration: which skills to chain and when

1. Identify the project type first (landing page, dashboard/report, or app) and pick the component stack from section 1 and the motion stack from section 2.
2. Build with the section 4 checklist applied as you go, not reviewed only at the end.
3. When the work involves scroll storytelling, parallax, or a cinematic page flow, load `scroll-experience`. When it involves a 3D scene or WebGL, load `3d-web-experience` instead of hand-rolling Three.js setup.
4. Before delivery, always run:
   - `avoid-ai-writing` (edit mode) on every piece of user-facing copy
   - `accessibility` on any new or changed interface
   - `web-design-guidelines` or `web-quality-audit` on the finished UI
5. Run the matching specialist skill when the task fits it: `page-cro` for a marketing page that needs to convert better, `form-cro` / `signup-flow-cro` / `popup-cro` / `onboarding-cro` / `paywall-upgrade-cro` for the specific flow named, `dataviz` for any chart or report screen, `core-web-vitals` for anything load-time sensitive.
6. Report back in the same voice this rule enforces: no emoji, no em dash, specific claims over generic ones.

## 6. Delivery gate

Confirm every line before saying the work is done:

- [ ] No emoji anywhere in the delivered UI or copy
- [ ] No em dash or double-hyphen anywhere in the delivered UI or copy
- [ ] `avoid-ai-writing` ran on the final copy and found nothing left to fix
- [ ] One clear primary action per screen, placed above the fold where relevant
- [ ] Contrast and keyboard navigation checked
- [ ] Verified at mobile width
- [ ] Component stack matches section 1 for this project type, or the vanilla-compatible equivalent already used in this repo
- [ ] Motion stack matches section 2: no 3D, particles, or decorative motion on the BI dashboard unless explicitly requested
