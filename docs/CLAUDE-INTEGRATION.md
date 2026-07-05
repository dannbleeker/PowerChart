# Making PowerChart accessible to Claude in PowerPoint

Research into how Claude can create and edit think-cell-style charts in users'
decks, and which integration path to build first. (Researched July 2026.)

## The landscape

Three facts from Anthropic's current product surface shape the answer:

1. **Claude for PowerPoint exists** — an official Anthropic add-in (research
   preview since Feb 2026; AppSource; Max/Team/Enterprise). It reads the slide
   master and generates **native, editable PowerPoint elements**, and — key —
   **Skills enabled in Claude settings are available inside the add-in**, as
   are **remote MCP connectors** via the sidebar.
2. **Agent Skills are the packaging mechanism** — a folder with `SKILL.md`,
   reference docs, and scripts, uploadable as a ZIP to claude.ai
   (Customize → Skills) and usable across claude.ai, Claude Code, the API
   (`skill_id` + code execution), and Claude for PowerPoint.
3. **Anthropic's own `pptx` skill** builds decks with **PptxgenJS** (from
   scratch) or **raw OOXML manipulation** (template editing) — *not*
   python-pptx — and mandates visual QA via slide thumbnails. This is exactly
   the level PowerChart's scene graph sits at.

## Why PowerChart is unusually well-positioned

The engine was built API-first without knowing it:

- **`ChartConfig` is pure JSON** — precisely the interface an LLM wants to
  emit. The Automation box and `.ppttc`-inspired batch CLI already consume it.
- **The scene graph is renderer-agnostic** — rects, lines, texts, wedges with
  absolute coordinates. Mapping scene → PptxgenJS shapes (or DrawingML XML) is
  mechanical, the same way the Office.js renderer works today.
- **`dist-lib/powerchart.js`** already exposes `buildChart`/`sampleConfig`
  headlessly for Node — the exact runtime a skill's scripts use.

## Integration options, ranked

### Option A — a **PowerChart Agent Skill** (recommended start)

Package the engine as a custom skill:

```
powerchart-skill/
├── SKILL.md            # triggers: waterfall, Mekko, Gantt, "consulting chart",
│                       # think-cell-style; the ChartConfig JSON schema in brief
├── reference.md        # full config reference: kinds, datasheet-row
│                       # conventions (e, 100%=, X extent, Start/End, stacks…),
│                       # decorations, style files
├── scripts/
│   ├── render-svg.mjs  # config JSON → SVG (exists: scripts/render-batch.mjs)
│   └── render-pptx.mjs # config JSON → .pptx with NATIVE shapes (to build:
│                       # scene → PptxgenJS; ~1 day, mirrors the Office.js renderer)
└── dist-lib/powerchart.js
```

Claude's flow: user asks for "an EBITDA bridge from this Excel" → Claude emits
`ChartConfig` JSON → runs `render-pptx.mjs` → native-shape slide, visually QA'd
via the thumbnail pattern from Anthropic's pptx skill.

**Why first:** one artifact covers *every* Claude surface — claude.ai, Claude
Code, the API, and (via settings) **Claude for PowerPoint itself**. No hosting,
no auth, private-by-default, and ~90% of it already exists in this repo.
**Missing piece:** the scene→pptx converter script (the only new code).

### Option B — a **remote MCP connector**

Host a small MCP server exposing `render_chart(config) → pptx/svg` (and
`list_kinds`, `sample_config`). Claude for PowerPoint and claude.ai both take
custom remote MCP connectors, so Claude-in-PowerPoint could call PowerChart
live. Cloudflare Workers would host `dist-lib` trivially.
**Cost:** hosting + auth + connector review friction. **When:** after A, if
team-wide distribution (vs per-user skill upload) matters.

### Option C — **Claude inside the PowerChart pane**

A "Describe your chart…" box in our task pane calling the Claude API
(`claude-sonnet-5`, tool-forced to emit `ChartConfig`), rendering instantly via
the existing preview → Insert. Great demo, but every user needs an API key (or
we need a proxy service), and it only reaches people who installed PowerChart.
**When:** as a showcase feature once A exists.

### Option D — the manual bridge (works today)

Ask any Claude to output PowerChart `ChartConfig` JSON (paste `reference.md`
as context) → paste into the pane's **Automation → Import**. Zero new code —
this is the fallback and the testing path for A.

## Status: Option A is built

The skill lives in `skill/` (SKILL.md, reference.md, `scripts/render-pptx.mjs`).
`npm run skill` assembles and zips it to `skill-dist/powerchart-charts.zip`;
upload at claude.ai → Customize → Skills, and it becomes available in Claude
for PowerPoint too. The pptx renderer emits native shapes with **exact**
adjustable pie geometry (verified in OOXML: `prst="pie"` with correct
adj1/adj2 angles), validated end-to-end with python-pptx.

## Recommendation

1. **Build the scene→pptx converter** (`render-pptx.mjs`, PptxgenJS: rect→
   `addShape("rect")`, line→`addShape("line")`, text→`addText`, wedge→
   `addShape("pie")` — PptxgenJS *does* support angle-adjustable pies, so decks
   generated this way get **exact** pie geometry, better than the live add-in).
2. **Author `SKILL.md` + `reference.md`** condensing the datasheet conventions
   and config schema (most content exists in README/ARCHITECTURE).
3. **Zip and upload** to claude.ai → it appears in Claude for PowerPoint.
4. Revisit B (MCP) for org-wide distribution, C for in-pane NL authoring.

## Sources

- [Use Claude for PowerPoint — Claude Help Center](https://support.claude.com/en/articles/13521390-use-claude-for-powerpoint)
- [Use skills in Claude — Claude Help Center](https://support.claude.com/en/articles/12512180-use-skills-in-claude)
- [Claude for Microsoft 365 with third-party platforms](https://support.claude.com/en/articles/13945233-use-claude-for-microsoft-365-with-third-party-platforms)
- [Agent Skills quickstart — Claude Platform Docs](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/quickstart)
- [anthropics/skills — pptx skill](https://github.com/anthropics/skills/blob/main/skills/pptx/SKILL.md)
- [Claude by Anthropic for PowerPoint — Microsoft AppSource](https://marketplace.microsoft.com/en-us/product/office/wa200010001)
- Reviews/analyses: [Plus AI](https://plusai.com/blog/in-depth-review-of-claude-for-powerpoint/), [MindStudio](https://www.mindstudio.ai/blog/claude-powerpoint-add-in-beta-capabilities), [prezent.ai](https://www.prezent.ai/blog/claude-for-powerpoint)
