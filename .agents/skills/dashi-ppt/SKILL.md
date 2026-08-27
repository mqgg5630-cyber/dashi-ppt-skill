---
name: dashi-ppt
description: Creates Chinese or English presentation decks as editable offline HTML, PPTX, or PDF using the Dashi PPT MCP tools. Use when a user asks for a PPT, presentation, slide deck, report deck, pitch deck, or editable PowerPoint.
---

# Dashi PPT for Antigravity

This is the Antigravity workspace adapter for the Dashi PPT skill. The full renderer and 12-theme layout library live in `skills/dashi-ppt/`; the workspace MCP server in `mcp/server.mjs` exposes the safe operations below. Prefer MCP tool calls instead of manually reconstructing the renderer.

## Tool-first workflow

1. Clarify the topic, audience, page count, visual theme, media intent, language, and final format. If the user has not delegated the choice, ask for a theme before generating. “做 PPT” means a presentation/HTML deck; only an explicit PPTX, PowerPoint, editable PPTX, export PPTX, or PPT file request requires a PPTX file.
2. Call `dashi_ppt_scaffold` with `title`, `goal`, `theme`, and `pages`. Add `audience`, `owner`, `roles`, `content_briefs`, `seed`, and media intent when known. The tool returns a new workspace-relative `goal_path`; never reuse an old `output/` goal file.
3. Write complete, user-specific copy into the returned `goal.json`. Keep each slide’s `content.presentation` as the canonical fact source. For a normal delivery, every slide must contain real title/body/metric/item copy rather than relying on template defaults. Use the existing repository editing tools or shell only for this content-authoring step.
4. Use `dashi_ppt_layout_query` to find alternatives and `dashi_ppt_inspect_layout` before writing unfamiliar object, array, count, or media props. Keep three template variants and one bespoke variant in schema version 2 unless the user explicitly requests a legacy one-layout deck.
5. Stage user media with `dashi_ppt_stage_media` and use the returned relative paths in `goal.json`. Do not put temporary paths, absolute paths, `file://` URLs, or remote URLs in the deck spec.
6. Call `dashi_ppt_render` only after the goal spec is complete. It runs safe-props normalization and the built-in validation gates, writes `ppt/index.html`, and starts a local preview server.
7. If the user wants a file, call `dashi_ppt_export` with `format: "pptx"` or `format: "pdf"`. For HTML delivery, return the preview URL from the render result. Do not claim an export succeeded unless the returned file exists.

## Available MCP tools

- `dashi_ppt_scaffold`: theme-aware goal JSON and layout allocation.
- `dashi_ppt_layout_query`: capacity/role/media-aware layout search.
- `dashi_ppt_inspect_layout`: copy keys, fill plan, prop shapes, media slots, and controls.
- `dashi_ppt_stage_media`: copies local image/video assets into the deck.
- `dashi_ppt_render`: renders and validates the editable HTML deck.
- `dashi_ppt_preview`: starts or restarts a preview for an existing rendered deck.
- `dashi_ppt_export`: exports editable PPTX or PDF.
- `dashi_ppt_validate`: validates a goal spec and, when supplied, its rendered HTML.

All MCP paths are relative to the cloned repository. Keep generated decks in `output/<deck-name>/`; that directory is intentionally ignored by Git. Node.js 20+ and npm are required. PPTX/PDF export additionally needs Chrome, Chromium, or Edge; set `CHROME_PATH` if the browser is not auto-detected.

## Theme selection

- `theme01` — soft neumorphism; product and company updates
- `theme02` — luminous purple/green; AI, robotics, and technology launches
- `theme03` — light/dark code; technical architecture and developer content
- `theme04` — glass candy; youth-oriented brands and consumer products
- `theme05` — spectrum charts; data reports and market analysis
- `theme06` — dark atlas; dense strategic, industry, or financial analysis
- `theme07` — cool white research; surveys, white papers, and policy reports
- `theme08` — black and gold experimental; premium launches and brand proposals
- `theme09` — navy magazine; brand stories, interviews, and profiles
- `theme10` — gold index; finance, investment, and ranking reports; do not auto-select unless clearly appropriate
- `theme11` — high-energy growth; business plans, growth reviews, and fundraising
- `theme12` — soundwave neon; music, entertainment, and youth events

## Content and quality rules

- Follow the user’s language. For English decks set the top-level `language` to `en` and replace all default Chinese copy.
- Use the repository’s schema-v2 contract: one logical slide with three distinct template layouts plus one `bespoke` composition. Keep the three templates grounded in the same facts; the bespoke variant may reorganize emphasis but must not invent facts.
- Put real content in `slide.content.presentation` before selecting layouts. Respect `fillPlan.text`, `fillPlan.arrays`, `copyBudgets`, `numericBounds`, and `mediaSlots` returned by inspection. Use `<br>`, `<b>`, and `<em>` only where the renderer permits HTML copy; never inject free-form HTML or CSS.
- Do not edit theme source, component CSS, class names, or generated metadata to solve an ordinary content request. Use props and the goal spec.
- Keep media usage unique to one logical slide. Before delivery, check that every referenced asset exists under the rendered deck and that no unrelated default copy remains.
- Validate the narrative (opening, evidence, conclusion/action), visible copy, chart insights, page count, and output file. A passing script is a technical baseline, not a substitute for checking whether the deck answers the user’s request.

## Checked-in 12-theme samples

`examples/theme-samples/` contains one eight-slide Chinese sample for every bundled theme. The same `briefs.json` content source is paired with each MCP scaffold goal and native editable PPTX. From the repository root, regenerate the goals through the local MCP server and then regenerate and verify the PPTX files with:

```bash
node examples/theme-samples/scaffold-all.mjs
node examples/theme-samples/generate-all.mjs
node examples/theme-samples/verify-all.mjs
```

The sample PPTX generator keeps text, shapes, lines, and chart-drawing primitives native and uses an offline OOXML fallback when the optional PptxGenJS package is not installed. These checked-in samples use `layout_variants: 1` to make all twelve theme packs comparable; normal user deliveries should still use schema-v2 three-template plus bespoke variants unless the user asks for a single-layout deck.

## Fallback when MCP is unavailable

The same repository can be used without MCP. Run commands from the repository root:

```bash
node skills/dashi-ppt/project/scripts/goal-scaffold.mjs --title "标题" --goal "汇报目标" --theme theme07 --pages 10 --out output/example/goal.json
bash skills/dashi-ppt/scripts/render_goal_deck.sh output/example/goal.json output/example/ppt/index.html
node skills/dashi-ppt/project/scripts/export-pptx.mjs output/example/ppt output/example/example.pptx
```

On Windows, use `skills/dashi-ppt/scripts/render_goal_deck.ps1` instead of the shell script. The legacy full instructions remain in `skills/dashi-ppt/SKILL.md`; read that file when the task needs detailed fill-plan, media, or visual-QA rules.
