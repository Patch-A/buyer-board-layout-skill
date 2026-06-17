# buyer-board-layout

Codex skill for buyer-board PPT template decomposition, buyer research, layout-config generation, content filling, image placement, and final export.

## What this repo contains

- `buyer-board-layout/`: the Codex skill
- `scripts/run_buyer_board_pipeline.py`: one-click unified entry script

## Current maturity

This repository now supports:

- template-based buyer-board generation
- layout-config scaffold generation from a reference PPT
- structured buyer text filling
- country + procurement-need driven buyer research
- public-website buyer asset fetching
- same-domain page expansion from homepage links
- search-engine candidate page discovery as a supplement to site crawling
- image candidate ranking plus size, aspect-ratio, and file-size filtering
- asset cache reuse to avoid repeated downloads from the same site
- per-run asset fetch reporting in `workspace/asset_fetch_report.json`
- social/profile/map candidate-page fallback discovery
- smarter right-side visual preprocessing with whitespace trim, content-aware crop, and extreme-ratio fallback
- optional AI right-side visual fallback via `--enable-ai-visual-fallback`
- separate logo and site-image placement
- PowerPoint COM image placement with preview export
- Python fallback image placement when PowerPoint COM is unavailable
- one-click pipeline execution

## Input modes

### Mode 1: Existing buyers.json

Use this when you already have buyer data prepared.

### Mode 2: Auto-research mode

Use this when you only have:

- a PPT template
- a target country
- a procurement need

The pipeline will:

1. generate `layout-config.json` if you do not provide one
2. research buyers and generate `buyers.json`
3. try to fetch public logo and right-side visuals from each buyer's public web presence
4. fill the PPT
5. place verified image assets when available
6. remove placeholder images when assets are unavailable

## Dependencies for auto-research mode

Install dependencies first:

```bash
pip install -r requirements.txt
```

Set your API key:

```bash
set OPENAI_API_KEY=your_key_here
```

Optional model override:

```bash
set BUYER_RESEARCH_MODEL=gpt-4.1
```

## One-click usage

### Existing buyers.json mode

```bash
python scripts/run_buyer_board_pipeline.py ^
  --template "buyer-board-layout/assets/examples/buyer-manual-reference.pptx" ^
  --buyers "buyer-board-layout/assets/examples/sa-buyers.json" ^
  --layout-config "buyer-board-layout/assets/examples/sa-layout-config.json" ^
  --output "output/finished.pptx" ^
  --preview-dir "output/previews" ^
  --workspace "output/workspace"
```

### Auto-research mode

```bash
python scripts/run_buyer_board_pipeline.py ^
  --template "path/to/template.pptx" ^
  --country "南非" ^
  --procurement-need "动力传动" ^
  --output "output/finished.pptx" ^
  --preview-dir "output/previews" ^
  --workspace "output/workspace"
```

Optional controls:

- `--buyer-count 5`
- `--layout-config path/to/layout-config.json`
- `--cover-title "南非动力传动买家"`
- `--cover-country "国家：南非"`
- `--content-title "南非动力传动买家"`
- `--openai-model gpt-4.1`

## Workspace outputs

The unified pipeline may produce these reusable artifacts inside `--workspace`:

- `buyers.generated.json`: AI-researched buyer list
- `buyers.with-assets.json`: buyer list enriched with fetched asset paths
- `layout-config.generated.json`: starter layout config when one is not supplied
- `asset-cache.json`: site-level asset cache to avoid duplicate fetching
- `asset_fetch_report.json`: per-buyer hit report for logo and right-side visual sourcing
- `assets/`: downloaded public image assets
- `research/`: intermediate buyer research files

## Current boundary

The current V4.4 workflow can automatically generate:

- buyer name
- website
- procurement products
- 120-Chinese-character company bio

It can also continue the PPT pipeline even when no verified logo or right-side image is available.

Current limitations:

- public buyer text research is automated but still depends on model quality and source availability
- public website asset fetching is automatic but best-effort
- search-engine candidate pages supplement official-site crawling, but they still prefer same-domain pages for downloads
- social/profile/map pages are fallback sources, not the first-choice primary source for brand assets
- AI right-side visual fallback is opt-in and requires a valid OpenAI API key
- when no verified image is available, the workflow clears placeholder graphics instead of inventing a risky fake logo
- if a logo asset is SVG, the Python fallback path still requires a working cairo runtime in addition to `cairosvg`

## Privacy note

Do not publish client-specific finished decks, private sample projects, generated preview outputs, or real customer deliverables in the public repo or release assets.

When sharing this skill publicly, prefer:

- blank or redacted templates
- generic `layout-config.json` samples
- sanitized `buyers.json` examples
- workflow documentation rather than real customer deliverables

## Layout-config generator scaffold

Use the generator when you have a new manually adjusted PPT and want a starter config:

```bash
python buyer-board-layout/scripts/generate_layout_config.py ^
  --template "path/to/reference.pptx" ^
  --output "path/to/layout-config.json" ^
  --cover-title "请替换封面标题" ^
  --cover-country "国家：请替换" ^
  --content-title "请替换内容页标题"
```

The generator currently:

- detects likely cover title and country text boxes
- detects the first content table and row labels
- detects likely content title and footer text boxes
- extracts per-slide logo and right-image slot heuristics

This output is a starter scaffold and should still be visually checked before production use.

## Sharing

This repository is public. Other users can open it, clone it, download the ZIP, and install or adapt the skill.

## Feedback loop

Use:

- `Issues` for bugs and feature requests
- `Discussions` for general feedback, usage reports, and template-sharing
- `Releases` for stable downloadable versions
