# Architecture

BidAgent is a local Python command-line application. Its build path is organized into input parsing, research, evidence checks, review, and document assembly.

## Main parts

| Part | Modules | Responsibility |
|---|---|---|
| Command line and configuration | `cli.py`, `config.py` | Parse commands and load project-relative settings |
| Input processing | `bom.py`, `util.py` | Read BOM rows and extract supported tender file text |
| Product knowledge | `ingest.py`, `library.py`, `datasheets.py` | Build the reusable product library and index optional local datasheets |
| AI and web research | `llm.py`, `web_research.py`, `agents/` | Call configured models, ground BOM research with Google Search, and draft tender-specific sections |
| Compliance and review | `agents/compliance.py`, `agents/reviewer.py` | Compare requirements to evidence and produce a review checklist |
| Word generation | `assembler.py` | Copy the user's template, insert proposal sections, tables, images, and source references |

## Build sequence

```text
BOM + tender files
        │
        ├── BOM parser ── product matching ── local library
        │                                  └── optional PDF index
        │                                  └── Gemini + Google Search grounding
        ├── Tender reader ── project facts + item requirements
        ├── Writer ── introduction sections
        ├── Compliance checker + static/AI reviewer
        └── Word assembler ── proposal.docx + review.md
```

The orchestrator is `pipeline.py`. Each major stage saves results under `work/<bid>/steps/`; model responses are cached under `work/.cache/`. A `--force` build ignores step results for that bid.

## Product data lifecycle

Ingested products seed the local library. Products researched from current web sources are stored as draft entries and included in the proposal by default. The review report flags draft entries; a reviewer should inspect the linked source and verify the offered specifications before approving a product. Approval makes that library entry available for later bids.

The web research call is made once per BOM model/item combination for a bid. Search results are stored in the per-bid step cache, so rerunning an unchanged bid does not automatically refresh them. Use `--force` to run research again.

## Local versus shared data

The repository contains code and non-sensitive defaults. Bid inputs, product libraries, local datasheets, generated work files, output documents, and API keys belong in ignored local paths. Do not commit confidential tender inputs or company product data.
