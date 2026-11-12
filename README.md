# BidAgent

BidAgent turns a bill of materials (BOM) and tender requirements into a draft technical bid in Word. It matches BOM items to a reusable product library, researches current product information with Gemini and Google Search, checks requirements against product evidence, and creates a review report.

Generated proposals are drafts. Review every web-researched specification, compliance result, and highlighted placeholder before submission.

## What it does

- Reads BOM workbooks and tender information from PDF, Word, text, or Excel files.
- Searches Google for each BOM model when building a new bid and carries source links into the proposal.
- Reuses product descriptions, features, pictures, company text, and formatting from an ingested bid.
- Builds a compliance statement and a separate review report for gaps and checks.
- Caches work per bid so an unchanged rerun can reuse prior results.

Local product datasheet PDFs are optional. A tender document or equivalent requirements text is still needed for tender-specific analysis; public web search cannot know private bid requirements.

## Project layout

```text
.
├── src/
│   └── bidagent/           # Python package: CLI, pipeline, agents, library, and Word builder
├── tests/                  # Unit and end-to-end tests
├── docs/                   # Architecture and workflow notes
├── knowledge/              # User-maintained product rules (safe to publish if non-sensitive)
├── config.yaml             # Models, providers, paths, tender limits, and proposal sections
├── requirements.txt        # Runtime dependencies for pip
├── pyproject.toml          # Installable package and CLI entry point
└── .env.example            # Key names only; copy to local .env
```

Bid-specific files are local and ignored by Git: `.env`, `library/`, `datasheets/`, `work/`, and `output/`.

## Requirements

- Python 3.11 or newer
- A Gemini API key for Google Search grounding and online research
- A finished bid `.docx` to ingest as your formatting template and initial product library

The current Google Search integration uses Gemini's native Google Search grounding. Google's Custom Search JSON API is closed to new customers, so this project does not require a Custom Search Engine ID.

## Install

From the repository root, create a virtual environment and install the package:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
```

If the `py` launcher cannot find your installed Python, create the environment with the full path to `python.exe`, then continue with the activation and install commands. For example:

```powershell
& "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe" -m venv .venv
```

## Configure API keys

Copy the example file and add your key:

```powershell
Copy-Item .env.example .env
```

Edit `.env` and set `GEMINI_API_KEY`. Create or manage the key in [Google AI Studio](https://aistudio.google.com/api-keys). The project uses Gemini for Google Search grounding and other AI steps. Optional Groq and OpenRouter keys can be set for model fallbacks.

Do not commit `.env` or paste API keys into `config.yaml`. Free-tier limits and availability depend on the model and Google project. Check the current [Gemini API pricing](https://ai.google.dev/gemini-api/docs/pricing) and [rate limits](https://ai.google.dev/gemini-api/docs/rate-limits). Free-tier data handling may differ from paid tiers; check Google's terms before sending confidential tender information.

## First-time setup

1. Check that the configured models and API key are available:

   ```powershell
   python -m bidagent doctor
   ```

2. Ingest one of your finished bids to create the Word template and seed the product library:

   ```powershell
   python -m bidagent ingest "My finished bid.docx"
   ```

3. Review products created under `library/products/`. Ingested product specifications should be checked against current manufacturer information before use.

The template can use `{{PROJECT_NAME}}`, `{{CLIENT}}`, `{{TENDER_NO}}`, `{{DATE}}`, and `{{LOCATION}}` placeholders. Configure retained template sections and proposal headings in `config.yaml`.

## Build a bid

```powershell
python -m bidagent build `
  --bom "MOD BOM.xlsx" `
  --tender "tender.docx" `
  --name "MOD Antivirus" `
  --client "MOD"
```

The `--tender` option accepts one or more supported files. If a tender is long, the default configuration reads up to 32 text chunks; increase `tender.max_chunks` in `config.yaml` if the review report says later content was skipped.

The build creates:

- `output/<name>.docx` — proposal assembled in your ingested Word template.
- `output/<name>_review.md` — issues, compliance gaps, warnings, draft products, and AI usage.

Read the review report first. Web-researched products are saved as drafts; check their sources and specifications, then approve a product for future bids:

```powershell
python -m bidagent library
python -m bidagent approve <product-id>
```

## How the pipeline works

1. **Read inputs:** extract BOM rows and tender facts and requirements.
2. **Match products:** compare BOM models against the local library.
3. **Research products:** use Gemini's Google Search grounding and keep URLs returned in the grounding metadata as source references. Local datasheets can also be indexed when supplied.
4. **Draft proposal sections:** write project introduction sections from tender facts and the company's saved writing style.
5. **Check compliance:** compare item requirements with library or researched product features, and flag unsupported or weak matches.
6. **Review and assemble:** create the Word proposal and a Markdown checklist for human review.

For more detail, see [the architecture notes](docs/architecture.md).

## Other commands

```powershell
python -m bidagent index             # inspect optional local datasheets
python -m bidagent usage             # show AI usage and cache counts
python -m bidagent build --help      # see build options
python -m pytest                    # run tests
```

Use `--no-llm` for an offline placeholder build and `--force` to ignore saved step results and research again.

## Evidence and limitations

- Search is based on Google's indexed results; it cannot guarantee that every website is searched or that every source is authoritative.
- The project checks that a web source URL came back in Google's grounding metadata, but it does not independently verify every web claim against the page. Confirm specifications on the linked manufacturer source before approval.
- Datasheet PDF extraction checks its quotes against the local PDF's extracted text. Scanned image-only PDFs need OCR.
- Numeric checks catch only some unit and threshold conflicts. Review every compliance result and requirement marked for confirmation.
- The generated proposal is not ready to submit until a person checks the tender interpretation, product offer, prices, delivery, warranties, legal declarations, and all placeholders.

## License

No license has been added yet. Add a license file before inviting reuse or contributions.
