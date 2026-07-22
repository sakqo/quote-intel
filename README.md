# Quote Intel — Document Intelligence for Quote Generation

Automates the "find the last time we quoted this part and copy it into a new
quote" workflow. A fake archive of messy past quotes (.docx + .pdf) is
indexed into SQLite using Claude for structured extraction; new quotes are
generated from the most recent archived data per part number.

## Pipeline

```
generate_archive.py   ->  archive/*.docx + *.pdf   (+ data/answer_key.json, test-only)
index_archive.py      ->  data/quotes.db           (Claude extraction, driven by rules.md)
make_quote.py         ->  output/Q-YYYY-NNNN.docx  (most recent price/description wins)
test_extraction.py    ->  accuracy report vs. answer key (>= 95% gate)
```

## Setup

```powershell
pip install -r requirements.txt
```

API key (needed only for indexing): set `ANTHROPIC_API_KEY` one of these ways —

```powershell
# this session only
$env:ANTHROPIC_API_KEY = "sk-ant-..."
# permanently (new shells)
[Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

or put `ANTHROPIC_API_KEY=sk-ant-...` in a `.env` file next to the scripts
(gitignored; a real env var wins over `.env`).

## Usage

```powershell
# 1. Generate the fake archive (default 25 quotes; scale up for stress tests)
python generate_archive.py            # or --count 300 --seed 7

# 2. Index it (resumable: re-running skips already-indexed files;
#    --force re-extracts everything, e.g. after editing rules.md)
python index_archive.py               # prints estimated API cost when done

# 3. Generate a quote
python make_quote.py --parts VLV-2043,PMP-118 --customer "Jane Doe" --company "Apex Fabrication LLC"
#    optional: --qty 3,1   --allow-partial

# 4. Validate extraction against ground truth (the 95% gate)
python test_extraction.py

# 5. Unit tests (no API calls)
python -m pytest tests/ -q
```

Unknown part numbers fail gracefully: the script lists what was found,
suggests close matches (fuzzy), and writes nothing unless `--allow-partial`.

## Live demo website

`app.py` is a small Flask app that puts the whole thing in a browser —
searchable part list, per-part quote history (most recent highlighted),
"did you mean" fuzzy suggestions, and one-click quote generation with an
HTML preview and .docx download.

```powershell
python app.py            # local, http://127.0.0.1:5001
```

Demo properties:

- **Read-only** — no extraction API calls at runtime; page queries use a
  read-only SQLite connection (`mode=ro`); generated documents exist only in
  memory behind a random token that expires after 10 minutes.
- **Bundled data** — `data/quotes.db` (pre-indexed) and
  `data/validation_report.json` (written by `test_extraction.py`) ship with
  the repo so the site shows the honest accuracy numbers, including the two
  documented misses on the "How it works" page.
- **Abuse protection** — 20 quote generations per IP per hour (HTTP 429
  after), and customer/company inputs are sanitized (length-capped,
  restricted character set) before touching the template.
- **Deploy on Render free tier** — `render.yaml` is included: connect the
  repo at https://render.com, and it builds with `pip install -r
  requirements.txt` and serves via `gunicorn app:app`. No environment
  variables needed (the demo never calls the API).
- Frontend is plain HTML/CSS (`templates/` + `static/style.css`); all
  presentation lives in the stylesheet so a redesign never touches Python.

Update `GITHUB_URL` at the top of `app.py` once the repo has its permanent
home.

## Customizing rules.md

`rules.md` is a plain-English rules file loaded verbatim into the extraction
prompt — the owner edits it, no code changes needed. It controls:

- **Field synonyms** — e.g. teach it that "Stock Code" also means part number
  by adding it to the synonyms list.
- **Layout interpretation** — how table/paragraph/letter quotes are read,
  including wrapped PDF table rows.
- **Price/date normalization** — output formats, currency assumptions,
  US vs EU date order.
- **Missing data policy** — output `null`, never invent values.
- **Tie-breaking** — "most recent quote date wins" (change it to e.g.
  "highest price wins" if that suits your business).

After editing, re-run `python index_archive.py --force` so existing rows are
re-extracted under the new rules, then `python test_extraction.py` to confirm
accuracy held.

## Data integrity

`data/answer_key.json` is ground truth written by the generator and read
ONLY by `test_extraction.py`. The indexer/extractor never touches it.

## Accuracy report (final validation run, 2026-07-21)

Archive: 25 quotes / 112 line items, each saved as both .docx and .pdf
(50 files, all indexed). Model: `claude-sonnet-4-6`. Indexing cost for the
full archive: **~$0.60** (108,664 input / 18,144 output tokens).

| Field        | Correct | Total | Accuracy |
|--------------|--------:|------:|---------:|
| part_number  |     112 |   112 |   100.0% |
| product_name |     110 |   112 |    98.2% |
| description  |     112 |   112 |   100.0% |
| unit_price   |     112 |   112 |   100.0% |
| quantity     |     112 |   112 |   100.0% |
| quote_date   |      50 |    50 |   100.0% |
| customer     |      50 |    50 |   100.0% |
| **OVERALL**  | **658** | **660** | **99.7%** |

**PASS** — 99.7% ≥ the 95% requirement. Zero spurious (hallucinated) line
items and zero missed line items.

The only 2 failures (both `product_name`, both PDFs): pdfplumber renders a
wrapped table cell across physical lines, and the model dropped the wrapped
fragment — e.g. `Centrifugal Transfer Pump / 1.5HP` extracted as
`Centrifugal Transfer Pump`. The .docx twins of the same quotes extracted
perfectly. If this matters for your archive, strengthen the "reassemble
wrapped cells" rule in `rules.md` with examples and re-index with `--force`.

At quote-generation time, if the most recent occurrence of a part is missing
its name/description (messy source), `make_quote.py` backfills from the most
recent occurrence that has one — the price still always comes from the most
recent quote.
