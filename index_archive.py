"""Index the quote archive into SQLite using Claude for structured extraction.

Walks archive/, extracts text from each .docx/.pdf, sends it to Claude
(claude-sonnet-4-6) with the plain-English rules from rules.md, and stores
quotes + line items in data/quotes.db.

Resumable: files already present in the quotes table are skipped. Use
--force to re-index everything (e.g. after editing rules.md).
Prints an estimated API cost when done.
"""

import argparse
import json
import time
from pathlib import Path

import anthropic

from common import ARCHIVE_DIR, ROOT, db_connect, require_api_key, extract_text

MODEL = "claude-sonnet-4-6"
# Pricing per million tokens (for the cost estimate printout).
PRICE_IN_PER_M = 3.00
PRICE_OUT_PER_M = 15.00

EXTRACT_TOOL = {
    "name": "record_quote",
    "description": "Record the structured contents of one sales quotation document.",
    "input_schema": {
        "type": "object",
        "properties": {
            "quote_number": {"type": ["string", "null"],
                             "description": "Quotation number exactly as printed, e.g. Q-2023-0147"},
            "quote_date": {"type": ["string", "null"],
                           "description": "Quote date in ISO YYYY-MM-DD"},
            "customer": {"type": ["string", "null"],
                         "description": "Customer/company the quote is addressed to, or null"},
            "line_items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "part_number": {"type": ["string", "null"]},
                        "product_name": {"type": ["string", "null"]},
                        "description": {"type": ["string", "null"],
                                        "description": "Technical description text, or null if absent"},
                        "unit_price": {"type": ["number", "null"],
                                       "description": "Per-unit price as plain decimal, no symbols"},
                        "quantity": {"type": ["integer", "null"]},
                    },
                    "required": ["part_number", "product_name", "description",
                                 "unit_price", "quantity"],
                },
            },
        },
        "required": ["quote_number", "quote_date", "customer", "line_items"],
    },
}


def build_system_prompt():
    rules = (ROOT / "rules.md").read_text(encoding="utf-8")
    return (
        "You are a meticulous data-entry clerk extracting structured data from "
        "industrial sales quotations. Follow the extraction rules below exactly. "
        "Record every line item. Use null for genuinely missing fields — never "
        "invent values.\n\n"
        "=== EXTRACTION RULES (rules.md) ===\n" + rules
    )


def extract_quote(client, system_prompt, text, retries=4):
    """One API call -> (parsed dict, input_tokens, output_tokens)."""
    delay = 5.0
    for attempt in range(retries + 1):
        try:
            resp = client.messages.create(
                model=MODEL,
                max_tokens=2048,
                system=system_prompt,
                tools=[EXTRACT_TOOL],
                tool_choice={"type": "tool", "name": "record_quote"},
                messages=[{"role": "user", "content":
                           "Extract this quotation document:\n\n" + text}],
            )
            block = next(b for b in resp.content if b.type == "tool_use")
            return block.input, resp.usage.input_tokens, resp.usage.output_tokens
        except (anthropic.RateLimitError, anthropic.APIStatusError) as e:
            if attempt == retries:
                raise
            print(f"    transient API error ({type(e).__name__}), retrying in {delay:.0f}s...")
            time.sleep(delay)
            delay *= 2


def store(con, filename, data, tok_in, tok_out):
    cur = con.execute(
        "INSERT INTO quotes (file, quote_number, quote_date, customer, "
        "input_tokens, output_tokens) VALUES (?,?,?,?,?,?)",
        (filename, data.get("quote_number"), data.get("quote_date"),
         data.get("customer"), tok_in, tok_out))
    qid = cur.lastrowid
    for it in data.get("line_items") or []:
        pn = it.get("part_number")
        con.execute(
            "INSERT INTO line_items (quote_id, part_number, product_name, "
            "description, unit_price, quantity) VALUES (?,?,?,?,?,?)",
            (qid, pn.upper().strip() if pn else None, it.get("product_name"),
             it.get("description"), it.get("unit_price"), it.get("quantity")))
    con.commit()


def run_index(con, extractor, files, force=False, limit=None, sleep=0.0):
    """Index `files` using `extractor(text) -> (data, tok_in, tok_out)`.

    Resumable: files already present in the quotes table are skipped unless
    force=True (which wipes and re-indexes). Returns
    (indexed, skipped, failures, total_in, total_out).
    """
    already = {row[0] for row in con.execute("SELECT file FROM quotes")}
    if force and already:
        con.execute("DELETE FROM quotes")  # cascades to line_items
        con.commit()
        already = set()

    todo = [f for f in files if f.name not in already]
    skipped = len(files) - len(todo)
    if limit:
        todo = todo[:limit]
    print(f"{len(files)} files in archive; {skipped} already indexed; "
          f"indexing {len(todo)} now.")

    total_in = total_out = 0
    failures = []
    for i, f in enumerate(todo, 1):
        print(f"[{i}/{len(todo)}] {f.name} ...", flush=True)
        try:
            text = extract_text(f)
            data, tok_in, tok_out = extractor(text)
            store(con, f.name, data, tok_in, tok_out)
            total_in += tok_in
            total_out += tok_out
            n = len(data.get("line_items") or [])
            print(f"    ok: {n} line item(s), quote {data.get('quote_number')}")
        except Exception as e:
            failures.append((f.name, str(e)))
            print(f"    FAILED: {e}")
        if sleep and i < len(todo):
            time.sleep(sleep)
    return len(todo), skipped, failures, total_in, total_out


def main():
    ap = argparse.ArgumentParser(description="Index quote archive into SQLite")
    ap.add_argument("--force", action="store_true",
                    help="re-index files even if already in the database")
    ap.add_argument("--limit", type=int, default=None,
                    help="index at most N new files (for testing)")
    ap.add_argument("--sleep", type=float, default=0.5,
                    help="seconds to pause between API calls (politeness)")
    args = ap.parse_args()

    require_api_key()
    client = anthropic.Anthropic()
    con = db_connect()
    system_prompt = build_system_prompt()

    files = sorted(p for p in ARCHIVE_DIR.iterdir()
                   if p.suffix.lower() in (".docx", ".pdf"))
    if not files:
        raise SystemExit(f"No .docx/.pdf files found in {ARCHIVE_DIR}. "
                         "Run generate_archive.py first.")

    extractor = lambda text: extract_quote(client, system_prompt, text)
    _, _, failures, total_in, total_out = run_index(
        con, extractor, files, force=args.force, limit=args.limit,
        sleep=args.sleep)

    # Cost estimate for THIS run + cumulative across the whole DB.
    run_cost = total_in / 1e6 * PRICE_IN_PER_M + total_out / 1e6 * PRICE_OUT_PER_M
    db_in, db_out = con.execute(
        "SELECT COALESCE(SUM(input_tokens),0), COALESCE(SUM(output_tokens),0) "
        "FROM quotes").fetchone()
    db_cost = db_in / 1e6 * PRICE_IN_PER_M + db_out / 1e6 * PRICE_OUT_PER_M

    nq, ni = con.execute(
        "SELECT (SELECT COUNT(*) FROM quotes), (SELECT COUNT(*) FROM line_items)"
    ).fetchone()
    print(f"\nDone. Database now holds {nq} quote files, {ni} line items.")
    print(f"This run:  {total_in:,} in / {total_out:,} out tokens  "
          f"~= ${run_cost:.4f}")
    print(f"All time:  {db_in:,} in / {db_out:,} out tokens  ~= ${db_cost:.4f}  "
          f"(rates: ${PRICE_IN_PER_M}/M in, ${PRICE_OUT_PER_M}/M out)")
    if failures:
        print(f"\n{len(failures)} file(s) FAILED (re-run to retry just these):")
        for name, err in failures:
            print(f"  {name}: {err}")


if __name__ == "__main__":
    main()
