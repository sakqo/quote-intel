"""Validation gate: compare SQLite extraction results against data/answer_key.json.

Scores per-field accuracy (part number, product name, description, unit price,
quantity, plus quote-level date and customer) across every archived file, and
an overall accuracy. Requirement: >= 95% overall. Failures are printed grouped
by cause so rules.md / extraction logic can be improved.

This is the ONLY code allowed to read answer_key.json.
"""

import difflib
import json
import re
import sys
from collections import defaultdict

from common import ROOT, db_connect

ANSWER_KEY = ROOT / "data" / "answer_key.json"
THRESHOLD = 0.95


def norm_text(s):
    if s is None:
        return None
    return re.sub(r"\s+", " ", s).strip().lower()


def text_match(expected, got, fuzz=0.92):
    """Exact after whitespace/case normalization, or >= fuzz similarity."""
    e, g = norm_text(expected), norm_text(got)
    if e == g:
        return True
    if e is None or g is None:
        return False
    return difflib.SequenceMatcher(None, e, g).ratio() >= fuzz


def price_match(expected, got):
    if expected is None and got is None:
        return True
    if expected is None or got is None:
        return False
    return abs(expected - got) < 0.005


def load_db_quotes(con):
    """{file: {quote row..., items: [...]}}"""
    out = {}
    for qid, file, qnum, qdate, cust in con.execute(
            "SELECT id, file, quote_number, quote_date, customer FROM quotes"):
        items = [dict(zip(("part_number", "product_name", "description",
                           "unit_price", "quantity"), row))
                 for row in con.execute(
                     "SELECT part_number, product_name, description, unit_price, "
                     "quantity FROM line_items WHERE quote_id = ?", (qid,))]
        out[file] = {"quote_number": qnum, "quote_date": qdate,
                     "customer": cust, "items": items}
    return out


def validate(db_quotes, key):
    """Returns (field_stats, failures).

    field_stats: {field: [correct, total]}
    failures: {cause: [message, ...]}
    """
    stats = defaultdict(lambda: [0, 0])
    failures = defaultdict(list)

    def score(field, ok, cause, msg):
        stats[field][1] += 1
        if ok:
            stats[field][0] += 1
        else:
            failures[cause].append(msg)

    for quote in key["quotes"]:
        for fname in quote["files"]:
            got = db_quotes.get(fname)
            if got is None:
                failures["file not indexed"].append(fname)
                # Count every expected field as wrong so unindexed files
                # cannot inflate accuracy.
                for field in ("quote_date", "customer"):
                    stats[field][1] += 1
                for it in quote["items"]:
                    for field in ("part_number", "product_name", "description",
                                  "unit_price", "quantity"):
                        stats[field][1] += 1
                continue

            score("quote_date", got["quote_date"] == quote["quote_date"],
                  "date mismatch",
                  f"{fname}: expected {quote['quote_date']}, got {got['quote_date']}")
            score("customer", text_match(quote["customer"], got["customer"]),
                  "customer mismatch",
                  f"{fname}: expected {quote['customer']!r}, got {got['customer']!r}")

            # Align items by part number.
            got_by_pn = {}
            for it in got["items"]:
                pn = (it["part_number"] or "").upper()
                got_by_pn.setdefault(pn, []).append(it)

            matched_pns = set()
            for exp in quote["items"]:
                pn = exp["part_number"].upper()
                cands = got_by_pn.get(pn, [])
                if not cands:
                    score("part_number", False, "line item missing",
                          f"{fname}: item {pn} not extracted")
                    for field in ("product_name", "description",
                                  "unit_price", "quantity"):
                        stats[field][1] += 1
                    continue
                matched_pns.add(pn)
                it = cands[0]
                score("part_number", True, "", "")
                score("product_name",
                      text_match(exp["product_name"], it["product_name"]),
                      "product name mismatch",
                      f"{fname} {pn}: expected {exp['product_name']!r}, "
                      f"got {it['product_name']!r}")
                score("description",
                      text_match(exp["description"], it["description"]),
                      "description mismatch",
                      f"{fname} {pn}: expected {str(exp['description'])[:80]!r}..., "
                      f"got {str(it['description'])[:80]!r}...")
                score("unit_price",
                      price_match(exp["unit_price"], it["unit_price"]),
                      "price mismatch",
                      f"{fname} {pn}: expected {exp['unit_price']}, "
                      f"got {it['unit_price']}")
                score("quantity", exp["quantity"] == it["quantity"],
                      "quantity mismatch",
                      f"{fname} {pn}: expected {exp['quantity']}, "
                      f"got {it['quantity']}")

            # Spurious extracted items = false positives; they count against
            # part_number accuracy.
            expected_pns = {it["part_number"].upper() for it in quote["items"]}
            for pn, its in got_by_pn.items():
                extra = len(its) - (1 if pn in expected_pns else 0)
                for _ in range(max(extra, 0)):
                    score("part_number", False, "spurious line item",
                          f"{fname}: extracted item {pn!r} not in source document")

    return stats, failures


def main():
    if not ANSWER_KEY.exists():
        raise SystemExit("No answer key. Run generate_archive.py first.")
    con = db_connect()
    key = json.loads(ANSWER_KEY.read_text(encoding="utf-8"))
    db_quotes = load_db_quotes(con)

    n_expected_files = sum(len(q["files"]) for q in key["quotes"])
    print(f"Answer key: {len(key['quotes'])} quotes across {n_expected_files} files; "
          f"database holds {len(db_quotes)} indexed files.\n")

    stats, failures = validate(db_quotes, key)

    print(f"{'Field':<14} {'Correct':>8} {'Total':>6} {'Accuracy':>9}")
    print("-" * 41)
    grand_ok = grand_total = 0
    for field in ("part_number", "product_name", "description", "unit_price",
                  "quantity", "quote_date", "customer"):
        ok, total = stats[field]
        grand_ok += ok
        grand_total += total
        pct = ok / total * 100 if total else 0
        print(f"{field:<14} {ok:>8} {total:>6} {pct:>8.1f}%")
    overall = grand_ok / grand_total if grand_total else 0
    print("-" * 41)
    print(f"{'OVERALL':<14} {grand_ok:>8} {grand_total:>6} {overall * 100:>8.1f}%")

    if failures:
        print("\nFailures grouped by cause:")
        for cause, msgs in sorted(failures.items(), key=lambda kv: -len(kv[1])):
            if not cause:
                continue
            print(f"\n  [{cause}] x{len(msgs)}")
            for m in msgs[:10]:
                print(f"    - {m}")
            if len(msgs) > 10:
                print(f"    ... and {len(msgs) - 10} more")

    print()
    if overall >= THRESHOLD:
        print(f"PASS: overall accuracy {overall * 100:.1f}% >= {THRESHOLD * 100:.0f}%")
        return 0
    print(f"FAIL: overall accuracy {overall * 100:.1f}% < {THRESHOLD * 100:.0f}% "
          f"— improve rules.md or extraction logic and re-run indexing.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
