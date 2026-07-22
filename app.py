"""Public read-only demo website for the quote generator.

Reuses make_quote's search / fuzzy-match / template-population logic — no
extraction API calls at runtime, no database writes. Generated .docx files
live only in memory (served once from a token URL, expired after 10 min).

Run locally:   python app.py          (http://127.0.0.1:5000)
Production:    gunicorn app:app       (see render.yaml)
"""

import json
import re
import secrets
import shutil
import sqlite3
import tempfile
import threading
import time
from collections import deque
from datetime import date
from pathlib import Path

from flask import (Flask, abort, redirect, render_template, request,
                   send_file, url_for)
import io

from common import ROOT, DB_PATH
import make_quote

app = Flask(__name__)

GITHUB_URL = "https://github.com/sakqo/quote-intel"

# ------------------------------------------------------------ read-only DB ---

def ro_connect():
    """Read-only SQLite connection — the demo can never write to the DB."""
    return sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)


def load_stats():
    p = ROOT / "data" / "validation_report.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return None


STATS = load_stats()


def all_parts():
    """[(part_number, latest product name, #occurrences, latest price)]"""
    con = ro_connect()
    try:
        rows = con.execute("""
            SELECT li.part_number,
                   (SELECT li2.product_name FROM line_items li2
                    JOIN quotes q2 ON q2.id = li2.quote_id
                    WHERE li2.part_number = li.part_number
                      AND li2.product_name IS NOT NULL
                    ORDER BY q2.quote_date DESC LIMIT 1),
                   COUNT(DISTINCT q.quote_number),
                   (SELECT li3.unit_price FROM line_items li3
                    JOIN quotes q3 ON q3.id = li3.quote_id
                    WHERE li3.part_number = li.part_number
                      AND li3.unit_price IS NOT NULL
                    ORDER BY q3.quote_date DESC LIMIT 1)
            FROM line_items li JOIN quotes q ON q.id = li.quote_id
            WHERE li.part_number IS NOT NULL
            GROUP BY li.part_number ORDER BY li.part_number""").fetchall()
        return rows
    finally:
        con.close()


def part_history(part):
    """Occurrences of a part grouped by quote (docx+pdf twins merged),
    newest first."""
    con = ro_connect()
    try:
        rows = con.execute("""
            SELECT q.quote_number, q.quote_date, q.customer, q.file,
                   li.product_name, li.description, li.unit_price, li.quantity
            FROM line_items li JOIN quotes q ON q.id = li.quote_id
            WHERE UPPER(li.part_number) = UPPER(?)
            ORDER BY q.quote_date DESC, q.id ASC""", (part,)).fetchall()
    finally:
        con.close()
    groups = {}
    order = []
    for qnum, qdate, cust, fname, name, desc, price, qty in rows:
        g = groups.get(qnum)
        if g is None:
            g = {"quote_number": qnum, "quote_date": qdate, "customer": cust,
                 "files": [], "product_name": name, "description": desc,
                 "unit_price": price, "quantity": qty}
            groups[qnum] = g
            order.append(g)
        g["files"].append(fname)
    return order


# --------------------------------------------------- abuse protection bits ---

RATE_LIMIT = 20          # generation requests
RATE_WINDOW = 3600       # per hour
_rate: dict[str, deque] = {}
_rate_lock = threading.Lock()


def client_ip():
    fwd = request.headers.get("X-Forwarded-For", "")
    return fwd.split(",")[0].strip() if fwd else (request.remote_addr or "?")


def rate_limited(ip):
    now = time.time()
    with _rate_lock:
        dq = _rate.setdefault(ip, deque())
        while dq and now - dq[0] > RATE_WINDOW:
            dq.popleft()
        if len(dq) >= RATE_LIMIT:
            return True
        dq.append(now)
        # Don't let the map grow without bound.
        if len(_rate) > 5000:
            for k in [k for k, v in _rate.items() if not v]:
                del _rate[k]
        return False


def sanitize(s, max_len=80):
    """Names/companies: keep word chars and light punctuation only."""
    s = re.sub(r"[^\w\s.,&'()\-/]", "", s or "", flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()[:max_len]


# ------------------------------------------- in-memory generated documents ---

DOC_TTL = 600            # seconds a generated docx stays downloadable
_docs: dict[str, tuple[bytes, str, float]] = {}
_docs_lock = threading.Lock()


def stash_doc(data: bytes, filename: str) -> str:
    token = secrets.token_urlsafe(12)
    now = time.time()
    with _docs_lock:
        for k in [k for k, (_, _, t) in _docs.items() if now - t > DOC_TTL]:
            del _docs[k]
        while len(_docs) > 100:          # hard cap on memory
            del _docs[next(iter(_docs))]
        _docs[token] = (data, filename, now)
    return token


# -------------------------------------------------------------------- pages ---

@app.route("/")
def home():
    return render_template("home.html", parts=all_parts(), stats=STATS,
                           github=GITHUB_URL)


@app.route("/search")
def search():
    q = (request.args.get("q") or "").strip()
    if not q:
        return redirect(url_for("home"))
    return redirect(url_for("part_page", part=q.upper()))


@app.route("/part/<part>")
def part_page(part):
    history = part_history(part)
    if not history:
        con = ro_connect()
        try:
            suggestions = make_quote.suggest(con, part)
        finally:
            con.close()
        return render_template("not_found.html", part=part,
                               suggestions=suggestions), 404
    return render_template("part.html", part=part.upper(), history=history,
                           today=date.today())


@app.route("/generate", methods=["POST"])
def generate():
    ip = client_ip()
    if rate_limited(ip):
        return render_template(
            "error.html",
            message="Rate limit reached (20 quote generations per hour). "
                    "Please try again later."), 429

    parts = [p.strip().upper() for p in
             (request.form.get("parts") or "").split(",") if p.strip()][:10]
    customer = sanitize(request.form.get("customer")) or "Demo Visitor"
    company = sanitize(request.form.get("company")) or "Example Manufacturing Co."
    try:
        qty = max(1, min(int(request.form.get("qty", "1")), 999))
    except ValueError:
        qty = 1
    if not parts:
        return render_template("error.html",
                               message="No part numbers given."), 400

    tmp = Path(tempfile.mkdtemp(prefix="quote-demo-"))
    try:
        out_path, found, missing = make_quote.generate(
            parts, customer, company, quantities=[qty] * len(parts),
            out_dir=tmp, allow_partial=True)
        if out_path is None:
            con = ro_connect()
            try:
                sugg = {p: make_quote.suggest(con, p) for p, _ in missing}
            finally:
                con.close()
            return render_template("not_found.html", part=", ".join(parts),
                                   suggestions=[s for v in sugg.values()
                                                for s in v]), 404
        data = out_path.read_bytes()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)   # temp file never lingers

    qnum = out_path.stem
    token = stash_doc(data, f"{qnum}.docx")
    total = sum((f["unit_price"] or 0) * qty for f in found)
    return render_template(
        "preview.html", quote_number=qnum, customer=customer, company=company,
        today=date.today(), items=found, qty=qty, total=total,
        token=token, missing=missing)


@app.route("/download/<token>")
def download(token):
    with _docs_lock:
        entry = _docs.get(token)
    if entry is None:
        abort(404)
    data, filename, _ = entry
    return send_file(io.BytesIO(data), as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument"
                              ".wordprocessingml.document")


@app.route("/how-it-works")
def how_it_works():
    return render_template("how.html", stats=STATS, github=GITHUB_URL)


if __name__ == "__main__":
    app.run(debug=False, port=5001)
