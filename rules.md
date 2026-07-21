# Extraction Rules — Meridian Quote Archive

This file is loaded into the extraction prompt verbatim. Edit it in plain
English to change how quotes are read — no code changes needed. Re-run
`python index_archive.py --force` after editing so already-indexed files are
re-extracted under the new rules.

## Field synonyms

Different quotes use different labels for the same field. Treat all of these
as the SAME field:

- **Part number**: "Part No.", "PN", "Part #", "Item Number", "Part Number",
  "Item #". Part numbers look like `ABC-1234` (2-4 uppercase letters, hyphen,
  digits). Always normalize to uppercase with the hyphen kept.
  <!-- Add your own patterns here if your suppliers use other formats. -->
- **Quantity**: "Qty", "Quantity", "Qty.", "Order Qty".
- **Unit price**: "Unit Price", "Price", "Unit Cost", "Price Each",
  "Price/Unit". This is the per-unit price, never a line total.
- **Description**: "Description", "Product Description", "Details",
  "Specification". The product NAME (short title) is separate from the
  description (the 2-4 sentence technical text).

## Layouts you will see

1. **Table-based** — one row per line item under a header row. PDF table rows
   may wrap: a long product name or description can spill onto following
   lines while other columns stay on the first line. Reassemble wrapped cells
   into one value.
2. **Paragraph-based** — each item is a block: "Line N: <name>", then the
   part number line, then description, then qty/price line.
3. **Letter-style** — items appear inside prose bullets like
   "Product Name (PN ABC-1234), quantity 3 — $99.00 per unit." with the
   description in the following sentence(s).

## Prices

Formats vary: `$1,204.50`, `1204.5 USD`, `USD 1,204.50`, `$1204.50 /ea`,
bare `1,204.50`. Always output the plain decimal number (e.g. `1204.5`) —
strip currency symbols, thousands separators, and suffixes like `/ea` or
`per unit`. Assume USD; do not convert currencies.

## Dates

Formats vary: `March 14, 2023`, `03/14/2023`, `2023-03-14`, `14-Mar-2023`.
Always output ISO format `YYYY-MM-DD`. Dates in `MM/DD/YYYY` are US style
(month first). <!-- If your archive has DD/MM dates, say so here. -->

## Missing data

- If a field is genuinely absent from the document, output `null`. NEVER
  guess or invent a value.
- A description is only the technical text about that product — do not
  substitute boilerplate (terms, validity, greetings) for a missing
  description.
- If the customer name is missing, output `null` for customer.

## What is NOT a line item

- Address lines, phone numbers, terms ("Net 30", "FOB Toledo"), validity
  statements, greetings and signatures are never line items.
- The quote number (e.g. `Q-2023-0147`, `MIS-23-0147`, `QT230147`) is not a
  part number, even though it can look similar.

## Tie-breaking (used at quote-generation time)

When the same part number appears in multiple archived quotes, the **most
recent quote date wins** — its price, name, and description are what go into
a new quote. <!-- Change this to e.g. "highest price wins" if you prefer. -->
