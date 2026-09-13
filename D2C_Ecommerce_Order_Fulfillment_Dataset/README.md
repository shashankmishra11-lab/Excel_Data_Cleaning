# D2C Order Fulfillment — Data Cleaning & Analysis

**~200 orders · 25 columns · a synthetic dataset modeled on real D2C order fulfillment data**

E-commerce order data usually isn't one clean file — it's an order management export, a payments log, and a courier feed, stitched together with three different date formats and two different measurement units. This project simulates exactly that, and works through it the way an analyst actually would: investigate before fixing, only correct what's provably safe to correct, and document the rest.

## Try it yourself
Grab `raw_dataset.xlsx` from this repo and see how many of these issues you can catch on your own before checking `cleaned_dataset.xlsx`. Good practice if you're building your own data cleaning portfolio.

## What was found & fixed

- **order_id** — Traced duplicate keys to genuine ID collisions (two different orders, one shared ID — not true duplicates). Reassigned new IDs with a collision-free formula (expanding COUNTIF + MAX offset), keeping a full audit trail of every original value.
- **order_date, dispatch_date, delivery_date** — Reconciled three columns each mixing real dates, text-formatted dates in two different regional formats (MM/DD/YYYY and DD/MM/YY), and sentinel "broken" dates (year 1900) — parsed and flagged without guessing at the unrecoverable ones.
- **weight** — Caught a unit-inconsistency bug: gram and kilogram values mixed in a single column. Fixed and verified by checking that weight-per-unit held consistent against order quantity, across every product.
- **customer_email** — Standardized formatting, and told apart corrupted entries from legitimate repeat-customer duplicates.
- **sku ↔ product_name** — Built an independent reference table and cross-validated every row with XLOOKUP to catch mismatches.
- **order_value / unit_price** — Cleaned inconsistent currency formatting (mixed ₹ symbols, "INR" prefixes, and plain numbers) into one consistent format.
- **payment_mode** — Standardized inconsistent category labels ("NetBanking" → "Net Banking", "Cash on Delivery" → "COD").
- **courier_partner / courier_ref** — Investigated mismatches before assuming error — several turned out to be legitimate subcontracted deliveries, not data issues.
- **oms_status ↔ delivery_status** — Cross-checked for logical contradictions; found and corrected an order marked *Cancelled* that also showed as *Delivered*.
- **awb** — Detected duplicate AWBs — some pointing to the same real order recorded twice under two different order_ids, likely from merged export sources.

## Flagged, not force-fixed
A handful of order_id/AWB pairs represent the same real order captured under two different IDs. Rather than guess which record to keep, these are flagged for review — some data problems don't have a clean automatable answer, and forcing one just hides the uncertainty.

## Tools
Excel — COUNTIF, XLOOKUP, IF/IFS, DATE, LEFT/MID/RIGHT/TEXT, ISNUMBER, SUBSTITUTE

## Files
- `raw_dataset.xlsx` — original messy export
- `cleaned_dataset.xlsx` — final cleaned version, audit-trail columns preserved
- `README.md` — this file

*(rename to match your actual file names before pushing)*

## Let's connect
[LinkedIn](https://in.linkedin.com/in/shashank-mishra-58678b375) · [GitHub](https://github.com/shashankmishra11-labs)

If you're also building a data analyst portfolio, feel free to fork this, try the cleaning yourself, or reach out — always happy to talk data cleaning, Excel formulas, or e-commerce analytics.
