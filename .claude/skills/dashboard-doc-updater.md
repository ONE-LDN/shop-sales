# dashboard-doc-updater

You are updating the Notion system document for the **Shop & Cafe Sales Dashboard** (`shop-sales` repo).

**Target Notion page:** https://www.notion.so/3680164f7fc681f09594f7f7eb4cf34b
**Page ID:** `3680164f-7fc6-81f0-9594-f7f7eb4cf34b`

## What this skill does

Reads the current `index.html`, compares its state against the Notion system document, and updates any sections that are out of date. Run after any meaningful code change to keep documentation in sync.

## Steps

1. **Read the source file**
   - Read `index.html` from the repository root
   - Because the file is large (~195KB), extract key sections with `grep` rather than reading the whole file at once:
     ```bash
     # Extract named functions
     grep -n 'function [a-zA-Z]\+(' index.html

     # Extract Supabase table references
     grep -n "shop_sales\|shop_product_lookup\|shop_deliveries\|shop_upload_log\|shop_invoice_aliases" index.html | grep "sbGet\|sbUpsert\|sbInsert\|sbDelete\|sbPatch" | head -30

     # Extract category and flag thresholds
     grep -n 'order-now\|order-soon\|coversDays\|avgPerWeek\|packSize\|CATEGORIES\|CAT_COLORS' index.html | head -40

     # Extract CDN script tags
     grep -n '<script src' index.html

     # Extract canvas elements (charts)
     grep -n '<canvas id' index.html

     # Extract period filter constants
     grep -n 'rollingWeeks\|periodMode\|monthFilter\|summarySegment' index.html | head -20
     ```

2. **Identify what has changed** compared to the documented state. Focus on:
   - New or renamed JavaScript functions
   - Changes to order flag thresholds (cafe: 7/14 days, merch: 14/56 days)
   - Changes to order cycle lengths (cafe: 2-week, merch: 8-week)
   - New categories in `CATEGORIES` or `CAT_COLORS`
   - New tabs, sub-tabs, or modals
   - New or changed Supabase tables or columns
   - Changes to fuzzy match threshold (currently 0.35)
   - New CDN dependencies or version bumps
   - New or changed known limitations
   - Changes to `EXACT_CAT`, `EXACT_PRICE`, `BRAND_PREFIXES`, `PREFIX_CAT`, `PREFIX_PRICE` fallback maps

3. **Fetch the current Notion page** to read its content:
   Use `notion-fetch` with id `3680164f-7fc6-81f0-9594-f7f7eb4cf34b`

4. **Update changed sections** using `notion-update-page`.
   - Only update sections that have actually changed
   - Preserve the page structure (13 numbered sections)
   - Use the same Notion-flavored Markdown format as the existing page
   - Always update the `> Last updated:` date at the top to today's date

5. **Report what changed** — list each section updated and a one-line summary of what changed.

## Section map (for targeted updates)

| Section | What to check in index.html |
| --- | --- |
| 1. Purpose | No change expected unless segment strategy changes |
| 2. Tech Stack | CDN library versions in `<script src=...>` tags |
| 3. Architecture | `init()`, `loadFromSupabase()`, `handleFiles()`, `handleInvoicePdfs()` flow |
| 4. Database Schema | Supabase column references — check `sbGet`/`sbUpsert` calls |
| 5. Business Logic | Flag thresholds, velocity window (4 weeks), order cycles (2/8 weeks), GP% formula, fuzzy match threshold |
| 6. Dashboard Views | Tab panel structure, sub-tab IDs |
| 7. Modals | Modal IDs and their step structure |
| 8. Key Functions | All `function` declarations |
| 9. Charts | All `<canvas id=...>` elements and `new Chart(...)` calls |
| 10. State Management | Global `let`/`var`/`const` declarations at top of `<script>` |
| 11. UI Colour System | CSS custom properties or colour constants |
| 12. Known Limitations | Update when known issues change |
| 13. Maintenance Guide | Update when workflows change |

## Notes

- Do **not** update the page if nothing has changed
- If a section is completely rewritten, replace it entirely rather than patching
- The `shop_invoice_aliases` table builds up automatically through PDF imports — no manual intervention needed
- When new suppliers are added (beyond Muscle Finesse and Tapstitch), update the supplier detection section in Section 8 and the Log Delivery modal docs in Section 7
- Amazon CSV parsing (`parseAmazonCSV`) is currently a stub — update Section 12 Known Limitations if this is implemented
