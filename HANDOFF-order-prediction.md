# Session Handoff — Order Prediction tab

**Date:** 2026-06-16
**Branch:** `claude/determined-darwin-uq1m82`
**File:** everything lives in the single-file dashboard `index.html`

---

## 1. What we changed on this tab THIS session

### Added an "Order by" column (merged — PR #29)
- New rightmost column **`Order by`** in **both** the Cafe and Merch prediction tables.
- Shows the recommended *next order date* per product, computed by `buildNextOrder()`.
- Rendering: a per-row `orderByCell` that displays:
  - `—` (grey) when there's no usable prediction (`nextOrderLabel` is empty/`—`)
  - **`Now`** (red, bold) when the item is already due
  - otherwise the date label, e.g. `24 Jun`
- Header colspans were bumped to fit the new column (Cafe loading row `colspan=10`, Merch `colspan=11`).

That was the only Order-Prediction-specific change this session. (The rest of the session was Weekly Report layout — PRs #26/#28/#30/#31 — and a Stock Reconciliation bootstrap fix — PR #27. None of those touch this tab.)

### Explicitly deferred this session
- **Brand grouping / brand logic** in Order Prediction — user said "leave it for now." Still open.

---

## 2. Why there are TWO tables (the open question)

The Order Prediction tab is **one tab with two sub-tabs**: **Cafe** and **Merch**. They are *not* shown at the same time — clicking the sub-tab toggles which `.sub-section` is visible.

HTML (`index.html` ~line 920–1023):
- `<div class="section" data-tab="predict">`
  - `.sub-tabs`: `data-sub="pred-cafe"` / `data-sub="pred-merch"` (line ~923)
  - `.sub-section[data-sub="pred-cafe"]` → table `#predTable` / tbody `#predTbody` (line ~928)
  - `.sub-section[data-sub="pred-merch"]` → table `#predMerchTable` / tbody `#predMerchTbody` (line ~982)

They are separate tables because **Cafe and Merch carry different columns and use different stock models**:

| | Cafe table (`#predTable`) | Merch table (`#predMerchTable`) |
|---|---|---|
| Identity columns | Product, **Brand**, **Category** | Product, **Collection**, **Colourway**, **Size** |
| Filters | category / brand / flag / order-cycle | collection / flag |
| Order horizon | order cycle selectable (1–4 wks, default 2) | fixed 8-week horizon |
| Supplier lead time | per `orderModel(false)` | 14-day lead (`orderModel(true)`) |
| Order link | Muscle Finesse | none |
| Velocity source | sales rows where `category !== 'Merch'` | merch rows |

Render path:
- `renderOrderPrediction()` (line ~3364) checks the active sub-tab; if it's `pred-merch` it delegates to `renderMerchOrderPrediction()` (line ~3564), otherwise it renders the Cafe table itself.
- Sub-tab switching wires re-render: Cafe filters call `renderOrderPrediction`, Merch filters call `renderMerchOrderPrediction` (line ~5355–5378).

**So "why 2 tables" = Cafe vs Merch have genuinely different shapes/logic.** If the next session wants a single unified table, that's a real refactor (shared columns + conditional cells), not just a CSS change. Worth confirming with the user whether they want them merged or just made clearer/labelled.

---

## 3. Key code locations (Order Prediction)

| What | Location |
|---|---|
| Tab button | `index.html:680` |
| Tab HTML (both sub-tables) | `index.html:920–1023` |
| Cafe render | `renderOrderPrediction()` ~`3364` |
| Merch render | `renderMerchOrderPrediction()` ~`3564` |
| Next-order-date calc | `buildNextOrder()` ~`3336` |
| Stock flag classifier | `classifyStock()` ~`1974` |
| Shopping list (above each table) | `renderShoppingList()` ~`3344` |
| Deep-link from Overview / Weekly | `goToOrderPrediction()` ~`2075` |
| Order model (horizon + lead days) | `orderModel()` — search for definition |
| Exclusions list | "Order Prediction exclusions" ~`1424` |

`buildNextOrder()` logic: order is due when remaining cover falls to the supplier lead time → `days = max(0, coversDays - leadDays)`; label is `Now` if `days<=0` else a `d MMM` date.

---

## 4. Things flagged as unclear / to change (for next session)

These are the items the user wants to revisit — **confirm specifics with the user before implementing**:

1. **"Why are there 2 tables?"** — decide: keep Cafe/Merch as separate sub-tab tables (current), make the split clearer (better labelling/headers), or merge into one table. See section 2.
2. **Brand grouping / brand logic** — deferred from this session. Original idea was grouping rows by brand in the Cafe table.
3. Anything else the user raises about clarity of the columns (Covers, Predicted Order, Est. Stock, Order by, Last Delivery) — the column meanings are explained in the `.pred-note` "How this works" blurb at the top of each sub-section.

---

## 5. Unrelated loose end (not this tab)

- **Supabase insert still pending:** add "Barebell Bars - Cinnamon Bun" to `shop_product_lookup` (project `ljjwssicvvyyueyznmou`). Blocked all session on the `mcp__Supabase__execute_sql` approval prompt. Fields: `product_name_raw='Barebell Bars - Cinnamon Bun'`, `display_name='Barebell Bar - Cinnamon Bun'`, `category='Bars & Snacks'`, `brand='Barebell'`, `active=true`, `stock_tracked=true`.
