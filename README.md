# Vibe Sheet v1.1 (Excel add-in)

Select cells → type what you want → preview → Accept → Undo if needed.
The AI can only change your selection (plus empty cells next to it) unless you tick **Allow edits outside selection**.

## What it does
| Problem | How you ask |
|---|---|
| Lookups / #N/A | "pull the price from Sheet2 using the product code" |
| Messy data | **Clean data** chip, or "split full name into first and last" |
| Dates | "convert these to dd/mm/yyyy", "days between B and C" |
| IF / SUMIFS logic | "total sales for Melbourne in March" |
| Errors | **Fix / explain errors** chip |
| Can't find a ribbon button | "make negatives red", "add a Yes/No dropdown", "freeze top row" |
| Pivot tables | **Pivot table** chip, or "sum of sales by region and month" |
| Charts | **Chart it** chip, or "line chart of revenue with title" |
| Repetitive work | "highlight duplicates", "sort by date newest first", "remove duplicates" |
| $A$1 copy problems | AI writes each row's formula with the right references |
| **Named cells** | "call this Span", or the **Name cells** chip, then formulas read `=Load_w*Span^2/8` |
| **Images** | paste / drop / 📷 a photo: table → cells, formula → working Excel formula with labelled inputs |

Every change shows a before/after preview. You can untick individual changes. Prompts are saved with the cells — select a cell to see which prompt made it.

---

## Setup
Open **SETUP-GUIDE.html** (double-click it). It covers the secure server (Cloudflare Worker that holds your Anthropic key), connecting the add-in, sharing it, and reading feedback on **admin.html**.

## Files
- `taskpane.html / .css / .js` the add-in panel (Prompt, History, Help, Feedback, Settings)
- `config.js` the only file you edit: your Worker address
- `manifest.xml` tells Excel where the add-in lives
- `admin.html` read feedback and usage (needs your ADMIN_TOKEN)
- `server/worker.js` paste into Cloudflare; do NOT upload to GitHub (it holds no secrets, but isn't needed there)

## Tips
- **Ctrl/⌘ + Enter** sends.
- Select a single empty cell and ask for a table: it fills empty cells down/right.
- Changes outside your selection that would overwrite data are **unticked by default** (yellow badge).
- Undo works in the panel (newest first). Excel's own Ctrl+Z may not undo add-in changes.

## Limits
- The AI key lives only in Cloudflare. Each person gets a daily prompt limit; set a monthly spend limit in Anthropic as the final safety net.
- Very large ranges (25,000+ filled cells) apply but can't be undone.
- Charts, pivots and dropdowns undo by deleting them. Previous dropdown rules on those cells aren't restored.
- Needs Excel 2021 / Microsoft 365 / Excel on the web (ExcelApi 1.9).
