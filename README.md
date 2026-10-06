# Vibe Sheet — MVP v1.0 (Excel add-in)

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

## Setup (already set for github.com/quickshortscompany-hub)

### 1. Upload to GitHub (replace everything)
1. Open your repo **quickshortscompany-hub/vibesheet**.
2. **Add file → Upload files** → drag in EVERYTHING from this folder: `taskpane.html`, `taskpane.js`, `taskpane.css`, `manifest.xml`, `README.md` and the `assets` folder → **Commit changes**. Same-name files are replaced.
3. Wait 2 minutes. Open https://quickshortscompany-hub.github.io/vibesheet/taskpane.html — the bottom of **Settings** must say **v1.0.2**.

### 2. Excel desktop (Windows)
Excel desktop still needs internet: the panel loads from GitHub and talks to the AI.
1. Make a folder, e.g. `C:\VibeSheet`, and copy `manifest.xml` into it.
2. Right-click the folder → **Properties → Sharing → Share…** → add yourself → **Share**. Copy the **network path** shown (like `\\YOUR-PC\VibeSheet`).
3. Excel → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
4. Paste the network path in **Catalog Url** → **Add catalog** → tick **Show in Menu** → **OK** → **OK**.
5. Close and reopen Excel.
6. **Home → Add-ins → More Add-ins** (or **Insert → My Add-ins**) → **SHARED FOLDER** tab → **Vibe Sheet** → **Add**.
7. Click **Vibe Sheet** on the Home tab.

### 3. Excel on the web (alternative)
**Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** → choose `manifest.xml`. If an older Vibe Sheet is there, remove it first (… → Remove).

### 4. First run
**Settings** → paste your Anthropic key → **Load** → pick a Sonnet model → **Save**.
If Load shows an error, the red box now shows Anthropic's exact message. You can also type a model ID in the box under the dropdown and press Save.

## Tips
- **Ctrl/⌘ + Enter** sends.
- Select a single empty cell and ask for a table: it fills empty cells down/right.
- Changes outside your selection that would overwrite data are **unticked by default** (yellow badge).
- Undo works in the panel (newest first). Excel's own Ctrl+Z may not undo add-in changes.

## Limits of this MVP
- Your API key is stored in this browser only. Fine for you and testers using their own keys. **Before a public launch, move the key to a small server** (v2) so users don't need their own key and you can charge for it.
- Very large ranges (25,000+ filled cells) apply but can't be undone.
- Charts, pivots and dropdowns undo by deleting them. Previous dropdown rules on those cells aren't restored.
- Needs Excel 2021 / Microsoft 365 / Excel on the web (ExcelApi 1.9).

## Files
- `manifest.xml`: tells Excel where the add-in lives and adds the ribbon button
- `taskpane.html / .css / .js`: the panel, AI call, preview, apply and undo engine
- `assets/`: icons
