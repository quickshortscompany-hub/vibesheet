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

## Setup (about 15 minutes, once)

### 1. Get an AI key
1. Go to **console.anthropic.com**, sign up, and add a small amount of credit (e.g. $5 USD).
2. **API Keys → Create key** and copy it (starts with `sk-ant-`).

### 2. Put the add-in online (free, GitHub Pages)
1. Create a free account at **github.com**.
2. **New repository** → name it exactly `vibesheet` → Public → Create.
3. Click **uploading an existing file** and drag in everything from this folder (`taskpane.html`, `taskpane.js`, `taskpane.css`, `manifest.xml`, `README.md`, and the `assets` folder) → **Commit changes**.
4. Repo **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main`, folder `/ (root)` → Save.
5. Wait about a minute. Check that `https://YOUR-USERNAME.github.io/vibesheet/taskpane.html` opens.

### 3. Point the manifest at your site
Open `manifest.xml` in Notepad / TextEdit → **Find & Replace** `YOUR-GITHUB-USERNAME` with your GitHub username (lowercase) → Save.
(Keep this edited copy on your computer; it's the file you upload in step 4.)

### 4. Load it into Excel
**Excel on the web (easiest, works on any computer):**
excel.cloud.microsoft → open a workbook → **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** → choose `manifest.xml`.

**Windows desktop:** put `manifest.xml` in a folder, share it (right-click → Properties → Sharing → Share), copy the network path (e.g. `\\YOURPC\addins`).
Excel → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → paste the path → Add → tick *Show in Menu* → OK → restart Excel → **Home → Add-ins → Advanced → Shared Folder** → Vibe Sheet.

**Mac desktop:** copy `manifest.xml` into
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create `wef` if missing) → restart Excel → **Home → Add-ins → My Add-ins**.

> If you use a uni / work Microsoft account and uploading is blocked, use a personal Microsoft account on Excel for the web instead.

### 5. First run
**Home → Vibe Sheet** → **Settings** tab → paste key → **Load** (choose a model; Sonnet is a good default) → **Save**.

---

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
