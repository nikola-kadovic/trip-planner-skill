# Spreadsheet style & build guide

Read the `xlsx` skill first. Build with `openpyxl` via the bundled **`scripts/trip_xlsx.py`** library (below), which handles styling, links, the link audit, and the recalc step. Font: **Arial** throughout. `tx.finalize()` runs the `xlsx` skill's `recalc.py` for you and reports any formula errors — fix until zero.

## Palette & formats

The palette and currency formats live in `scripts/trip_xlsx.py` as overridable constants — you don't redeclare them. For reference, the defaults are: `DARK` title bands, `MID` section/header bands, `LIGHT` zebra stripes, `ACCENT` emphasis, `CHOSEN` for the recommended row; text colors `C_INPUT` (blue inputs), `C_XREF` (green cross-sheet refs), `C_LINK` (blue-underline hyperlinks), `C_NOTE` (gray italic). Currency formats come from `tx.money("CAD")` etc. (positive; negative-in-parens; zero-as-dash). Override any constant after import if a trip wants a different look.

## Use the bundled builder — don't rewrite the boilerplate

The styling, hyperlink mechanics, currency formats, link audit, and save/recalc are
already written and tested in **`scripts/trip_xlsx.py`**. Import it instead of
redefining a `sc()` helper and re-deriving fills/borders every run — that rewrite is
slow and is exactly where the "blue text that isn't actually clickable" bug comes
from. Your build script should contain only the **trip-specific** part: which tabs
this trip needs, their columns, and the cost formulas. Compose freely — the library
doesn't dictate tabs (a ski trip has a ski-pass tab; a beach trip won't).

```python
import sys; sys.path.insert(0, "<skill-dir>/scripts")
import trip_xlsx as tx
```

API (all cell helpers return the cell so you can tweak further):
- `wb = tx.new_workbook()`
- `ws = tx.sheet(wb, "Hotels", first=True, widths={1:26, 2:12})` — `first=True` for the first tab so no blank "Sheet" is left behind.
- Bands/headers: `tx.title(ws, r, text, ncols=…)`, `tx.band(ws, r, text, ncols=…)`, `tx.header(ws, r, [col names…])`.
- Cells: `tx.lbl` (text) · `tx.num` (blue editable input, e.g. `fmt=tx.money("EUR")`) · `tx.formula` (black computed) · `tx.xref` (green cross-sheet ref) · `tx.link(ws, r, c, text, url)` (sets `cell.hyperlink` **and** styles it — raises if `url` is empty) · `tx.note` (gray italic, supports `colspan`).
- Layout: `tx.fill_row`, `tx.zebra(ws, r1, r2, c1, c2)` (preserves any existing fill, so it won't clobber a chosen-row highlight), `tx.highlight_row(ws, r, c1, c2)` (CHOSEN fill), `tx.borders(ws, r1, c1, r2, c2)`.
- Currency: `tx.money("CAD")` → a ready number format; extend `tx._CURRENCY_SYMBOL` for exotic currencies.
- Close-out: `tx.finalize(wb, out_path, link_checks=[…])` — see Final checks below.

Palette/format constants live in the module and follow the semantics in the Palette section above (title=DARK, sections=MID, zebra=LIGHT, chosen=CHOSEN; inputs blue, formulas black, cross-sheet refs green, hyperlinks blue-underlined). Font is Arial throughout.

**The hyperlink rule still holds, the library just enforces it:** `tx.link()` sets `cell.hyperlink` for you and refuses an empty URL, so a priced row literally can't ship link-less if you build it with `tx.link`. Put the link on a cell that reads naturally (hotel name, an airline/"Book" cell, the activity's "Book/Source" cell).

## Build order

Build **Hotels** and **Flights** first and capture the cell address of each *chosen* row's Total (in home currency). Then build **Trip Summary** so its cost lines can cross-reference those cells (`='Hotels'!J7`, `='Flights'!G5`). Put the Summary tab at index 0 so it opens first.

## Trip Summary tab

1. Title band + subtitle (trip name, nights, travelers, dates, budget tier).
2. **Assumptions** block — blue input cells: Travelers, Total nights, nights per city, `EUR→home` FX, `USD→home` FX. Everything downstream references these.
3. **Cost Breakdown** table — columns: Category · Details · Qty · Unit (home cur) · Subtotal · Notes/source. Each `Subtotal = Qty*Unit` as a formula. Lines:
   - Flights → Unit = cross-sheet link to chosen flight total (green).
   - Hotel per city → Unit = cross-sheet link to chosen hotel total (green).
   - Estimated lines (train, food, local transport, activities, city tax, ETIAS, insurance) → Unit as a formula off the FX cell, e.g. `=60*$C$10`, with the source/basis in Notes.
   - Food qty = `travelers * nights`; per-person items qty = `travelers`.
4. **Subtotal** `=SUM(...)`, **Contingency buffer** (e.g. `=Subtotal*0.05`), **Estimated Total**, **Per person** `=Total/travelers`.

## Hotels tab

One section per city. Columns: Property (hyperlinked name) · Platform · Area/District · Stars · Score · Reviews · Total stay · Cur · Total (home cur) · Chosen · (spacer). Rules:
- **Every property row's name cell must be a working hyperlink** to that option's booking page (`cell.hyperlink = url`). No row ships with a price but no link.
- Total (home cur): if `Cur == home` → `=H{r}`; else convert via FX cell, e.g. `=H{r}*'Trip Summary'!$C$11`.
- Show options from **all four platforms** (Booking.com, Expedia, Tripadvisor, lastminute.com); the Platform column must visibly include each source that returned results. Highlight the chosen row with CHOSEN fill and a ★; record its Total cell for the Summary link.
- Header note: the date the prices were pulled + "live prices change, click a name to book."
- **Platform coverage row:** under each city section, add a small italic line listing which platforms were queried and how many results each returned (from `platform_coverage`), e.g. *"Searched: Booking.com (8), Expedia (6), Tripadvisor (5), lastminute.com (4)."* If any of the four returned nothing or errored, say so explicitly (*"lastminute.com — no results"*) so the user can see coverage at a glance rather than wondering whether a platform was skipped.

## Flights tab

Columns: Airline · Route · Stops · Duration (out/ret) · Fare pp · Total (home cur) · **Book** · Notes. Total = `fare_pp * travelers * USD→home` (Expedia fares are USD). Prefer/mark nonstop. **Every flight row's "Book" cell hyperlinks to that option's own Expedia Flights-Search URL** (built per `references/connectors.md` for that exact route/dates/party) — one link per row, not a single shared header link. The chosen flight row gets the CHOSEN fill and ★, and its Total cell is the one the Summary cross-references.

## Activities tab

Columns: Activity · Stop · Per person · Qty (people or units) · Total (home cur) · Source/Book · Notes. **Every row's Source/Book cell must be a working hyperlink** (`cell.hyperlink = url`) to the booking page or official source — including free activities, which link their info page. Total = `per_person * qty` (convert currency via an FX cell if not in home currency). Mark estimates vs. live prices in Notes. Roll the activities total into the Trip Summary "Activities & tours" line via a cross-sheet link (green), replacing the lump estimate.

## Itinerary tab (when useful)

Columns: Day · City · Plan · Notes. One row per day; keep arrival days light; flag intercity travel days and any timed/bookable items.

## Final checks

Close out with one call — `tx.finalize()` saves the file (creating the outputs dir if needed), runs the link-coverage audit, and runs `recalc.py`:

```python
report = tx.finalize(
    wb,
    "/mnt/user-data/outputs/<Trip_Name>.xlsx",
    link_checks=[("Hotels", "I"), ("Flights", "F"), ("Activities", "E")],
)
# link_checks = (sheet, price-column letter) for the BOOKABLE tabs only — never the
# Trip Summary, whose per-person formula cells are priced but shouldn't link.
# A row counts as covered if ANY cell in it has a hyperlink, so the link can sit on
# the name column or a dedicated "Book" column — whichever reads naturally.
```

Then **gate on the report — do not present a workbook that fails either check:**
- `report["missing_links"]` must be empty. If it lists rows, recover each URL from the phase output (or build a search URL for that exact hotel/flight/activity) with `tx.link(...)`, then call `tx.finalize()` again.
- `report["recalc_ok"]` must be `True` (zero formula errors). If `False`, read `report["recalc_output"]`, fix the formulas, and re-finalize. (`None` means `recalc.py` wasn't found — run the `xlsx` skill's recalc manually.)
- Spot-check that cross-sheet links resolved and that non-home-currency conversions look right.

Once both checks pass, `present_files` the path with a short summary (total, per person, chosen picks, and a note that the link audit passed).
