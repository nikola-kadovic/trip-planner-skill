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
- **Choose-an-option:** `tx.choose_validation(ws, r1, r2, c)` (dropdown on a segment's Choose column) · `tx.choose_mark(ws, r, c)` (write the marker into the recommended row so a default is pre-selected) · `tx.chosen_total(value_range, marker_range, sheet="Hotels")` (SUMIF that totals whichever row is marked) · `tx.chosen_label(...)` (INDEX/MATCH the marked row's name). The marker glyph is `tx.CHOSEN_MARK`.
- **Per-person:** `tx.per_person(total_cell, travelers_cell)` → `'=F12/$C$4'`, for the per-person column.
- Currency: `tx.money("CAD")` → a ready number format; extend `tx._CURRENCY_SYMBOL` for exotic currencies.
- Close-out: `tx.finalize(wb, out_path, link_checks=[…])` — see Final checks below.

Palette/format constants live in the module and follow the semantics in the Palette section above (title=DARK, sections=MID, zebra=LIGHT, chosen=CHOSEN; inputs blue, formulas black, cross-sheet refs green, hyperlinks blue-underlined). Font is Arial throughout.

**The hyperlink rule still holds, the library just enforces it:** `tx.link()` sets `cell.hyperlink` for you and refuses an empty URL, so a priced row literally can't ship link-less if you build it with `tx.link`. Put the link on a cell that reads naturally (hotel name, an airline/"Book" cell, the activity's "Book/Source" cell).

## Build order

Build **Hotels** and **Flights** first. For each segment, record the **range** of option rows — the Total (home) column range and the Choose-marker column range — not a single chosen cell. Then build **Trip Summary** so its cost lines reference the *marked* option via `tx.chosen_total(value_range, marker_range, sheet=…)` (a SUMIF), e.g. `=SUMIF('Hotels'!$K$4:$K$8,"✓",'Hotels'!$I$4:$I$8)` (mark column K, sum the Total-home column I). This is what makes the pick switchable: the user moves the marker, and the Summary total follows on recalc — no formula editing. Put the Summary tab at index 0 so it opens first.

## Choosing options (switchable picks)

Every segment that offers alternatives (each city's hotels, each flight leg) gets a **Choose** column as its last data column. Build it the same way everywhere:

1. Add a dropdown to the segment's marker cells with `tx.choose_validation(ws, r1, r2, choose_col)` — the user picks from the marker glyph (`tx.CHOSEN_MARK`, `✓`) or blank.
2. Pre-select the recommended option: `tx.choose_mark(ws, rec_row, choose_col)` and `tx.highlight_row(ws, rec_row, c1, c2)` so the workbook opens with a valid default and the recommendation is visible.
3. On Trip Summary, pull that segment's chosen total with `tx.chosen_total(total_range, marker_range, sheet=…)`. Optionally echo the chosen name into the Details column with `tx.chosen_label(name_range, marker_range, sheet=…)` so the Summary relabels itself when the pick changes.

Mark exactly one row per segment (note this in a small italic line under the section). Because the lookup is a SUMIF over the segment's own rows, picks in different cities/legs are independent.

## Trip Summary tab

1. Title band + subtitle (trip name, nights, travelers, dates, budget tier).
2. **Assumptions** block — blue input cells: Travelers, Total nights, nights per city, `EUR→home` FX, `USD→home` FX. Everything downstream references these.
3. **Cost Breakdown** table — columns: Category · Details · Qty · Unit (home cur) · **Per person** · Group total · Notes/source. The **Per person column comes before the Group total and is the emphasized one** (bold / ACCENT), because "what does one flight / one room-share cost me" is the figure people scan for; the group total sits to its right for the household sum. Build each line as:
   - Group total = `Qty*Unit` as a formula.
   - Per person = `tx.per_person(group_total_cell, travelers_cell)` (i.e. `=GroupTotal/$C$<travelers>`). For inherently per-person lines (a flight seat, an activity priced per head) this resolves to the single-unit price, which is exactly the point.
   - Flights → Unit/Group total = `tx.chosen_total(...)` cross-sheet SUMIF over the Flights options (green), so it tracks the chosen flight.
   - Hotel per city → same, a `tx.chosen_total(...)` SUMIF over that city's Hotels options (green).
   - Estimated lines (train, food, local transport, activities, city tax, ETIAS, insurance) → Unit as a formula off the FX cell, e.g. `=60*$C$10`, with the source/basis in Notes.
   - Food qty = `travelers * nights`; per-person items qty = `travelers`.
4. **Subtotal** `=SUM(...)` (group) with its per-person counterpart, **Contingency buffer** (e.g. `=Subtotal*0.05`), **Estimated Total**, and a prominent **Per person total** `=Total/travelers`.

## Hotels tab

One section per city. Columns: Property (hyperlinked name) · Platform · Area/District · Stars · Score · Reviews · Total stay · Cur · Total (home cur) · **Per person** · **Choose**. Rules:
- **Every property row's name cell must be a working hyperlink** to that option's booking page (`cell.hyperlink = url`). No row ships with a price but no link.
- Total (home cur): if `Cur == home` → `=H{r}`; else convert via FX cell, e.g. `=H{r}*'Trip Summary'!$C$11`.
- **Per person** = `tx.per_person(total_home_cell, travelers_cell)` — the room/stay cost split across the party, so a single traveler's share reads at a glance.
- Show options from **all four platforms** (Booking.com, Expedia, Tripadvisor, lastminute.com); the Platform column must visibly include each source that returned results.
- **Choose column:** add a dropdown over the city's option rows with `tx.choose_validation(...)`, pre-mark the recommended row with `tx.choose_mark(...)` and `tx.highlight_row(...)` (CHOSEN fill). The Trip Summary's hotel line for this city pulls the marked row's total via `tx.chosen_total(total_range, marker_range, sheet="Hotels")`, so the user can switch hotels and the trip total updates. Record the two ranges (totals, markers) for the Summary lookup.
- Header note: the date the prices were pulled + "live prices change, click a name to book."
- **Platform coverage row:** under each city section, add a small italic line listing which platforms were queried and how many results each returned (from `platform_coverage`), e.g. *"Searched: Booking.com (8), Expedia (6), Tripadvisor (5), lastminute.com (4)."* If any of the four returned nothing or errored, say so explicitly (*"lastminute.com — no results"*) so the user can see coverage at a glance rather than wondering whether a platform was skipped.

## Flights tab

Columns: Airline · Route · Stops · Duration (out/ret) · **Fare pp (home cur)** · Total (home cur) · **Choose** · Book · Notes. Fare pp = `fare_pp * USD→home` (the price of a single seat — the figure travelers want); Total = `Fare_pp_home * travelers`. Prefer/mark nonstop. **Every flight row's "Book" cell hyperlinks to that option's own Expedia Flights-Search URL** (built per `references/connectors.md` for that exact route/dates/party) — one link per row, not a single shared header link. **Choose column:** dropdown via `tx.choose_validation(...)`, recommended row pre-marked with `tx.choose_mark(...)` + CHOSEN fill; the Summary's flight line uses `tx.chosen_total(total_range, marker_range, sheet="Flights")` so switching flights updates the trip total.

## Activities tab

Columns: Activity · Stop · **Per person** · Qty (people or units) · Total (home cur) · Source/Book · Notes. The **Per person column is the prominent one** — it's the single-ticket price. **Every row's Source/Book cell must be a working hyperlink** (`cell.hyperlink = url`) to the booking page or official source — including free activities, which link their info page. Total = `per_person * qty` (convert currency via an FX cell if not in home currency). Mark estimates vs. live prices in Notes. Roll the activities total into the Trip Summary "Activities & tours" line via a cross-sheet link (green), replacing the lump estimate. (Activities are typically all-included rather than either/or, so they don't need a Choose column — but if you list alternatives for one slot, give that group a Choose column and SUMIF the same way.)

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
