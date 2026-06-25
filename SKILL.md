---
name: trip-planner
description: Plans a trip end to end through an orchestrated, multi-phase workflow and produces a clean Excel spreadsheet whose centerpiece is a full cost breakdown, alongside live, bookable hotel and flight options and priced activities pulled from real sites for the exact dates. Use this whenever the user asks to plan a trip, vacation, getaway, or holiday; asks to find hotels, flights, or things to do for specific destinations and dates; asks when is the best time to visit somewhere; or asks how much a trip would cost. Trigger on phrases like "plan me a trip to X", "help me plan a vacation", "find me hotels and flights for...", "what should I do in Y", "when should I go to Z", "what would a week in Y cost", even when the user never mentions a spreadsheet. Always trigger for multi-stop or multi-city trips.
---

# Trip Planner

An **orchestrator** runs the plan through distinct **phases** and delegates each phase to a **subagent**, then assembles everything into a polished, formula-driven Excel workbook. The deliverable always contains a **cost breakdown**, **live hotel options**, **live flight options**, **priced activities**, and (when useful) a day-by-day **itinerary**.

**The core promise: every priced item is independently clickable.** Each hotel, each flight option, and each activity gets its **own** working booking/source link, right there on its row — so the user opens the file and books anything with a single click. A price with no link next to it is a defect, not a deliverable. This is the whole reason the workbook exists; if links are missing, the trip plan has failed at its primary job no matter how good the numbers look. Phase 5 includes a hard link-coverage gate that the workbook must pass before it ships.

## Architecture: orchestrator + phases

The orchestrator owns intake, the budget gate, sequencing, user check-ins, and final spreadsheet assembly. Each phase is delegated to a focused subagent whose brief lives in `agents/`:

| Phase | What it answers | Subagent brief |
|---|---|---|
| 0 — Intake | budget, dates flexibility, travelers, origin, region, prefs | *(orchestrator does this directly)* |
| 1 — Timing | when to go (weather, prices, crowds, events) | `agents/timing.md` |
| 2 — Itinerary | which cities, night split, candidate activities, intercity legs | `agents/itinerary.md` |
| 3 — Lodging & flights | live, bookable hotels (all four platforms) + flights | `agents/lodging-and-flights.md` |
| 4 — Pricing & extras | price every activity + food, transport, taxes, visa, insurance | `agents/activity-pricing.md` |
| 5 — Assembly | build + select + present the spreadsheet | *(orchestrator does this directly)* |

**Dependencies / parallelism:** 1 → 2 are sequential (cities/activities can depend on season). Once Phase 2 fixes the stops + dates, **Phases 3 and 4 are independent — run them in parallel.** Phase 5 waits for both.

## Execution modes

- **Orchestrated (subagents available — e.g. Claude Code, Cowork):** for each phase, spawn a subagent, handing it the brief from `agents/<phase>.md` plus the shared context (below). Spawn the Phase 3 and Phase 4 subagents **concurrently**. Collect each subagent's structured return, validate it, and proceed. If a subagent lacks a needed tool (e.g. the travel connectors), the orchestrator runs that phase itself.
- **Inline (no subagents — e.g. Claude.ai chat):** the orchestrator performs each phase itself, in sequence, following the same briefs and producing the same structured outputs. Same phases, no delegation or parallelism.

Either way, the phase briefs and output contracts are identical — they're the source of truth for what each phase must produce.

## Phase 0 — Intake & budget gate (orchestrator)

**Ask about budget FIRST. Always.** Use `ask_user_input_v0` with tappable options: *Economical* · *Mid-range* · *Comfortable* · *Splurge* · *I'll give a number*. Translate a tier into per-night price ceilings per city (scaled to the destination's cost level); treat a number as the ceiling. Never assume a tier.

In the same prompt, gather anything else missing: **destinations/region**, **dates** (fixed, or a flexible window to optimize in Phase 1), **travelers** (adults + **children's ages**), **origin** airport, **currency** (default to the user's home currency), and **preferences** (interests, pace, neighborhood, cabin class, nonstop). Bundle questions; don't interrogate one at a time. Then populate the shared context and begin the phases.

## Phases 1–4 (delegated)

Run each per its brief in `agents/`. Between phases, do quick **pulse-checks** with the user when a real choice exists ("These date windows trade weather for price — which way?", "Two riverside hotels vs. one cheaper inland — lean which way?"). Carry each phase's structured output forward in the shared context.

- **Phase 1 — Timing** (`agents/timing.md`): recommend/validate the date window.
- **Phase 2 — Itinerary** (`agents/itinerary.md`): lock stops, nights per city, candidate activities, intercity legs.
- **Phase 3 — Lodging & flights** (`agents/lodging-and-flights.md`): live hotels across **all four platforms (Booking.com, Expedia, Tripadvisor, lastminute.com)** + live flights. Follows `references/connectors.md`. Returns per-segment platform coverage so skipped sources are visible.
- **Phase 4 — Pricing & extras** (`agents/activity-pricing.md`): price every activity and all on-the-ground costs.

## Phase 5 — Assembly (orchestrator)

Read the `xlsx` skill, then build per `references/spreadsheet-style.md`, which uses the **bundled `scripts/trip_xlsx.py` builder** — import it rather than re-writing styling/hyperlink/audit/recalc boilerplate each run (that rewrite is slow and is where link bugs creep in). Your build code holds only the trip-specific tabs, columns, and cost formulas. From the gathered options, select a recommended option per segment (price + reviews + location, within budget), confirming with the user. Tabs: **Trip Summary** (always), **Hotels** (always, names hyperlinked, chosen feeds the total), **Flights** (always, nonstop preferred, chosen feeds the total), **Activities** (priced, with links), **Itinerary** (when useful). Inputs blue, formulas black, cross-sheet links green, hyperlinks blue-underlined; currency format throughout.

**Link-coverage gate (must pass before shipping).** Every priced row on Hotels, Flights, and Activities must carry a real, clickable hyperlink (an actual `cell.hyperlink`, not just blue-colored text). Build links with `tx.link()` (it sets the hyperlink and refuses an empty URL), then let `tx.finalize(..., link_checks=[…])` audit the finished file — its report lists any priced row missing a link and flags formula errors. If `report["missing_links"]` is non-empty, **go back and fill them in** — recover the URL from the phase output, or for a row that genuinely has no bookable page, link the provider's search results or the official/authoritative source so the cell is still clickable — then re-finalize. Do not ship a workbook the audit flags. If a specific link truly cannot be obtained, the cell must still link *somewhere* useful (a search for that exact hotel/flight/activity) and the Notes column must say the deep link wasn't available — never leave a bare price.

Save to `/mnt/user-data/outputs/`, `present_files`, and give a short summary (total, per person, chosen picks) — and confirm the link-audit passed.

## Shared context (orchestrator maintains; passed to every subagent)

```
trip_name, origin_airport,
destinations / region,
dates: { fixed: bool, window or start+end },
nights_total, nights_per_city,
travelers: { adults, children_ages: [] },
budget: { tier, per_night_ceiling_by_city, total_ceiling, currency },
prefs: { interests, pace, neighborhood, cabin_class, nonstop, ... },
fx: { eur_to_home, usd_to_home },
phase_outputs: { timing, itinerary, lodging_flights, pricing }
```

## Defaults & rules (memorize)

1. **Ask budget FIRST — always.** No exceptions.
2. **Flights: nonstop whenever available, economy** — unless the user states otherwise.
3. **Query ALL FOUR hotel platforms every time — Booking.com, Expedia, Tripadvisor, and lastminute.com** — for every stay segment; choose on price + reviews + location; pulse-check with the user. Record which platforms were actually queried per segment and name any that errored or returned nothing — never let one (e.g. Booking.com) get silently skipped.
4. **Every priced row gets its own clickable booking/source link — no exceptions.** Hotels, flights, and activities each carry a working hyperlink on their row. A price without a link beside it is incomplete work. If a true deep link isn't available for some row, link a search for that exact item (or the official source) and note it — never leave a bare number. Phase 5's link-audit enforces this.
5. **Never auto-book or enter payment details** — the user books via the links.
6. **Default to the user's home currency;** keep FX cells adjustable.
7. **Flag estimates vs. live prices** in the notes column.
8. **Note any provider or subagent that failed** so the user knows the data isn't complete.
9. **Run Phases 3 & 4 in parallel** when subagents are available.
