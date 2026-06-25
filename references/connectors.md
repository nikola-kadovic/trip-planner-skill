# Travel Connectors — parameters, response fields, and gotchas

**Use every connector available, every time.** Hotels come from **four** sources — `Booking.com`, `Expedia`, `Tripadvisor`, and `lastminute.com` — and you query **all four** for each stay segment, not a favorite two or three. Flights come from `Expedia` (and lastminute's package route when relevant). The Booking.com / Expedia / Tripadvisor tools render an interactive widget in the chat; lastminute returns structured data you read directly. Either way, read the data fields below to populate the spreadsheet; don't paste raw JSON back to the user.

## Hotels

### Booking.com:accommodations_search
Key params: `destination` (single place name; or `coordinates` for "near X"), `checkin_date`, `checkout_date` (`YYYY-MM-DD`), `number_of_adults`, `children_ages` (REQUIRED if any children), `number_of_rooms`, `currency` (set to the user's home currency, e.g. `CAD`), `user_country_code` (e.g. `ca`), `user_locale` (e.g. `en`), `price` (`{minimum, maximum}` per night), `minimum_review_score` (7, 8, or 9), `star_rating`, `facilities`, `accommodation_types`, `user_query` (short natural-language summary of only what the user explicitly asked).
Read from each item in `accommodations[]`: `name`, `url` (**directly bookable**, already includes the dates + currency), `price.book` (**total for the whole stay** in the currency you requested), `rating.review_score` (0–10), `rating.number_of_reviews`, `rating.stars`, `location.district_name`, `location.address`, `location.coordinates`.

### Expedia:search_hotels
Key params: `destination`, `check_in_date`, `check_out_date`, `adult_count`, `children_age_list`, `rooms`, `amenities`, `star_ratings` (`[1..5]`), `guest_rating` (`GOOD`/`VERY_GOOD`/`WONDERFUL`), `max_nightly_price` or `max_total_price`, `sort_type` (`CHEAPEST`/`MOST_EXPENSIVE`/`NEAREST`/`FARTHEST`), `property_themes`, `query_text` (verbatim user messages), `user_location` (`{city, region, country}` — the user's home, not the destination), `user_locale`, `client_device_info` (`{agent_name: "ClaudeAI", device_type: "desktop"}`).
Read from each item in `data[]`: `hotel_name`, `star_rating`, `guest_rating` (0–10), `guest_review_count`, `avg_nightly_price`, `total_price` (**whole stay**), `currency` (**usually USD — convert to display currency**), `url` (bookable), `geo_location`.

### Tripadvisor:search_hotels
Key params: `location` (string; for "near X" pass just the landmark, e.g. `Louvre Museum, Paris`), `checkIn`, `checkOut`, `guests`, `childrenAges`, `rooms`, `filters` (`{amenities, starRatings, minPrice, maxPrice, minReviewScore (0–5 bubble), selectFreeCancellation, ...}`), `sortBy` (`BEST_VALUE`/`PRICE_LOW_TO_HIGH`/`DISTANCE`/`POPULARITY`), `limit` (use 20), `mcpServerVersion` (`{version: "V2026_0327"}`), `requestContext` (`{clientType: "DESKTOP"}`).
Note: Tripadvisor uses a **0–5 "bubble rating"**, not a 0–10 score — keep that straight when comparing across platforms. This endpoint can throw an internal error; if it does, retry once with a simpler `location`, then skip it and note the omission.

### lastminute.com:search_only_hotel
**Mandatory pre-step:** lastminute won't take a raw city name. First call `lastminute.com:resolve_destination_id(city_name, lang)` (e.g. `lang="en"`). If it returns `success: true`, use the numeric `id` as `destination`. If `success: false`, it returns a `suggestions` list — pick the right city (or ask the user) before searching; don't guess. A 3-letter IATA code (e.g. `MAD`) works as a fallback only when the user explicitly gives one.
Key params: `destination` (the resolved numeric ID, preferred), `date_from`, `date_to` (`YYYY-MM-DD`), `adults` (string; for multiple rooms `"2;2"`), `ch_ages` (string of child ages, REQUIRED if any children, e.g. `"6,9"`), `hotel_stars` (CSV like `"3,4,5"`), `accommodation_type`, `accommodation_facilities` (CSV of numeric IDs — e.g. `0`=Free Wi-Fi, `4`=Pool, `6`=Spa), `price` (per-person/night cap as `"max"` or `"min,max"`, digits only), `sort` (`recommended`/`price`/`stars`/`review_score`/`distance`), `max_results` (use up to 20), `lang`.
Read from each result item: `index`, `hotel.name`, `hotel.stars`, `hotel.rating`, `price_total` + `currency` (**whole stay**), `hotel.distance_km`, `hotel.amenities`, `hotel.main_image`, plus the internal `search_id` and `internal_id_hotel` (keep these — you need them to get a bookable link; don't show them to the user).
**Getting a bookable link (important — lastminute search results don't include one).** For the option(s) you'll put in the sheet, drill in: `lastminute.com:select_hotel_options(search_id, hotel_internal_id, date_from, date_to)` → returns room rates with a `pricing_id` and per-room `rate_id`. Then `lastminute.com:generate_booking_link(pricing_id, rate_id)` → returns the **direct booking URL**. Use that URL as the row's hyperlink. At minimum do this for the chosen lastminute option; do it for the other lastminute rows you list too so every one is clickable.

### lastminute.com:search_flight_and_hotel_package (optional — package deals)
When a bundled flight+hotel package may beat booking separately, also run `search_flight_and_hotel_package(origin, destination, date_from, date_to, adults, ch_ages, ...)`. Here `origin`/`destination` must be **3-letter IATA codes** (e.g. `YYZ`, `BCN`), not city names. Surface a strong package as an alternative row/note in the workbook (label it "package — flight+hotel"), with its booking link via the same `select_hotel_options` → `generate_booking_link` flow.

## Flights

### Expedia:search_flights
Key params: `origin`, `destination` (IATA codes preferred, e.g. `YYZ`, `LIS`), `departure_date`, `return_date` (omit for one-way), `adult_count`, `children_age_list`, `cabin_class` (default `ECONOMY`), `number_of_stops` (**set `0` to bias nonstop**), `filter_nearby_airport`, `filter_basic_economy`, `price_min`/`price_max`, `sort_type` (`PRICE`/`DURATION`), `limit`, `query_text`, `user_location`, `user_locale`, `client_device_info`.
Read from each item in `options[]`: `slices[]` (one `outbound`, one `inbound`), each slice's `number_of_stops`, `flight_duration`, and `legs[]` (`marketing_airline_name`, `flight_number`); and `price.total_price` (`{value, currency}` — **USD, for ALL adults combined**), `price.average_price_per_ticket`, plus `refundable` and `fare_options[].baggage_fees` if relevant.
**Nonstop preference:** even with `number_of_stops=0`, verify every slice's `number_of_stops == 0` before marking an option nonstop. No per-option booking deep link is returned — build an Expedia Flights-Search URL for the link, e.g.
`https://www.expedia.com/Flights-Search?leg1=from:Toronto(YYZ),to:Lisbon(LIS),departure:MM/DD/YYYY&leg2=from:Lisbon(LIS),to:Toronto(YYZ),departure:MM/DD/YYYY&passengers=adults:2&trip=roundtrip`

## Cross-platform gotchas

- **Query all four hotel sources, and prove it.** Booking.com, Expedia, Tripadvisor, and lastminute.com each get called for every stay segment. Record per segment which ones you actually hit and how many results each returned. A platform that silently never gets queried is the failure mode to avoid — if one returns nothing or errors, that's fine, but it must be *named* in `unavailable_providers`, not quietly dropped.
- **lastminute needs a resolved destination ID and a follow-up call for the link.** Always `resolve_destination_id` first; the bookable URL only comes from `select_hotel_options` → `generate_booking_link`, not the search result. Budget the extra calls for it.
- **Currency normalization.** Booking.com returns the currency you request (set it to the user's home currency → no conversion). Expedia hotels **and** flights return **USD**. lastminute returns its own `currency` field (often EUR/GBP) — read it and convert via an FX cell. Convert any non-home-currency row to the display currency via an FX input cell in the workbook.
- **What the price covers.** Booking `price.book` and Expedia hotel `total_price` are the **whole stay**. Expedia flight `total_price` is **all passengers combined** (don't multiply by travelers again); `average_price_per_ticket` is per person.
- **Rating scales differ.** Booking & Expedia: 0–10. Tripadvisor: 0–5 bubbles. Normalize or label when comparing.
- **Booking, not browsing, is off-limits.** These are consumer travel partners; the user naming them (or invoking trip planning) authorizes *searching*. Never auto-book, submit a form, or enter payment — only surface links.
