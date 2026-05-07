# ASI to RMS migration database

This project builds a cleaned migration database from:

- SiteMinder reservation summary CSV files in `tables/`
- ASI Guest Arrival Report `.xls` files in `tables/QVB ASI Backup/`
- ASI Guest Departure Report `.xls` files in `tables/QVB ASI Backup/`

Run:

```bash
PYTHONPATH=.codex_deps /Users/wsq/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_hotel_database.py
```

Outputs are written to `outputs/`:

- `hotel_migration.sqlite`: SQLite database with all cleaned tables.
- `booking_master.csv`: SiteMinder booking-level master table, with matched ASI arrival and departure remarks appended where confidence is high.
- `siteminder_bookings_clean.csv`: cleaned SiteMinder bookings.
- `asi_arrivals_clean.csv`: cleaned ASI arrival rows.
- `asi_departures_clean.csv`: cleaned ASI departure rows.
- `booking_matches.csv`: row-level ASI-to-SiteMinder match audit with score and status.
- `booking_match_candidates.csv`: all candidate SiteMinder bookings considered for matched or ambiguous ASI rows.
- `booking_integrity_checks.csv`: booking-level room-count completeness checks, especially for multi-room bookings.
- `database_summary.json`: counts and match summary.

Matching rules:

- SiteMinder cancelled bookings are kept in the database but are not used for default ASI matching.
- ASI rows match SiteMinder bookings by same check-in/check-out dates plus normalized guest name.
- Reversed names such as `Sarah Adams` and `Adams Sarah` are treated as equivalent.
- Short-name/full-name cases such as `Thayna Lima` and `Thayna Mendes De Freitas Lima` are treated as high-confidence matches.
- Rows with multiple close candidates are marked `ambiguous` and are not merged into `booking_master.csv` remarks.
- Rows with no sufficiently similar active SiteMinder booking are marked `unmatched`.

Integrity checks:

- For active SiteMinder bookings, `expected_room_qty` comes from the `N x ...` room quantity in SiteMinder.
- `arrival_matched_rows` is the main check that rooms were created in ASI.
- `complete_by_arrival` means matched ASI Arrival rows are greater than or equal to the expected room quantity.
- `needs_ambiguous_review` means the booking may be complete, but only after accepting ambiguous matches.
- `missing_asi_room_rows` means the active SiteMinder booking does not currently have enough matched ASI Arrival rows.
- `complete_by_departure_only` means Departure confirms enough checked-out rows, but Arrival did not fully match.

## Absolute RMS import master

The RMS import pipeline uses `tables/new/21860_AdvanceGuestSearch_20260507_203734.xls` as the absolute master file.
It never deletes master rows and appends only three fields to the right: `Payment Type`, `Note`, and `Active/Cancel`.

Run:

```bash
PYTHONPATH=.codex_deps /Users/wsq/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_rms_import_master.py
```

Outputs are written to `outputs/rms_import/`:

- `rms_absolute_master_merged.csv`: final merged CSV for RMS import review.
- `rms_absolute_master_merged.xlsx`: Excel version of the same final merged table.
- `rms_absolute_master_summary.json`: row counts and merge/payment summary.
- `audit_payment_type_blank.csv`: rows where the requested source rules leave `Payment Type` blank.
- `audit_expedia_payment_unmatched.csv`: Expedia rows whose `CRS Folio #` did not match `reservationsList.csv`.
- `audit_arrival_remark_not_matched.csv`: master rows where Arrival Remark was not added because exact name/date matching failed.
- `audit_active_cancel_unmatched.csv`: rows where `Active/Cancel` is still blank because SiteMinder did not match and Arrival List fallback was not allowed for that source.
- `audit_active_cancel_arrival_fallback.csv`: ASI/Mobile/blank-source rows where SiteMinder did not match and `Active/Cancel` was decided by Arrival List.
- `audit_business_source_ctrip_fixed.csv`: rows where blank `Business Source` was set to `Ctrip` from SiteMinder channel data.

Rules:

- The absolute master row count must stay unchanged; the script stops if the output row count differs.
- `Note` is copied from ASI Arrival Remark only when `First Name + Last Name`, check-in date, and check-out date exactly match the ASI Arrival Report.
- `Active/Cancel` is checked against all SiteMinder rows by `CRS Folio # = Booking reference`; cancelled bookings become `Cancel`, booked/modified bookings become `Active`.
- If SiteMinder does not match, `Active/Cancel` falls back to Arrival List exact matching only for `ASI`, `Mobile`, or blank `Business Source`: present in Arrival List = `Active`, otherwise `Cancel`.
- If master `Business Source` is blank but `CRS Folio #` matches SiteMinder and the SiteMinder channel is Ctrip/Trip.com, the output `Business Source` is set to `Ctrip`.
- Expedia `Payment Type` is based on `CRS Folio # = Reservation ID` in `tables/new/reservationsList.csv`.
- `Payment Type` mapping: Agoda/Ctrip/AirBnBXML = `Prepaid`; Hotelbeds/Hopper/Jetstar-Hooroo-Qantas/Restel/Traveloka/WebBeds = `VCC`; Anand Systems Booking Engine/Booking.com/Mobile = `POA`; Expedia uses the Expedia payment type file.

## Web app

Install dependencies on the server:

```bash
python3 -m pip install -r requirements.txt
```

Run the simple upload web app:

```bash
python3 web_app.py
```

Then open:

```text
http://SERVER_IP:8000
```

The upload page requires:

- ASI Advance Guest Search: the absolute master table. The output keeps all rows from this file.
- ASI Guest Arrival Report: the Arrival report containing the ASI `Remark`; exact guest name plus check-in/check-out date is required before copying it into `Note`.
- SiteMinder Reservations Summary CSV: one or more SiteMinder CSV exports; used to set `Active/Cancel`.
- Expedia Reservations List CSV: `reservationsList.csv`; used only for Expedia `Payment Type`.

For production, run Flask behind a proper WSGI server such as gunicorn or uwsgi, and protect the page because uploaded files contain guest personal data.

### Windows server quick start

Install Python 3.11+ for Windows, then run:

```bat
run_windows.bat
```

The script creates `.venv`, installs dependencies, and starts the app on port `8000`.

Open:

```text
http://SERVER_IP:8000
```

If Windows Firewall blocks the port, allow inbound TCP `8000` or change the port in `web_app.py`.
