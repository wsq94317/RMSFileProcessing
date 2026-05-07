#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
import sqlite3
from collections import Counter, defaultdict
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
TABLES_DIR = ROOT / "tables"
ASI_DIR = TABLES_DIR / "QVB ASI Backup"
OUTPUT_DIR = ROOT / "outputs"


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def clean_single_line(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def as_datetime(value: Any) -> datetime | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    text = clean_single_line(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def as_date(value: Any) -> date | None:
    dt = as_datetime(value)
    return dt.date() if dt else None


def iso_dt(value: datetime | None) -> str:
    return value.isoformat(sep=" ") if value else ""


def iso_date(value: date | None) -> str:
    return value.isoformat() if value else ""


def amount_to_float(value: Any) -> float | None:
    text = clean_single_line(value)
    if not text:
        return None
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", ".", "-", "-."}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_name(value: Any) -> str:
    text = clean_single_line(value).upper()
    text = re.sub(r"[^A-Z0-9\u4E00-\u9FFF]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def name_key(value: Any) -> str:
    tokens = normalize_name(value).split()
    return " ".join(sorted(tokens))


def parse_room_qty(room_text: str) -> int:
    match = re.match(r"\s*(\d+)\s*x\b", room_text, re.I)
    return int(match.group(1)) if match else 1


def read_siteminder() -> pd.DataFrame:
    frames = []
    for path in sorted(TABLES_DIR.glob("reservations_summary_report*.csv")):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["source_file"] = path.name
        frames.append(df)
    if not frames:
        return pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True)
    original_count = len(raw)
    raw = raw.drop_duplicates(subset=[c for c in raw.columns if c != "source_file"]).copy()
    raw["source_files"] = raw.groupby("Booking reference")["source_file"].transform(
        lambda s: "|".join(sorted(set(s)))
    )
    raw = raw.drop_duplicates(subset=["Booking reference"], keep="last").copy()

    raw["booking_reference"] = raw["Booking reference"].map(clean_single_line)
    raw["guest_names"] = raw["Guest names"].map(clean_single_line)
    raw["guest_name_norm"] = raw["guest_names"].map(normalize_name)
    raw["guest_name_key"] = raw["guest_names"].map(name_key)
    raw["checkin_date"] = raw["Check-in"].map(lambda v: iso_date(as_date(v)))
    raw["checkout_date"] = raw["Check-out"].map(lambda v: iso_date(as_date(v)))
    raw["channel"] = raw["Channel"].map(clean_single_line)
    raw["affiliated_channel"] = raw["Affiliated Channel"].map(clean_single_line)
    raw["referral"] = raw["Referral"].map(clean_single_line)
    raw["room"] = raw["Room"].map(clean_single_line)
    raw["room_qty"] = raw["room"].map(parse_room_qty)
    raw["booked_on"] = raw["Booked-on date"].map(lambda v: iso_dt(as_datetime(v)))
    raw["modified_on"] = raw["Modified-on date"].map(lambda v: iso_dt(as_datetime(v)))
    raw["cancelled_on"] = raw["Cancelled-on date"].map(lambda v: iso_dt(as_datetime(v)))
    raw["booking_status"] = raw["Booking status"].map(clean_single_line)
    raw["occupancy"] = raw["Occupancy"].map(clean_single_line)
    raw["total_price_aud"] = raw["Total price"].map(amount_to_float)
    raw["connectivity_status"] = raw["Connectivity status"].map(clean_single_line)
    raw["connectivity_last_sent_at"] = raw["Connectivity last sent at"].map(
        lambda v: iso_dt(as_datetime(v))
    )
    raw["raw_duplicate_export_rows"] = original_count - len(raw)

    cols = [
        "booking_reference",
        "guest_names",
        "guest_name_norm",
        "guest_name_key",
        "checkin_date",
        "checkout_date",
        "channel",
        "affiliated_channel",
        "referral",
        "room",
        "room_qty",
        "booked_on",
        "modified_on",
        "cancelled_on",
        "booking_status",
        "occupancy",
        "total_price_aud",
        "connectivity_status",
        "connectivity_last_sent_at",
        "source_files",
        "raw_duplicate_export_rows",
    ]
    return raw[cols].sort_values(["checkin_date", "checkout_date", "guest_name_key"]).reset_index(drop=True)


def read_arrivals() -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for path in sorted(ASI_DIR.glob("*Arrival*.xls")):
        df = pd.read_excel(path, header=None, engine="xlrd")
        current_arrival_group = ""
        for idx, row in df.iterrows():
            first = clean_single_line(row.get(0, ""))
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", first):
                current_arrival_group = iso_date(as_date(first))
                continue
            arrival_dt = as_datetime(row.get(1))
            checkout_dt = as_datetime(row.get(4))
            guest_name = clean_single_line(row.get(8))
            room = clean_single_line(row.get(0))
            if not (arrival_dt and checkout_dt and guest_name and room):
                continue
            rows.append(
                {
                    "asi_source": "arrival",
                    "asi_row_id": f"arrival:{path.name}:{idx + 1}",
                    "source_file": path.name,
                    "source_row": idx + 1,
                    "arrival_group_date": current_arrival_group,
                    "room_name": room,
                    "checkin_at": iso_dt(arrival_dt),
                    "checkout_at": iso_dt(checkout_dt),
                    "checkin_date": iso_date(arrival_dt.date()),
                    "checkout_date": iso_date(checkout_dt.date()),
                    "guest_name": guest_name,
                    "guest_name_norm": normalize_name(guest_name),
                    "guest_name_key": name_key(guest_name),
                    "guest_category": clean_single_line(row.get(11)),
                    "guest_count": clean_single_line(row.get(12)),
                    "address": clean_single_line(row.get(16)),
                    "source": clean_single_line(row.get(19)),
                    "total_charges": amount_to_float(row.get(21)),
                    "payments": amount_to_float(row.get(25)),
                    "status": clean_single_line(row.get(26)),
                    "remark": clean_text(row.get(29)),
                }
            )
    return pd.DataFrame(rows)


def departure_date_from_file(path: Path) -> str:
    match = re.search(r"_(\d{2}-[A-Za-z]{3}-\d{4})_", path.name)
    if not match:
        return ""
    parsed = datetime.strptime(match.group(1), "%d-%b-%Y")
    return parsed.date().isoformat()


def read_departures() -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for path in sorted(ASI_DIR.glob("*Departure*.xls")):
        df = pd.read_excel(path, header=None, engine="xlrd")
        report_date = departure_date_from_file(path)
        for idx, row in df.iterrows():
            guest_name = clean_single_line(row.get(0))
            checkin_dt = as_datetime(row.get(5))
            checkout_dt = as_datetime(row.get(6))
            room = clean_single_line(row.get(3))
            if not (guest_name and checkin_dt and checkout_dt and room):
                continue
            rows.append(
                {
                    "asi_source": "departure",
                    "asi_row_id": f"departure:{path.name}:{idx + 1}",
                    "source_file": path.name,
                    "source_row": idx + 1,
                    "report_date": report_date,
                    "room_name": room,
                    "checkin_at": iso_dt(checkin_dt),
                    "checkout_at": iso_dt(checkout_dt),
                    "checkin_date": iso_date(checkin_dt.date()),
                    "checkout_date": iso_date(checkout_dt.date()),
                    "guest_name": guest_name,
                    "guest_name_norm": normalize_name(guest_name),
                    "guest_name_key": name_key(guest_name),
                    "balance": amount_to_float(row.get(7)),
                    "guest_remark": clean_text(row.get(9)),
                }
            )
    return pd.DataFrame(rows)


def name_score(asi_name: str, sm_name: str) -> int:
    asi_norm, sm_norm = normalize_name(asi_name), normalize_name(sm_name)
    asi_key, sm_key = name_key(asi_name), name_key(sm_name)
    asi_tokens = set(asi_key.split())
    sm_tokens = set(sm_key.split())
    if not asi_norm or not sm_norm:
        return 0
    if asi_norm == sm_norm:
        return 100
    if asi_key == sm_key:
        return 98
    if len(asi_tokens) >= 2 and asi_tokens.issubset(sm_tokens):
        return 94
    if len(sm_tokens) >= 2 and sm_tokens.issubset(asi_tokens):
        return 94
    if asi_norm in sm_norm or sm_norm in asi_norm:
        return 92
    ratio = SequenceMatcher(None, asi_key, sm_key).ratio()
    if ratio >= 0.92:
        return 88
    if ratio >= 0.86:
        return 82
    if asi_tokens & sm_tokens and ratio >= 0.78:
        return 72
    return int(ratio * 70)


def build_match_candidates(asi: pd.DataFrame, sm: pd.DataFrame, source_name: str) -> pd.DataFrame:
    if asi.empty or sm.empty:
        return pd.DataFrame()
    active_sm = sm[sm["booking_status"].str.upper().ne("CANCELLED")].copy()
    by_dates: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for rec in active_sm.to_dict("records"):
        by_dates[(rec["checkin_date"], rec["checkout_date"])].append(rec)

    rows: list[dict[str, Any]] = []
    for rec in asi.to_dict("records"):
        candidates = by_dates.get((rec["checkin_date"], rec["checkout_date"]), [])
        for cand in candidates:
            score = name_score(rec["guest_name"], cand["guest_names"])
            if score >= 72:
                rows.append(
                    {
                        "asi_source": source_name,
                        "asi_row_id": rec["asi_row_id"],
                        "booking_reference": cand["booking_reference"],
                        "match_score": score,
                        "asi_guest_name": rec["guest_name"],
                        "siteminder_guest_names": cand["guest_names"],
                        "checkin_date": rec["checkin_date"],
                        "checkout_date": rec["checkout_date"],
                        "siteminder_status": cand["booking_status"],
                        "siteminder_channel": cand["channel"],
                        "siteminder_room": cand["room"],
                    }
                )
    if not rows:
        return pd.DataFrame()
    candidates = pd.DataFrame(rows)
    candidates = candidates.sort_values(
        ["asi_source", "asi_row_id", "match_score", "booking_reference"],
        ascending=[True, True, False, True],
    ).reset_index(drop=True)
    candidates["candidate_rank"] = candidates.groupby(["asi_source", "asi_row_id"]).cumcount() + 1
    return candidates


def build_matches(asi: pd.DataFrame, candidates: pd.DataFrame, source_name: str) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    candidate_groups = {
        key: group.sort_values(["match_score", "booking_reference"], ascending=[False, True])
        for key, group in candidates[candidates["asi_source"].eq(source_name)].groupby("asi_row_id")
    } if not candidates.empty else {}
    for rec in asi.to_dict("records"):
        scored = candidate_groups.get(rec["asi_row_id"], pd.DataFrame())
        if scored.empty:
            rows.append(
                {
                    "asi_source": source_name,
                    "asi_row_id": rec["asi_row_id"],
                    "booking_reference": "",
                    "match_status": "unmatched",
                    "match_score": 0,
                    "candidate_count": 0,
                    "second_best_score": 0,
                    "match_reason": "same dates, no sufficiently similar active SiteMinder guest name",
                }
            )
            continue
        best = scored.iloc[0]
        best_score = int(best["match_score"])
        second_score = int(scored.iloc[1]["match_score"]) if len(scored) > 1 else 0
        ambiguous_same_top = int(scored["match_score"].eq(best_score).sum()) > 1
        status = "matched"
        reason = "same check-in/check-out dates and guest name"
        if ambiguous_same_top or best_score - second_score <= 3:
            status = "ambiguous"
            reason = "multiple SiteMinder candidates have similar scores"
        rows.append(
                {
                    "asi_source": source_name,
                    "asi_row_id": rec["asi_row_id"],
                    "booking_reference": best["booking_reference"],
                    "match_status": status,
                    "match_score": best_score,
                    "candidate_count": int(len(scored)),
                    "second_best_score": second_score,
                    "match_reason": reason,
                }
        )
    return pd.DataFrame(rows)


def write_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL)


def write_sqlite(path: Path, tables: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()
    con = sqlite3.connect(path)
    try:
        for name, df in tables.items():
            df.to_sql(name, con, index=False, if_exists="replace")
        con.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_sm_ref ON siteminder_bookings_clean(booking_reference);
            CREATE INDEX IF NOT EXISTS idx_sm_dates_name ON siteminder_bookings_clean(checkin_date, checkout_date, guest_name_key);
            CREATE INDEX IF NOT EXISTS idx_arrival_dates_name ON asi_arrivals_clean(checkin_date, checkout_date, guest_name_key);
            CREATE INDEX IF NOT EXISTS idx_departure_dates_name ON asi_departures_clean(checkin_date, checkout_date, guest_name_key);
            CREATE INDEX IF NOT EXISTS idx_matches_ref ON booking_matches(booking_reference);
            CREATE INDEX IF NOT EXISTS idx_matches_asi ON booking_matches(asi_row_id);
            """
        )
        con.commit()
    finally:
        con.close()


def build_booking_master(
    sm: pd.DataFrame, arrivals: pd.DataFrame, departures: pd.DataFrame, matches: pd.DataFrame
) -> pd.DataFrame:
    exact_matches = matches[matches["match_status"].eq("matched")].copy()
    arrival_map = arrivals.set_index("asi_row_id").to_dict("index") if not arrivals.empty else {}
    departure_map = departures.set_index("asi_row_id").to_dict("index") if not departures.empty else {}
    grouped = defaultdict(lambda: {"arrival_remarks": [], "departure_remarks": [], "asi_rooms": []})
    for m in exact_matches.to_dict("records"):
        ref = m["booking_reference"]
        if m["asi_source"] == "arrival" and m["asi_row_id"] in arrival_map:
            r = arrival_map[m["asi_row_id"]]
            if r.get("remark"):
                grouped[ref]["arrival_remarks"].append(r["remark"])
            grouped[ref]["asi_rooms"].append(r.get("room_name", ""))
        if m["asi_source"] == "departure" and m["asi_row_id"] in departure_map:
            r = departure_map[m["asi_row_id"]]
            if r.get("guest_remark"):
                grouped[ref]["departure_remarks"].append(r["guest_remark"])
            grouped[ref]["asi_rooms"].append(r.get("room_name", ""))

    rows = []
    for rec in sm.to_dict("records"):
        ref = rec["booking_reference"]
        extra = grouped.get(ref, {})
        row = dict(rec)
        row["matched_asi_room_count"] = len([x for x in extra.get("asi_rooms", []) if x])
        row["matched_asi_rooms"] = "|".join(sorted(set(x for x in extra.get("asi_rooms", []) if x)))
        row["arrival_remarks"] = "\n---\n".join(dict.fromkeys(extra.get("arrival_remarks", [])))
        row["departure_remarks"] = "\n---\n".join(dict.fromkeys(extra.get("departure_remarks", [])))
        rows.append(row)
    return pd.DataFrame(rows)


def build_integrity_checks(
    sm: pd.DataFrame, arrivals: pd.DataFrame, departures: pd.DataFrame, matches: pd.DataFrame
) -> pd.DataFrame:
    arrival_lookup = arrivals.set_index("asi_row_id").to_dict("index") if not arrivals.empty else {}
    departure_lookup = departures.set_index("asi_row_id").to_dict("index") if not departures.empty else {}
    active_refs = set(sm[sm["booking_status"].str.upper().ne("CANCELLED")]["booking_reference"])

    grouped: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "arrival_matched_rows": 0,
            "departure_matched_rows": 0,
            "arrival_ambiguous_rows": 0,
            "departure_ambiguous_rows": 0,
            "arrival_rooms": set(),
            "departure_rooms": set(),
            "arrival_remarks_count": 0,
            "departure_remarks_count": 0,
        }
    )
    for row in matches.to_dict("records"):
        ref = row.get("booking_reference", "")
        if not ref:
            continue
        bucket = grouped[ref]
        source = row["asi_source"]
        status = row["match_status"]
        if source == "arrival":
            rec = arrival_lookup.get(row["asi_row_id"], {})
            if status == "matched":
                bucket["arrival_matched_rows"] += 1
                if rec.get("room_name"):
                    bucket["arrival_rooms"].add(rec["room_name"])
                if rec.get("remark"):
                    bucket["arrival_remarks_count"] += 1
            elif status == "ambiguous":
                bucket["arrival_ambiguous_rows"] += 1
        elif source == "departure":
            rec = departure_lookup.get(row["asi_row_id"], {})
            if status == "matched":
                bucket["departure_matched_rows"] += 1
                if rec.get("room_name"):
                    bucket["departure_rooms"].add(rec["room_name"])
                if rec.get("guest_remark"):
                    bucket["departure_remarks_count"] += 1
            elif status == "ambiguous":
                bucket["departure_ambiguous_rows"] += 1

    max_departure_report_date = departures["report_date"].max() if not departures.empty else ""
    rows = []
    for rec in sm.to_dict("records"):
        ref = rec["booking_reference"]
        bucket = grouped.get(ref, {})
        expected = int(rec.get("room_qty") or 1)
        arrival_count = int(bucket.get("arrival_matched_rows", 0))
        departure_count = int(bucket.get("departure_matched_rows", 0))
        active = ref in active_refs
        status = "cancelled_not_checked"
        issue = ""
        if active:
            if arrival_count >= expected:
                status = "complete_by_arrival"
            elif arrival_count + int(bucket.get("arrival_ambiguous_rows", 0)) >= expected:
                status = "needs_ambiguous_review"
                issue = "Arrival has enough rows only if ambiguous matches are accepted."
            elif rec["checkout_date"] <= max_departure_report_date and departure_count >= expected:
                status = "complete_by_departure_only"
                issue = "Departure confirms checked-out rows, but Arrival does not fully match."
            else:
                status = "missing_asi_room_rows"
                issue = "Matched ASI Arrival room rows are fewer than SiteMinder room quantity."
        rows.append(
            {
                "booking_reference": ref,
                "guest_names": rec["guest_names"],
                "booking_status": rec["booking_status"],
                "checkin_date": rec["checkin_date"],
                "checkout_date": rec["checkout_date"],
                "channel": rec["channel"],
                "siteminder_room": rec["room"],
                "expected_room_qty": expected,
                "arrival_matched_rows": arrival_count,
                "arrival_ambiguous_rows": int(bucket.get("arrival_ambiguous_rows", 0)),
                "arrival_matched_rooms": "|".join(sorted(bucket.get("arrival_rooms", set()))),
                "departure_matched_rows": departure_count,
                "departure_ambiguous_rows": int(bucket.get("departure_ambiguous_rows", 0)),
                "departure_matched_rooms": "|".join(sorted(bucket.get("departure_rooms", set()))),
                "arrival_remarks_count": int(bucket.get("arrival_remarks_count", 0)),
                "departure_remarks_count": int(bucket.get("departure_remarks_count", 0)),
                "integrity_status": status,
                "integrity_issue": issue,
            }
        )
    return pd.DataFrame(rows)


def summarize(sm: pd.DataFrame, arrivals: pd.DataFrame, departures: pd.DataFrame, matches: pd.DataFrame) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "siteminder_bookings": int(len(sm)),
        "siteminder_active_bookings": int(sm["booking_status"].str.upper().ne("CANCELLED").sum()) if not sm.empty else 0,
        "asi_arrival_rows": int(len(arrivals)),
        "asi_departure_rows": int(len(departures)),
        "arrival_date_range": [
            arrivals["checkin_date"].min() if not arrivals.empty else "",
            arrivals["checkin_date"].max() if not arrivals.empty else "",
        ],
        "departure_report_date_range": [
            departures["report_date"].min() if not departures.empty else "",
            departures["report_date"].max() if not departures.empty else "",
        ],
    }
    if not matches.empty:
        counts = Counter(matches["match_status"])
        summary["match_status_counts"] = dict(counts)
        for source in ["arrival", "departure"]:
            sub = matches[matches["asi_source"].eq(source)]
            summary[f"{source}_match_status_counts"] = dict(Counter(sub["match_status"]))
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()

    sm = read_siteminder()
    arrivals = read_arrivals()
    departures = read_departures()
    arrival_candidates = build_match_candidates(arrivals, sm, "arrival")
    departure_candidates = build_match_candidates(departures, sm, "departure")
    candidates = pd.concat([arrival_candidates, departure_candidates], ignore_index=True)
    arrival_matches = build_matches(arrivals, candidates, "arrival")
    departure_matches = build_matches(departures, candidates, "departure")
    matches = pd.concat([arrival_matches, departure_matches], ignore_index=True)
    master = build_booking_master(sm, arrivals, departures, matches)
    integrity = build_integrity_checks(sm, arrivals, departures, matches)

    out = args.output_dir
    write_csv(sm, out / "siteminder_bookings_clean.csv")
    write_csv(arrivals, out / "asi_arrivals_clean.csv")
    write_csv(departures, out / "asi_departures_clean.csv")
    write_csv(matches, out / "booking_matches.csv")
    write_csv(candidates, out / "booking_match_candidates.csv")
    write_csv(master, out / "booking_master.csv")
    write_csv(integrity, out / "booking_integrity_checks.csv")
    summary = summarize(sm, arrivals, departures, matches)
    (out / "database_summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    write_sqlite(
        out / "hotel_migration.sqlite",
        {
            "siteminder_bookings_clean": sm,
            "asi_arrivals_clean": arrivals,
            "asi_departures_clean": departures,
            "booking_matches": matches,
            "booking_match_candidates": candidates,
            "booking_master": master,
            "booking_integrity_checks": integrity,
        },
    )
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
