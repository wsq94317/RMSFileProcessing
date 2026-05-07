#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
TABLES_DIR = ROOT / "tables"
NEW_DIR = TABLES_DIR / "new"
OUTPUT_DIR = ROOT / "outputs" / "rms_import"

MASTER_FILE = NEW_DIR / "21860_AdvanceGuestSearch_20260507_203734.xls"
ARRIVAL_FILE = NEW_DIR / "21860_GuestArrivalReport_07-May-2026_134356.xls"
EXPEDIA_FILE = NEW_DIR / "reservationsList.csv"


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def clean_line(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def as_datetime(value: Any) -> datetime | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    text = clean_line(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def as_date(value: Any) -> date | None:
    dt = as_datetime(value)
    return dt.date() if dt else None


def iso_date(value: date | None) -> str:
    return value.isoformat() if value else ""


def read_absolute_master(path: Path) -> pd.DataFrame:
    tables = pd.read_html(path, keep_default_na=False)
    if len(tables) != 1:
        raise ValueError(f"Expected one HTML table in {path}, found {len(tables)}")
    raw = tables[0]
    headers = [clean_line(v) for v in raw.iloc[0].tolist()]
    df = raw.iloc[1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df["_Absolute Master Row #"] = range(1, len(df) + 1)
    df.attrs["master_columns"] = headers
    for col in headers:
        df[col] = df[col].map(clean_line)
    required = ["First Name", "Last Name", "Date In", "Date Out", "CRS Folio #", "Business Source"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Master file missing required columns: {missing}")
    df["_master_guest_name_exact"] = (df["First Name"] + " " + df["Last Name"]).map(clean_line)
    df["_checkin_date"] = df["Date In"].map(lambda v: iso_date(as_date(v)))
    df["_checkout_date"] = df["Date Out"].map(lambda v: iso_date(as_date(v)))
    return df


def read_arrival_remarks(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_excel(path, header=None, engine="xlrd")
    rows: list[dict[str, Any]] = []
    for idx, row in df.iterrows():
        checkin = as_datetime(row.get(1))
        checkout = as_datetime(row.get(4))
        guest_name = clean_line(row.get(8))
        room = clean_line(row.get(0))
        if not (checkin and checkout and guest_name and room):
            continue
        rows.append(
            {
                "arrival_source_row": idx + 1,
                "arrival_room": room,
                "arrival_guest_name_exact": guest_name,
                "arrival_checkin_date": iso_date(checkin.date()),
                "arrival_checkout_date": iso_date(checkout.date()),
                "arrival_remark": clean_text(row.get(29)),
                "arrival_status": clean_line(row.get(26)),
                "arrival_source": clean_line(row.get(19)),
            }
        )
    arrival = pd.DataFrame(rows)
    if arrival.empty:
        return arrival, pd.DataFrame()

    grouped_rows = []
    for key, group in arrival.groupby(
        ["arrival_guest_name_exact", "arrival_checkin_date", "arrival_checkout_date"],
        dropna=False,
    ):
        remarks = [r for r in group["arrival_remark"].tolist() if clean_text(r)]
        unique_remarks = list(dict.fromkeys(remarks))
        grouped_rows.append(
            {
                "arrival_guest_name_exact": key[0],
                "arrival_checkin_date": key[1],
                "arrival_checkout_date": key[2],
                "ASI Arrival Remark": "\n---\n".join(unique_remarks),
                "ASI Arrival Match Count": int(len(group)),
                "ASI Arrival Rooms": "|".join(sorted(set(group["arrival_room"].astype(str)))),
                "ASI Arrival Statuses": "|".join(sorted(set(group["arrival_status"].astype(str)))),
                "ASI Arrival Source Rows": "|".join(str(x) for x in group["arrival_source_row"].tolist()),
            }
        )
    grouped = pd.DataFrame(grouped_rows)
    return arrival, grouped


def read_siteminder(paths: list[Path] | None = None) -> pd.DataFrame:
    frames = []
    source_paths = paths if paths is not None else sorted(TABLES_DIR.glob("reservations_summary_report*.csv"))
    for path in source_paths:
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["SiteMinder Source File"] = path.name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    sm = pd.concat(frames, ignore_index=True)
    sm = sm.drop_duplicates(subset=[c for c in sm.columns if c != "SiteMinder Source File"]).copy()
    grouped_rows = []
    for booking_ref, group in sm.groupby("Booking reference", dropna=False):
        statuses = [clean_line(v) for v in group["Booking status"].tolist() if clean_line(v)]
        channels = [clean_line(v) for v in group["Channel"].tolist() if clean_line(v)]
        affiliated_channels = [clean_line(v) for v in group["Affiliated Channel"].tolist() if clean_line(v)]
        non_cancelled = [s for s in statuses if s.upper() != "CANCELLED"]
        if not clean_line(booking_ref):
            continue
        grouped_rows.append(
            {
                "SM Booking reference": clean_line(booking_ref),
                "SM Statuses": "|".join(sorted(set(statuses))),
                "SM Channels": "|".join(sorted(set(channels))),
                "SM Affiliated Channels": "|".join(sorted(set(affiliated_channels))),
                "Active/Cancel": "Active" if non_cancelled else "Cancel",
                "SM Matched Row Count": int(len(group)),
            }
        )
    return pd.DataFrame(grouped_rows)


def read_expedia_payment_types(path: Path) -> pd.DataFrame:
    expedia = pd.read_csv(path, dtype=str, keep_default_na=False)
    required = ["Reservation ID", "Payment type"]
    missing = [c for c in required if c not in expedia.columns]
    if missing:
        raise ValueError(f"Expedia file missing required columns: {missing}")
    expedia = expedia.drop_duplicates(subset=["Reservation ID"], keep="last").copy()
    return expedia[["Reservation ID", "Payment type", "Confirmation #", "Status"]].rename(
        columns={
            "Reservation ID": "Expedia Reservation ID",
            "Payment type": "Expedia Payment Type",
            "Confirmation #": "Expedia Confirmation #",
            "Status": "Expedia Status",
        }
    )


def normalize_source(source: str) -> str:
    text = clean_line(source).lower()
    text = text.replace("&", "and")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def payment_method(row: pd.Series) -> str:
    source = normalize_source(row.get("Business Source", ""))
    expedia_payment_type = clean_line(row.get("Expedia Payment Type", ""))

    prepaid_sources = {"agoda", "ctrip", "airbnbxml"}
    vcc_sources = {
        "hotelbeds",
        "hopper",
        "jetstar hooroo qantas",
        "restel",
        "traveloka",
        "webbeds destinations of the world",
    }
    poa_sources = {"anand systems booking engine", "booking com", "mobile"}

    if source in prepaid_sources:
        return "Prepaid"
    if source in vcc_sources:
        return "VCC"
    if source in poa_sources:
        return "POA"
    if source == "expedia":
        if expedia_payment_type == "Hotel Collect":
            return "POA"
        if expedia_payment_type == "Expedia Collect":
            return "Prepaid"
        return ""
    return ""


def is_ctrip_siteminder(row: pd.Series) -> bool:
    text = normalize_source(
        " ".join(
            [
                clean_line(row.get("SM Channels", "")),
                clean_line(row.get("SM Affiliated Channels", "")),
            ]
        )
    )
    return any(token in text for token in ["ctrip", "trip com"])


def siteminder_match_key(row: pd.Series) -> str:
    crs_folio = clean_line(row.get("CRS Folio #", ""))
    if normalize_source(row.get("Business Source", "")) == "hotelbeds" and "-" in crs_folio:
        return crs_folio.split("-", 1)[1].strip()
    return crs_folio


def build_master(
    master_file: Path = MASTER_FILE,
    arrival_file: Path = ARRIVAL_FILE,
    expedia_file: Path = EXPEDIA_FILE,
    siteminder_files: list[Path] | None = None,
) -> tuple[pd.DataFrame, dict[str, Any], pd.DataFrame]:
    master = read_absolute_master(master_file)
    master_columns = list(master.attrs["master_columns"])
    master_original_rows = len(master)

    arrival_raw, arrival_grouped = read_arrival_remarks(arrival_file)
    merged = master.merge(
        arrival_grouped,
        how="left",
        left_on=["_master_guest_name_exact", "_checkin_date", "_checkout_date"],
        right_on=["arrival_guest_name_exact", "arrival_checkin_date", "arrival_checkout_date"],
        validate="many_to_one",
    )

    sm = read_siteminder(siteminder_files)
    if not sm.empty:
        merged["_siteminder_match_key"] = merged.apply(siteminder_match_key, axis=1)
        merged = merged.merge(
            sm,
            how="left",
            left_on="_siteminder_match_key",
            right_on="SM Booking reference",
            validate="many_to_one",
        )

    expedia = read_expedia_payment_types(expedia_file)
    merged = merged.merge(
        expedia,
        how="left",
        left_on="CRS Folio #",
        right_on="Expedia Reservation ID",
        validate="many_to_one",
    )
    original_source_norm = merged["Business Source"].map(normalize_source)
    fallback_allowed = original_source_norm.isin(["", "asi", "mobile", "anand systems booking engine"])
    source_was_blank = merged["Business Source"].fillna("").map(clean_line).eq("")
    ctrip_from_siteminder = source_was_blank & merged.apply(is_ctrip_siteminder, axis=1)
    merged.loc[ctrip_from_siteminder, "Business Source"] = "Ctrip"

    has_arrival = merged["ASI Arrival Match Count"].notna()
    missing_siteminder_status = merged["Active/Cancel"].fillna("").eq("")
    fallback_rows = missing_siteminder_status & fallback_allowed
    merged.loc[fallback_rows & has_arrival, "Active/Cancel"] = "Active"
    merged.loc[fallback_rows & ~has_arrival, "Active/Cancel"] = "Cancel"

    merged["Payment Type"] = merged.apply(payment_method, axis=1)
    merged["Note"] = merged["ASI Arrival Remark"].fillna("")
    merged["Active/Cancel"] = merged["Active/Cancel"].fillna("")

    merged["ASI Arrival Exact Match Key"] = (
        merged["_master_guest_name_exact"] + "|" + merged["_checkin_date"] + "|" + merged["_checkout_date"]
    )
    merged["ASI Arrival Remark Match Status"] = "not_matched"
    merged.loc[has_arrival, "ASI Arrival Remark Match Status"] = "matched_exact_name_dates"
    merged.loc[has_arrival & merged["ASI Arrival Remark"].fillna("").eq(""), "ASI Arrival Remark Match Status"] = (
        "matched_exact_name_dates_no_remark"
    )
    merged["SiteMinder Merge Status"] = "not_matched"
    merged.loc[merged.get("SM Booking reference", pd.Series("", index=merged.index)).fillna("").ne(""), "SiteMinder Merge Status"] = (
        "matched_active_not_cancelled"
    )
    merged["Expedia Payment Match Status"] = ""
    is_expedia = merged["Business Source"].map(normalize_source).eq("expedia")
    merged.loc[is_expedia, "Expedia Payment Match Status"] = "not_matched"
    merged.loc[is_expedia & merged["Expedia Payment Type"].fillna("").ne(""), "Expedia Payment Match Status"] = "matched"
    merged["Business Source Fix Status"] = ""
    merged.loc[ctrip_from_siteminder, "Business Source Fix Status"] = "blank_source_set_to_ctrip_from_siteminder"
    merged["Active/Cancel Source"] = "No SiteMinder match; fallback not applicable"
    merged.loc[fallback_rows, "Active/Cancel Source"] = "Arrival List fallback"
    merged.loc[merged.get("SM Booking reference", pd.Series("", index=merged.index)).fillna("").ne(""), "Active/Cancel Source"] = "SiteMinder"

    if len(merged) != master_original_rows:
        raise RuntimeError(f"Output row count changed from {master_original_rows} to {len(merged)}")
    if not merged["_Absolute Master Row #"].is_unique:
        raise RuntimeError("Internal absolute master row number is no longer unique")

    output = merged[master_columns + ["Payment Type", "Note", "Active/Cancel"]].copy()
    output = output.fillna("")

    summary = {
        "master_file": str(master_file),
        "arrival_file": str(arrival_file),
        "expedia_file": str(expedia_file),
        "siteminder_files": [str(p) for p in (siteminder_files or sorted(TABLES_DIR.glob("reservations_summary_report*.csv")))],
        "master_rows_input": master_original_rows,
        "master_rows_output": int(len(output)),
        "arrival_rows_raw": int(len(arrival_raw)),
        "arrival_exact_name_date_rows_matched": int(has_arrival.sum()),
        "arrival_exact_name_date_rows_with_remark": int(
            (has_arrival & merged["ASI Arrival Remark"].fillna("").ne("")).sum()
        ),
        "siteminder_booking_refs": int(len(sm)),
        "hotelbeds_rows_with_trimmed_siteminder_match_key": int(
            (original_source_norm.eq("hotelbeds") & merged.get("_siteminder_match_key", pd.Series("", index=merged.index)).ne(merged["CRS Folio #"])).sum()
        ),
        "active_cancel_counts": output["Active/Cancel"].value_counts(dropna=False).to_dict(),
        "active_cancel_source_counts": merged["Active/Cancel Source"].value_counts(dropna=False).to_dict(),
        "active_cancel_blank_after_restricted_fallback": int(output["Active/Cancel"].eq("").sum()),
        "business_source_blank_set_to_ctrip_from_siteminder": int(ctrip_from_siteminder.sum()),
        "expedia_rows": int(is_expedia.sum()),
        "expedia_payment_rows_matched": int(merged["Expedia Payment Match Status"].eq("matched").sum()),
        "payment_type_counts": output["Payment Type"].value_counts(dropna=False).to_dict(),
        "business_source_unmapped_payment_counts": output.loc[
            output["Payment Type"].eq(""), "Business Source"
        ].value_counts(dropna=False).to_dict(),
    }
    audit = merged.fillna("")
    return output, summary, audit


def write_outputs(output: pd.DataFrame, summary: dict[str, Any], audit: pd.DataFrame, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / "rms_absolute_master_merged.csv"
    xlsx_path = out_dir / "rms_absolute_master_merged.xlsx"
    summary_path = out_dir / "rms_absolute_master_summary.json"
    blank_payment_path = out_dir / "audit_payment_type_blank.csv"
    expedia_unmatched_path = out_dir / "audit_expedia_payment_unmatched.csv"
    arrival_no_match_path = out_dir / "audit_arrival_remark_not_matched.csv"
    active_cancel_unmatched_path = out_dir / "audit_active_cancel_unmatched.csv"
    active_cancel_fallback_path = out_dir / "audit_active_cancel_arrival_fallback.csv"
    ctrip_source_fixed_path = out_dir / "audit_business_source_ctrip_fixed.csv"

    output.to_csv(csv_path, index=False, quoting=csv.QUOTE_MINIMAL)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        output.to_excel(writer, index=False, sheet_name="RMS Import Master")
        ws = writer.book["RMS Import Master"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    audit[audit["Payment Type"].eq("")].to_csv(blank_payment_path, index=False)
    audit[
        audit["Business Source"].map(normalize_source).eq("expedia")
        & audit["Expedia Payment Match Status"].ne("matched")
    ].to_csv(expedia_unmatched_path, index=False)
    audit[audit["ASI Arrival Remark Match Status"].eq("not_matched")].to_csv(arrival_no_match_path, index=False)
    audit[audit["Active/Cancel"].eq("")].to_csv(active_cancel_unmatched_path, index=False)
    audit[audit["Active/Cancel Source"].eq("Arrival List fallback")].to_csv(active_cancel_fallback_path, index=False)
    audit[audit["Business Source Fix Status"].eq("blank_source_set_to_ctrip_from_siteminder")].to_csv(
        ctrip_source_fixed_path, index=False
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--master-file", type=Path, default=MASTER_FILE)
    parser.add_argument("--arrival-file", type=Path, default=ARRIVAL_FILE)
    parser.add_argument("--expedia-file", type=Path, default=EXPEDIA_FILE)
    parser.add_argument("--siteminder-file", type=Path, action="append", dest="siteminder_files")
    args = parser.parse_args()
    output, summary, audit = build_master(
        master_file=args.master_file,
        arrival_file=args.arrival_file,
        expedia_file=args.expedia_file,
        siteminder_files=args.siteminder_files,
    )
    write_outputs(output, summary, audit, args.output_dir)
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
