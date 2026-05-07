#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
import zipfile
from collections import defaultdict
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
TABLES_DIR = ROOT / "tables"
FINAL_INPUT_DIR = TABLES_DIR / "FinalInput"
NEW_DIR = TABLES_DIR / "new"
OUTPUT_DIR = ROOT / "outputs" / "rms_import"

MASTER_FILE = FINAL_INPUT_DIR / "main.xls"
ARRIVAL_FILE = FINAL_INPUT_DIR / "Arrival.xls"
EXPEDIA_FILE = FINAL_INPUT_DIR / "EXP.csv"
BATCH_FOLIO_FILE = FINAL_INPUT_DIR / "Batch Folio.xls"
ASI_BOOKING_REPORT_FILE = TABLES_DIR / "YEHS Hotel Sydney QVB Booking Report.xlsx"


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


def clean_line(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def normalize_booking_id(value: Any) -> str:
    text = clean_line(value)
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


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


def default_siteminder_files() -> list[Path]:
    final_input_siteminder = FINAL_INPUT_DIR / "SiteMinder.csv"
    if final_input_siteminder.exists():
        return [final_input_siteminder]
    return sorted(TABLES_DIR.glob("reservations_summary_report*.csv"))


def read_batch_folio(path: Path | None = None) -> pd.DataFrame:
    if path is None:
        return pd.DataFrame()
    if not path.exists():
        raise FileNotFoundError(f"Batch Folio file not found: {path}")

    raw = pd.read_excel(path, header=None, engine="xlrd", dtype=str).fillna("")
    header_row_index: int | None = None
    header_values: list[str] = []
    for idx, row in raw.iterrows():
        values = [clean_line(value) for value in row.tolist()]
        normalized = [normalize_source(value) for value in values]
        if "folio" in normalized or "folio number" in normalized:
            header_row_index = idx
            header_values = values
            break
    if header_row_index is None:
        raise ValueError("Batch Folio file missing Folio # header")

    normalized_headers = [normalize_source(value) for value in header_values]

    def header_col(*names: str) -> int | None:
        for name in names:
            normalized_name = normalize_source(name)
            if normalized_name in normalized_headers:
                return normalized_headers.index(normalized_name)
        return None

    folio_col = header_col("Folio #", "Folio No.", "Folio Number", "Folio")
    if folio_col is None:
        raise ValueError("Batch Folio file missing Folio # column")

    room_col = header_col("Room")
    first_col = header_col("First Name")
    last_col = header_col("Last Name")
    date_in_col = header_col("Date In")
    date_out_col = header_col("Date Out")
    source_col = header_col("Business Source")

    rows: list[dict[str, Any]] = []
    for idx, row in raw.iloc[header_row_index + 1 :].iterrows():
        folio = normalize_booking_id(row.get(folio_col, ""))
        if not folio or normalize_source(folio) in {"folio", "folio number"}:
            continue
        rows.append(
            {
                "Batch Folio Source Row": idx + 1,
                "Batch Folio #": folio,
                "Batch Folio Room": clean_line(row.get(room_col, "")) if room_col is not None else "",
                "Batch Folio First Name": clean_line(row.get(first_col, "")) if first_col is not None else "",
                "Batch Folio Last Name": clean_line(row.get(last_col, "")) if last_col is not None else "",
                "Batch Folio Date In": iso_date(as_date(row.get(date_in_col, ""))) if date_in_col is not None else "",
                "Batch Folio Date Out": iso_date(as_date(row.get(date_out_col, ""))) if date_out_col is not None else "",
                "Batch Folio Business Source": clean_line(row.get(source_col, "")) if source_col is not None else "",
            }
        )

    batch = pd.DataFrame(rows)
    if batch.empty:
        raise ValueError("Batch Folio file did not contain any usable Folio # rows")
    duplicates = batch["Batch Folio #"].value_counts()
    duplicated = duplicates[duplicates > 1]
    if not duplicated.empty:
        sample = ", ".join(duplicated.head(10).index.tolist())
        raise ValueError(f"Batch Folio contains duplicate Folio # values: {sample}")
    return batch


def read_siteminder(paths: list[Path] | None = None) -> pd.DataFrame:
    frames = []
    source_paths = paths if paths is not None else default_siteminder_files()
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
                "SM Active/Cancel": "Active" if non_cancelled else "Cancel",
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


def xlsx_column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    index = 0
    for char in match.group(1):
        index = index * 26 + ord(char) - 64
    return index - 1


def read_xlsx_first_sheet_values(path: Path) -> list[list[str]]:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"a": main_ns, "r": rel_ns, "pr": pkg_rel_ns}

    with zipfile.ZipFile(path) as workbook:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("a:si", ns):
                shared_strings.append("".join(node.text or "" for node in item.findall(".//a:t", ns)))

        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in workbook.namelist() and "xl/_rels/workbook.xml.rels" in workbook.namelist():
            workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
            rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
            first_sheet = workbook_root.find(".//a:sheets/a:sheet", ns)
            if first_sheet is not None:
                rel_id = first_sheet.attrib.get(f"{{{rel_ns}}}id", "")
                for rel in rels_root.findall("pr:Relationship", ns):
                    if rel.attrib.get("Id") == rel_id:
                        target = rel.attrib.get("Target", "worksheets/sheet1.xml")
                        sheet_path = target.lstrip("/")
                        if not sheet_path.startswith("xl/"):
                            sheet_path = f"xl/{sheet_path}"
                        break

        sheet_root = ET.fromstring(workbook.read(sheet_path))
        rows: list[list[str]] = []
        for row in sheet_root.findall(".//a:sheetData/a:row", ns):
            values: list[str] = []
            for cell in row.findall("a:c", ns):
                col_idx = xlsx_column_index(cell.attrib.get("r", ""))
                while len(values) <= col_idx:
                    values.append("")
                cell_type = cell.attrib.get("t", "")
                if cell_type == "inlineStr":
                    value = "".join(node.text or "" for node in cell.findall(".//a:t", ns))
                else:
                    node = cell.find("a:v", ns)
                    raw_value = node.text if node is not None and node.text is not None else ""
                    if cell_type == "s" and raw_value.isdigit():
                        value = shared_strings[int(raw_value)] if int(raw_value) < len(shared_strings) else raw_value
                    else:
                        value = raw_value
                values[col_idx] = clean_line(value)
            while values and values[-1] == "":
                values.pop()
            rows.append(values)
        return rows


def find_header_index(rows: list[list[str]], required_header: str) -> int:
    required_norm = normalize_source(required_header)
    for index, row in enumerate(rows):
        if any(normalize_source(value) == required_norm for value in row):
            return index
    raise ValueError(f"ASI Booking Report missing required header: {required_header}")


def make_unique_headers(headers: list[str]) -> list[str]:
    used: dict[str, int] = {}
    unique = []
    for idx, header in enumerate(headers, start=1):
        name = clean_line(header) or f"Column {idx}"
        count = used.get(name, 0)
        used[name] = count + 1
        unique.append(name if count == 0 else f"{name} {count + 1}")
    return unique


def table_from_header_rows(rows: list[list[Any]], required_header: str) -> pd.DataFrame:
    clean_rows = [[clean_line(value) for value in row] for row in rows]
    header_idx = find_header_index(clean_rows, required_header)
    max_cols = max(len(row) for row in clean_rows[header_idx:]) if clean_rows[header_idx:] else 0
    headers = make_unique_headers(clean_rows[header_idx] + [""] * (max_cols - len(clean_rows[header_idx])))
    data_rows = []
    for row in clean_rows[header_idx + 1 :]:
        padded = row + [""] * (max_cols - len(row))
        data_rows.append(padded[:max_cols])
    return pd.DataFrame(data_rows, columns=headers)


def is_cancel_value(value: Any) -> bool:
    return "cancel" in normalize_source(clean_line(value))


def asi_status_column_score(series: pd.Series) -> int:
    values = [clean_line(value) for value in series.tolist() if clean_line(value)]
    if not values:
        return 0
    simple_statuses = {"-", "active", "booked", "confirmed", "modified"}
    simple_count = 0
    cancel_count = 0
    for value in values:
        normalized = normalize_source(value)
        if is_cancel_value(value):
            cancel_count += 1
            simple_count += 1
        elif value == "-" or normalized in simple_statuses:
            simple_count += 1
    if simple_count / len(values) < 0.8:
        return 0
    return simple_count + (cancel_count * 1000)


def read_asi_booking_report(path: Path | None = None) -> pd.DataFrame:
    if path is None:
        return pd.DataFrame()
    if not path.exists():
        raise FileNotFoundError(f"ASI Booking Report not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(path, dtype=str, keep_default_na=False)
    elif suffix == ".xlsx":
        raw = table_from_header_rows(read_xlsx_first_sheet_values(path), "BookingID")
    else:
        raw_rows = pd.read_excel(path, header=None, dtype=str, keep_default_na=False).fillna("").values.tolist()
        raw = table_from_header_rows(raw_rows, "BookingID")

    if raw.empty:
        return pd.DataFrame()

    raw = raw.copy().fillna("")
    normalized_columns = {normalize_source(col): col for col in raw.columns}
    booking_col = normalized_columns.get("bookingid") or normalized_columns.get("booking id")
    if not booking_col:
        raise ValueError("ASI Booking Report missing BookingID column")

    status_col = (
        normalized_columns.get("booking status")
        or normalized_columns.get("status")
        or normalized_columns.get("active cancel")
    )
    if not status_col:
        cancel_counts = {
            col: asi_status_column_score(raw[col])
            for col in raw.columns
            if col != booking_col
        }
        status_col = max(cancel_counts, key=cancel_counts.get) if cancel_counts else ""
        if not status_col or cancel_counts.get(status_col, 0) == 0:
            raise ValueError("ASI Booking Report missing a status/cancel column")

    report = raw[[booking_col, status_col]].copy()
    report.columns = ["ASI Booking Report Booking ID", "ASI Booking Report Status"]
    report["ASI Booking Report Booking ID"] = report["ASI Booking Report Booking ID"].map(normalize_booking_id)
    report["ASI Booking Report Status"] = report["ASI Booking Report Status"].map(clean_line)
    report = report[report["ASI Booking Report Booking ID"].ne("")].copy()
    if report.empty:
        return report

    grouped_rows = []
    for booking_id, group in report.groupby("ASI Booking Report Booking ID", dropna=False):
        statuses = [clean_line(value) for value in group["ASI Booking Report Status"].tolist() if clean_line(value)]
        grouped_rows.append(
            {
                "ASI Booking Report Booking ID": booking_id,
                "ASI Booking Report Statuses": "|".join(sorted(set(statuses))),
                "ASI Booking Report Active/Cancel": "Cancel" if any(is_cancel_value(value) for value in statuses) else "",
                "ASI Booking Report Row Count": int(len(group)),
            }
        )
    return pd.DataFrame(grouped_rows)


def normalize_source(source: str) -> str:
    text = clean_line(source).lower()
    text = text.replace("&", "and")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_name(name: str) -> str:
    text = clean_line(name).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def name_similarity(left: str, right: str) -> int:
    left_norm = normalize_name(left)
    right_norm = normalize_name(right)
    if not left_norm or not right_norm:
        return 0
    if left_norm == right_norm:
        return 100
    left_tokens = set(left_norm.split())
    right_tokens = set(right_norm.split())
    if len(left_tokens) >= 2 and left_tokens.issubset(right_tokens):
        return 95
    if len(right_tokens) >= 2 and right_tokens.issubset(left_tokens):
        return 95
    return int(round(SequenceMatcher(None, left_norm, right_norm).ratio() * 100))


def payment_method(row: pd.Series) -> str:
    source = normalize_source(row.get("Business Source", ""))
    expedia_payment_type = clean_line(row.get("Expedia Payment Type", ""))

    prepaid_sources = {"agoda", "ctrip", "airbnbxml", "airbnb"}
    vcc_sources = {
        "hotelbeds",
        "hopper",
        "jetstar hooroo qantas",
        "qantas jetstar hotels",
        "qantas jetstar hotels holidays",
        "qantas jetstar hotels holidays new",
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


def canonical_source_from_siteminder(row: pd.Series) -> str:
    raw = clean_line(row.get("SM Channels", "")) or clean_line(row.get("SM Affiliated Channels", ""))
    text = normalize_source(
        " ".join(
            [
                clean_line(row.get("SM Channels", "")),
                clean_line(row.get("SM Affiliated Channels", "")),
            ]
        )
    )
    if not text:
        return ""
    if any(token in text for token in ["ctrip", "trip com"]):
        return "Ctrip"
    if "booking com" in text:
        return "Booking.com"
    if "agoda" in text:
        return "Agoda"
    if "expedia" in text:
        return "Expedia"
    if "hotelbeds" in text:
        return "Hotelbeds"
    if "airbnb" in text:
        return "AirBnBXML"
    if "traveloka" in text:
        return "Traveloka"
    if "hopper" in text:
        return "Hopper"
    if "webbeds" in text or "destinations of the world" in text:
        return "WebBeds - Destinations of the World"
    if "qantas" in text or "jetstar" in text or "hooroo" in text:
        return "Jetstar / Hooroo / Qantas"
    return raw


def siteminder_match_key(row: pd.Series) -> str:
    crs_folio = clean_line(row.get("CRS Folio #", ""))
    if normalize_source(row.get("Business Source", "")) == "hotelbeds" and "-" in crs_folio:
        return crs_folio.split("-", 1)[1].strip()
    return crs_folio


def add_arrival_candidate_fields(merged: pd.DataFrame, arrival_raw: pd.DataFrame) -> pd.DataFrame:
    merged = merged.copy()
    defaults = {
        "Arrival Same Date Candidate Count": 0,
        "Arrival Similar Candidate Name": "",
        "Arrival Similar Candidate Score": 0,
        "Arrival Similar Candidate Row": "",
        "Arrival Similar Candidate Room": "",
        "Arrival Similar Candidate Remark": "",
    }
    if arrival_raw.empty:
        for col, value in defaults.items():
            merged[col] = value
        return merged

    by_dates: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for rec in arrival_raw.to_dict("records"):
        by_dates[(rec["arrival_checkin_date"], rec["arrival_checkout_date"])].append(rec)

    candidate_rows = []
    for rec in merged.to_dict("records"):
        same_date = by_dates.get((rec.get("_checkin_date", ""), rec.get("_checkout_date", "")), [])
        guest_name = rec.get("_master_guest_name_exact", "")
        best: dict[str, Any] | None = None
        best_score = 0
        for candidate in same_date:
            score = name_similarity(guest_name, candidate.get("arrival_guest_name_exact", ""))
            if score > best_score:
                best_score = score
                best = candidate
        candidate_rows.append(
            {
                "Arrival Same Date Candidate Count": len(same_date),
                "Arrival Similar Candidate Name": best.get("arrival_guest_name_exact", "") if best else "",
                "Arrival Similar Candidate Score": best_score,
                "Arrival Similar Candidate Row": best.get("arrival_source_row", "") if best else "",
                "Arrival Similar Candidate Room": best.get("arrival_room", "") if best else "",
                "Arrival Similar Candidate Remark": best.get("arrival_remark", "") if best else "",
            }
        )
    candidate_df = pd.DataFrame(candidate_rows, index=merged.index)
    for col in candidate_df.columns:
        merged[col] = candidate_df[col]
    return merged


def build_master(
    master_file: Path = MASTER_FILE,
    arrival_file: Path = ARRIVAL_FILE,
    expedia_file: Path = EXPEDIA_FILE,
    siteminder_files: list[Path] | None = None,
    asi_booking_report_file: Path | None = None,
    batch_folio_file: Path | None = BATCH_FOLIO_FILE,
) -> tuple[pd.DataFrame, dict[str, Any], pd.DataFrame]:
    master_all = read_absolute_master(master_file)
    master_columns = list(master_all.attrs["master_columns"])
    master_original_rows = len(master_all)

    batch_folio = read_batch_folio(batch_folio_file)
    batch_folio_enabled = not batch_folio.empty
    batch_folio_unmatched_count = 0
    if batch_folio_enabled:
        master_all["_folio_match_key"] = master_all["Folio No."].map(normalize_booking_id)
        batch_folio_numbers = set(batch_folio["Batch Folio #"].map(normalize_booking_id))
        master_folio_numbers = set(master_all["_folio_match_key"])
        batch_folio_unmatched_count = len(batch_folio_numbers - master_folio_numbers)
        master = master_all.copy()
        master = master.merge(
            batch_folio,
            how="left",
            left_on="_folio_match_key",
            right_on="Batch Folio #",
            validate="many_to_one",
        )
        master["Batch Folio Match Status"] = "not_in_batch_folio"
        master.loc[master["Batch Folio #"].fillna("").ne(""), "Batch Folio Match Status"] = "matched_active"
    else:
        master = master_all.copy()
        master["Batch Folio Match Status"] = "not_uploaded"

    arrival_raw, arrival_grouped = read_arrival_remarks(arrival_file)
    merged = master.merge(
        arrival_grouped,
        how="left",
        left_on=["_master_guest_name_exact", "_checkin_date", "_checkout_date"],
        right_on=["arrival_guest_name_exact", "arrival_checkin_date", "arrival_checkout_date"],
        validate="many_to_one",
    )
    merged = add_arrival_candidate_fields(merged, arrival_raw)

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
    else:
        merged["SM Booking reference"] = ""
        merged["SM Statuses"] = ""
        merged["SM Channels"] = ""
        merged["SM Affiliated Channels"] = ""
        merged["SM Active/Cancel"] = ""
        merged["SM Matched Row Count"] = ""

    expedia = read_expedia_payment_types(expedia_file)
    merged = merged.merge(
        expedia,
        how="left",
        left_on="CRS Folio #",
        right_on="Expedia Reservation ID",
        validate="many_to_one",
    )

    asi_booking_report = read_asi_booking_report(asi_booking_report_file)
    if not asi_booking_report.empty:
        merged["_asi_booking_report_match_key"] = merged["CRS Folio #"].map(normalize_booking_id)
        merged = merged.merge(
            asi_booking_report,
            how="left",
            left_on="_asi_booking_report_match_key",
            right_on="ASI Booking Report Booking ID",
            validate="many_to_one",
        )
    else:
        merged["ASI Booking Report Booking ID"] = ""
        merged["ASI Booking Report Statuses"] = ""
        merged["ASI Booking Report Active/Cancel"] = ""
        merged["ASI Booking Report Row Count"] = ""

    original_source_norm = merged["Business Source"].map(normalize_source)
    source_was_blank = merged["Business Source"].fillna("").map(clean_line).eq("")
    siteminder_source_fill = merged.apply(canonical_source_from_siteminder, axis=1)
    source_from_siteminder = source_was_blank & siteminder_source_fill.ne("")
    ctrip_from_siteminder = source_from_siteminder & siteminder_source_fill.eq("Ctrip")
    merged.loc[source_from_siteminder, "Business Source"] = siteminder_source_fill[source_from_siteminder]

    has_arrival = merged["ASI Arrival Match Count"].notna()
    has_similar_arrival = (~has_arrival) & merged["Arrival Similar Candidate Score"].fillna(0).astype(int).ge(88)
    has_arrival_signal = has_arrival | has_similar_arrival
    has_siteminder = merged.get("SM Booking reference", pd.Series("", index=merged.index)).fillna("").ne("")
    siteminder_active = has_siteminder & merged["SM Active/Cancel"].fillna("").eq("Active")
    siteminder_cancel = has_siteminder & merged["SM Active/Cancel"].fillna("").eq("Cancel")
    no_siteminder = ~has_siteminder
    if batch_folio_enabled:
        has_batch_folio = merged["Batch Folio #"].fillna("").ne("")
        missing_batch_folio = ~has_batch_folio
        cancel_conflict = pd.Series(False, index=merged.index)
        no_siteminder_exact_arrival = pd.Series(False, index=merged.index)
        no_siteminder_similar_arrival = pd.Series(False, index=merged.index)
        no_siteminder_default_active = pd.Series(False, index=merged.index)
        batch_active_siteminder_cancel_conflict = has_batch_folio & siteminder_cancel
        batch_missing_siteminder_active_conflict = missing_batch_folio & siteminder_active
        batch_folio_cancel = missing_batch_folio & ~siteminder_active
        merged["Active/Cancel"] = "Active"
        merged.loc[batch_folio_cancel, "Active/Cancel"] = "Cancel"
    else:
        has_batch_folio = pd.Series(False, index=merged.index)
        missing_batch_folio = pd.Series(False, index=merged.index)
        cancel_conflict = siteminder_cancel & has_arrival_signal
        no_siteminder_exact_arrival = no_siteminder & has_arrival
        no_siteminder_similar_arrival = no_siteminder & ~has_arrival & has_similar_arrival
        no_siteminder_default_active = no_siteminder & ~has_arrival_signal
        batch_active_siteminder_cancel_conflict = pd.Series(False, index=merged.index)
        batch_missing_siteminder_active_conflict = pd.Series(False, index=merged.index)
        batch_folio_cancel = pd.Series(False, index=merged.index)
        merged.loc[has_siteminder, "Active/Cancel"] = merged.loc[has_siteminder, "SM Active/Cancel"]
        merged.loc[cancel_conflict, "Active/Cancel"] = "Active"
        merged.loc[no_siteminder, "Active/Cancel"] = "Active"

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
        "matched_siteminder"
    )
    merged["Expedia Payment Match Status"] = ""
    is_expedia = merged["Business Source"].map(normalize_source).eq("expedia")
    merged.loc[is_expedia, "Expedia Payment Match Status"] = "not_matched"
    merged.loc[is_expedia & merged["Expedia Payment Type"].fillna("").ne(""), "Expedia Payment Match Status"] = "matched"
    merged["Business Source Fix Status"] = ""
    merged["Business Source Filled From SiteMinder"] = ""
    merged.loc[source_from_siteminder, "Business Source Fix Status"] = "blank_source_set_from_siteminder"
    merged.loc[source_from_siteminder, "Business Source Filled From SiteMinder"] = siteminder_source_fill[source_from_siteminder]
    if batch_folio_enabled:
        merged["Active/Cancel Source"] = "Batch Folio active"
        merged.loc[batch_folio_cancel, "Active/Cancel Source"] = "Batch Folio missing; treated as cancelled"
        merged.loc[batch_missing_siteminder_active_conflict, "Active/Cancel Source"] = (
            "Batch Folio missing but SiteMinder active"
        )
    else:
        merged["Active/Cancel Source"] = "SiteMinder Active"
        merged.loc[siteminder_cancel, "Active/Cancel Source"] = "SiteMinder Cancel"
        merged.loc[cancel_conflict & has_arrival, "Active/Cancel Source"] = "SiteMinder Cancel overridden by exact Arrival match"
        merged.loc[cancel_conflict & ~has_arrival & has_similar_arrival, "Active/Cancel Source"] = "SiteMinder Cancel overridden by similar Arrival candidate"
        merged.loc[no_siteminder_exact_arrival, "Active/Cancel Source"] = "No SiteMinder match; exact Arrival match"
        merged.loc[no_siteminder_similar_arrival, "Active/Cancel Source"] = "No SiteMinder match; similar Arrival candidate"
        merged.loc[no_siteminder_default_active, "Active/Cancel Source"] = "No SiteMinder match; conservative default Active"
    merged["Status Needs Review"] = "No"
    review_rows = (
        (batch_missing_siteminder_active_conflict | batch_active_siteminder_cancel_conflict)
        if batch_folio_enabled
        else (cancel_conflict | no_siteminder)
    )
    merged.loc[review_rows, "Status Needs Review"] = "Yes"
    merged["Status Review Reason"] = ""
    if batch_folio_enabled:
        merged.loc[batch_missing_siteminder_active_conflict, "Status Review Reason"] = (
            "Batch Folio does not contain this Folio No. but SiteMinder says Active; kept Active for manual review."
        )
        merged.loc[batch_active_siteminder_cancel_conflict, "Status Review Reason"] = (
            "Batch Folio contains this Folio No. but SiteMinder says Cancel; kept Active for manual review."
        )
    else:
        merged.loc[cancel_conflict, "Status Review Reason"] = "SiteMinder says Cancel but Arrival List has an exact or similar same-date candidate; defaulted to Active."
        merged.loc[no_siteminder_exact_arrival, "Status Review Reason"] = "No SiteMinder match; exact Arrival List match exists; defaulted to Active."
        merged.loc[no_siteminder_similar_arrival, "Status Review Reason"] = "No SiteMinder match; similar same-date Arrival candidate exists; defaulted to Active."
        merged.loc[no_siteminder_default_active, "Status Review Reason"] = "No SiteMinder match and no strong Arrival candidate; conservative default is Active."
    blank_name_review = review_rows & merged["_master_guest_name_exact"].fillna("").eq("")
    merged.loc[blank_name_review, "Status Review Reason"] = (
        merged.loc[blank_name_review, "Status Review Reason"] + " Master guest name is blank, so Arrival name matching is unreliable."
    )

    asi_booking_report_uploaded = asi_booking_report_file is not None
    has_asi_booking_report = merged["ASI Booking Report Booking ID"].fillna("").ne("")
    asi_booking_report_cancel = merged["ASI Booking Report Active/Cancel"].fillna("").eq("Cancel")
    merged["ASI Booking Report Match Status"] = "not_uploaded"
    if asi_booking_report_uploaded:
        merged["ASI Booking Report Match Status"] = "not_matched"
        merged.loc[has_asi_booking_report, "ASI Booking Report Match Status"] = "matched"
    merged["ASI Booking Report Cancel Override"] = ""
    merged.loc[asi_booking_report_cancel, "ASI Booking Report Cancel Override"] = (
        merged.loc[asi_booking_report_cancel, "Active/Cancel"].fillna("") + " -> Cancel"
    )
    merged.loc[asi_booking_report_cancel, "Active/Cancel"] = "Cancel"
    merged.loc[asi_booking_report_cancel, "Active/Cancel Source"] = "ASI Booking Report Cancel"
    merged.loc[asi_booking_report_cancel, "Status Needs Review"] = "No"
    merged.loc[asi_booking_report_cancel, "Status Review Reason"] = ""

    expected_enriched_rows = master_original_rows
    if len(merged) != expected_enriched_rows:
        raise RuntimeError(f"Enriched row count changed from {expected_enriched_rows} to {len(merged)}")
    if not merged["_Absolute Master Row #"].is_unique:
        raise RuntimeError("Internal absolute master row number is no longer unique")

    output = merged[master_columns + ["Payment Type", "Note", "Active/Cancel"]].copy()
    output = output.fillna("")

    if len(output) != master_original_rows:
        raise RuntimeError(f"Output row count changed from {master_original_rows} to {len(output)}")

    summary = {
        "master_file": str(master_file),
        "arrival_file": str(arrival_file),
        "expedia_file": str(expedia_file),
        "asi_booking_report_file": str(asi_booking_report_file) if asi_booking_report_file else "",
        "batch_folio_file": str(batch_folio_file) if batch_folio_file else "",
        "siteminder_files": [str(p) for p in (siteminder_files if siteminder_files is not None else default_siteminder_files())],
        "master_rows_input": master_original_rows,
        "master_rows_output": int(len(output)),
        "master_rows_after_batch_folio_filter": int(has_batch_folio.sum()) if batch_folio_enabled else master_original_rows,
        "batch_folio_rows": int(len(batch_folio)),
        "batch_folio_unmatched_folio_numbers": int(batch_folio_unmatched_count),
        "batch_folio_cancelled_rows": int(batch_folio_cancel.sum()) if batch_folio_enabled else 0,
        "arrival_rows_raw": int(len(arrival_raw)),
        "arrival_exact_name_date_rows_matched": int(has_arrival.sum()),
        "arrival_exact_name_date_rows_with_remark": int(
            (has_arrival & merged["ASI Arrival Remark"].fillna("").ne("")).sum()
        ),
        "siteminder_booking_refs": int(len(sm)),
        "asi_booking_report_rows": int(len(asi_booking_report)),
        "asi_booking_report_rows_matched": int(has_asi_booking_report.sum()),
        "asi_booking_report_cancelled_rows_matched": int(asi_booking_report_cancel.sum()),
        "hotelbeds_rows_with_trimmed_siteminder_match_key": int(
            (original_source_norm.eq("hotelbeds") & merged.get("_siteminder_match_key", pd.Series("", index=merged.index)).ne(merged["CRS Folio #"])).sum()
        ),
        "active_cancel_counts": output["Active/Cancel"].value_counts(dropna=False).to_dict(),
        "active_cancel_source_counts": merged["Active/Cancel Source"].value_counts(dropna=False).to_dict(),
        "active_cancel_blank_after_conservative_status": int(output["Active/Cancel"].eq("").sum()),
        "status_needs_review_rows": int(merged["Status Needs Review"].eq("Yes").sum()),
        "siteminder_cancel_overridden_by_arrival_signal": int(cancel_conflict.sum()),
        "batch_folio_active_but_siteminder_cancel_rows": int(batch_active_siteminder_cancel_conflict.sum()),
        "batch_folio_missing_but_siteminder_active_rows": int(batch_missing_siteminder_active_conflict.sum()),
        "batch_folio_cancelled_but_siteminder_active_rows": int(batch_missing_siteminder_active_conflict.sum()),
        "business_source_blank_filled_from_siteminder": int(source_from_siteminder.sum()),
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
    cancelled_csv_path = out_dir / "cancelled_bookings.csv"
    cancelled_xlsx_path = out_dir / "cancelled_bookings.xlsx"
    status_review_csv_path = out_dir / "status_review.csv"
    status_review_xlsx_path = out_dir / "status_review.xlsx"
    summary_path = out_dir / "rms_absolute_master_summary.json"
    blank_payment_path = out_dir / "audit_payment_type_blank.csv"
    expedia_unmatched_path = out_dir / "audit_expedia_payment_unmatched.csv"
    arrival_no_match_path = out_dir / "audit_arrival_remark_not_matched.csv"
    active_cancel_unmatched_path = out_dir / "audit_active_cancel_unmatched.csv"
    active_cancel_fallback_path = out_dir / "audit_active_cancel_arrival_fallback.csv"
    ctrip_source_fixed_path = out_dir / "audit_business_source_ctrip_fixed.csv"
    status_needs_review_path = out_dir / "audit_status_needs_review.csv"
    source_filled_path = out_dir / "audit_business_source_filled_from_siteminder.csv"
    asi_booking_report_cancelled_path = out_dir / "audit_asi_booking_report_cancelled.csv"
    batch_folio_siteminder_conflicts_path = out_dir / "audit_batch_folio_siteminder_conflicts.csv"

    active_output = output[output["Active/Cancel"].ne("Cancel")].copy()
    cancelled_output = output[output["Active/Cancel"].eq("Cancel")].copy()
    status_review = audit[audit["Status Needs Review"].eq("Yes")].copy()
    review_columns = [
        "_Absolute Master Row #",
        "First Name",
        "Last Name",
        "Business Source",
        "Folio No.",
        "CRS Folio #",
        "Date In",
        "Date Out",
        "Room Type",
        "Room",
        "Payment Type",
        "Active/Cancel",
        "Active/Cancel Source",
        "Status Review Reason",
        "Batch Folio Match Status",
        "SM Booking reference",
        "SM Active/Cancel",
        "SM Statuses",
        "SM Channels",
        "ASI Arrival Match Count",
        "ASI Arrival Rooms",
        "ASI Arrival Source Rows",
        "Arrival Similar Candidate Name",
        "Arrival Similar Candidate Score",
        "Arrival Similar Candidate Row",
        "Arrival Similar Candidate Room",
        "Arrival Similar Candidate Remark",
    ]
    status_review = status_review[[c for c in review_columns if c in status_review.columns]]

    summary["active_import_rows"] = int(len(active_output))
    summary["cancelled_booking_rows"] = int(len(cancelled_output))
    summary["total_rows_across_active_and_cancelled_files"] = int(len(active_output) + len(cancelled_output))
    summary["status_review_excel_rows"] = int(len(status_review))
    summary["active_import_payment_type_counts"] = active_output["Payment Type"].value_counts(dropna=False).to_dict()
    summary["payment_type_counts"] = summary["active_import_payment_type_counts"]
    summary["active_import_unmapped_payment_source_counts"] = active_output.loc[
        active_output["Payment Type"].eq(""), "Business Source"
    ].value_counts(dropna=False).to_dict()
    summary["business_source_unmapped_payment_counts"] = summary["active_import_unmapped_payment_source_counts"]

    active_output.to_csv(csv_path, index=False, quoting=csv.QUOTE_MINIMAL)
    cancelled_output.to_csv(cancelled_csv_path, index=False, quoting=csv.QUOTE_MINIMAL)
    status_review.to_csv(status_review_csv_path, index=False, quoting=csv.QUOTE_MINIMAL)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        active_output.to_excel(writer, index=False, sheet_name="RMS Import Master")
        ws = writer.book["RMS Import Master"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    with pd.ExcelWriter(cancelled_xlsx_path, engine="openpyxl") as writer:
        cancelled_output.to_excel(writer, index=False, sheet_name="Cancelled Bookings")
        ws = writer.book["Cancelled Bookings"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    with pd.ExcelWriter(status_review_xlsx_path, engine="openpyxl") as writer:
        status_review.to_excel(writer, index=False, sheet_name="Status Review")
        ws = writer.book["Status Review"]
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
    audit[audit["Active/Cancel Source"].str.contains("Arrival", na=False)].to_csv(active_cancel_fallback_path, index=False)
    audit[audit["Business Source Fix Status"].eq("blank_source_set_from_siteminder")].to_csv(
        source_filled_path, index=False
    )
    audit[audit["Business Source Filled From SiteMinder"].eq("Ctrip")].to_csv(
        ctrip_source_fixed_path, index=False
    )
    audit[audit["Status Needs Review"].eq("Yes")].to_csv(status_needs_review_path, index=False)
    audit[audit["ASI Booking Report Active/Cancel"].eq("Cancel")].to_csv(
        asi_booking_report_cancelled_path, index=False
    )
    audit[
        audit["Status Review Reason"].str.contains("Batch Folio", na=False)
        & audit["SM Active/Cancel"].fillna("").ne("")
    ].to_csv(batch_folio_siteminder_conflicts_path, index=False)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--master-file", type=Path, default=MASTER_FILE)
    parser.add_argument("--arrival-file", type=Path, default=ARRIVAL_FILE)
    parser.add_argument("--expedia-file", type=Path, default=EXPEDIA_FILE)
    parser.add_argument("--batch-folio-file", type=Path, default=BATCH_FOLIO_FILE)
    parser.add_argument("--no-batch-folio", action="store_true")
    parser.add_argument("--asi-booking-report-file", type=Path, default=None)
    parser.add_argument("--siteminder-file", type=Path, action="append", dest="siteminder_files")
    parser.add_argument("--no-siteminder", action="store_true")
    args = parser.parse_args()
    output, summary, audit = build_master(
        master_file=args.master_file,
        arrival_file=args.arrival_file,
        expedia_file=args.expedia_file,
        siteminder_files=[] if args.no_siteminder else args.siteminder_files,
        asi_booking_report_file=args.asi_booking_report_file,
        batch_folio_file=None if args.no_batch_folio else args.batch_folio_file,
    )
    write_outputs(output, summary, audit, args.output_dir)
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
