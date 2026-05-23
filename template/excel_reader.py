"""
Excel Reader — Reads birth records and returns clean dicts for portal entry.
"""
import pandas as pd
from config.settings import GENDER_MAP


def load_records(filepath: str, sheet_name: str = "Sheet1") -> list[dict]:
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    records = []
    for idx, row in df.iterrows():
        try:
            records.append(_clean_row(idx, row))
        except Exception as e:
            print(f"  Warning Row {idx}: Skipped - {e}")
    print(f"  Loaded {len(records)} records from '{filepath}'")
    return records


def _clean_row(idx: int, row: pd.Series) -> dict:
    reg_number = _fmt_reg_number(row["रजि क्रमांक"])
    year = str(int(row["Unnamed: 1"])).strip()
    birth_date = _fmt_date(row["जन्म दिनांक"])
    reg_date = _fmt_date(row["रजि दिनांक"])
    gender_raw = str(row["लिंग"]).strip()
    gender = GENDER_MAP.get(gender_raw, gender_raw)
    child_name = str(row["बालक का नाम"]).strip()
    father_name = str(row["पिता का नाम"]).strip()
    mother_name = str(row["माता का नाम"]).strip()
    religion = str(row["हिन्दू मुस्लिम"]).strip()
    address = str(row["पता"]).strip()
    if address.lower() == "nan":
        address = str(row["पता जन्म स्थान"]).strip()
    return {
        "excel_row": idx,
        "reg_number": reg_number,
        "year": year,
        "birth_date": birth_date,
        "reg_date": reg_date,
        "gender": gender,
        "child_name": child_name,
        "father_name": father_name,
        "mother_name": mother_name,
        "religion": religion,
        "address": address,
    }


def _fmt_reg_number(val) -> str:
    if isinstance(val, (int, float)) and not pd.isna(val):
        return str(int(val)).strip()
    s = str(val).strip()
    if s.replace(".", "", 1).isdigit():
        return str(int(float(s)))
    return s


def _fmt_date(dt) -> str:
    if isinstance(dt, (pd.Timestamp,)):
        return dt.strftime("%d/%m/%Y")
    return str(dt).strip()


def get_record_summary(r: dict) -> str:
    return f"Row {r['excel_row']}: Reg#{r['reg_number']}/{r['year']} | {r['child_name']} | {r['father_name']} | {r['gender']}"
