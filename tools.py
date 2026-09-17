import os
import json
import re
import zipfile
import pandas as pd

from .document_utils import (
    extract_word_text,
    extract_placeholders,
    generate_certificate,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
EXCEL_PATH = os.path.join(INPUT_DIR, "Book1.xlsx")
WORD_PATH = os.path.join(INPUT_DIR, "template.docx")
CERTIFICATE_DIR = os.path.join(OUTPUT_DIR, "certificates")


def ensure_directories():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(CERTIFICATE_DIR, exist_ok=True)


def safe_filename(name):
    name = re.sub(r"[^A-Za-z0-9 _-]", "", str(name))
    name = name.strip().replace(" ", "_")
    return name or "Participant"


def inspect_files() -> str:
    """Inspect the user's Excel and Word inputs before generation."""
    ensure_directories()

    if not os.path.exists(EXCEL_PATH):
        return json.dumps({
            "status": "ERROR",
            "message": f"Excel file not found: {EXCEL_PATH}",
        })

    if not os.path.exists(WORD_PATH):
        return json.dumps({
            "status": "ERROR",
            "message": f"Word template not found: {WORD_PATH}",
        })

    df = pd.read_excel(EXCEL_PATH).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    word_text = extract_word_text(WORD_PATH)
    placeholders = extract_placeholders(WORD_PATH)

    return json.dumps({
        "status": "SUCCESS",
        "excel": {
            "file": "input/Book1.xlsx",
            "records": len(df),
            "columns": list(df.columns),
            "sample_records": df.head(5).fillna("").to_dict(orient="records"),
            "missing_values": df.isnull().sum().to_dict(),
            "duplicates": int(df.duplicated().sum()),
        },
        "word": {
            "file": "input/template.docx",
            "placeholders": placeholders,
            "placeholder_count": len(placeholders),
            "text_preview": word_text[:3000],
        },
    }, indent=2, default=str)


def generate_certificates(mapping_json: str) -> str:
    """
    Generate one personalized DOCX for each Excel row.
    mapping_json maps Word placeholder names to Excel column names.
    Example: {"NAME": "NAME"}
    """
    ensure_directories()

    if not os.path.exists(EXCEL_PATH):
        return json.dumps({
            "status": "ERROR",
            "message": "input/Book1.xlsx not found.",
        })

    if not os.path.exists(WORD_PATH):
        return json.dumps({
            "status": "ERROR",
            "message": "input/template.docx not found.",
        })

    try:
        mapping = json.loads(mapping_json)
        if not isinstance(mapping, dict):
            raise ValueError("Mapping must be a JSON object.")
    except Exception as e:
        return json.dumps({
            "status": "ERROR",
            "message": f"Invalid mapping JSON: {e}",
        })

    df = pd.read_excel(EXCEL_PATH).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    valid_columns = set(df.columns)
    invalid = [
        {"placeholder": p, "column": c}
        for p, c in mapping.items()
        if c is not None and c not in valid_columns
    ]

    if invalid:
        return json.dumps({
            "status": "ERROR",
            "message": "Mapping contains invalid Excel columns.",
            "invalid": invalid,
            "valid_columns": list(df.columns),
        }, indent=2)

    for filename in os.listdir(CERTIFICATE_DIR):
        path = os.path.join(CERTIFICATE_DIR, filename)
        if os.path.isfile(path):
            os.remove(path)

    results = []

    for index, row in df.iterrows():
        record = row.to_dict()

        name_column = mapping.get("NAME")
        recipient_name = str(record.get(name_column, "")) if name_column else ""
        if not recipient_name.strip():
            recipient_name = f"Participant_{index + 1}"

        filename = safe_filename(recipient_name) + ".docx"
        output_path = os.path.join(CERTIFICATE_DIR, filename)

        counter = 1
        while os.path.exists(output_path):
            filename = f"{safe_filename(recipient_name)}_{counter}.docx"
            output_path = os.path.join(CERTIFICATE_DIR, filename)
            counter += 1

        try:
            generate_certificate(WORD_PATH, output_path, record, mapping)
            results.append({
                "record": index + 1,
                "recipient": recipient_name,
                "file": output_path,
                "status": "SUCCESS",
            })
        except Exception as e:
            results.append({
                "record": index + 1,
                "recipient": recipient_name,
                "file": "",
                "status": "FAILED",
                "error": str(e),
            })

    report_path = os.path.join(OUTPUT_DIR, "generation_report.csv")
    pd.DataFrame(results).to_csv(report_path, index=False)

    zip_path = os.path.join(OUTPUT_DIR, "certificates.zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for filename in os.listdir(CERTIFICATE_DIR):
            path = os.path.join(CERTIFICATE_DIR, filename)
            if os.path.isfile(path):
                z.write(path, filename)

    success = sum(x["status"] == "SUCCESS" for x in results)
    failed = sum(x["status"] == "FAILED" for x in results)

    return json.dumps({
        "status": "SUCCESS",
        "total_records": len(results),
        "generated": success,
        "failed": failed,
        "certificate_folder": CERTIFICATE_DIR,
        "zip_file": zip_path,
        "report_file": report_path,
        "results": results,
    }, indent=2)
