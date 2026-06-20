# -*- coding: utf-8 -*-
"""
PDF Document Classifier with Local OCR (PaddleOCR)

Scans a folder of unclassified PDFs, runs OCR on each one, classifies the
document by type using rule-based pattern matching, extracts key metadata
(dates, IDs, plate numbers, document numbers), renames the file according to
a per-type naming convention, and writes a summary report to an Excel file.

Designed to run safely on Windows under Anaconda by guarding against the
multiprocessing/math-library conflict that can freeze local OCR.
"""

import os

# Prevent the OCR math libraries (OpenMP) from breaking multiprocessing in a
# local Windows environment. Must be set before importing numpy/OCR libs.
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import re
import shutil
import logging
import multiprocessing
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import numpy as np
import pandas as pd
from pdf2image import convert_from_path

# Optional progress bar
try:
    from tqdm import tqdm
except ImportError:
    def tqdm(x, **kwargs):
        return x


# =======================================================
# 0. CONFIGURATION AND LOGGING
# =======================================================
class PDFConfig:
    # --- Base directory ---
    # Override via the PDF_CLASSIFIER_BASE_DIR environment variable,
    # otherwise defaults to a `workspace` folder next to this project.
    BASE_DIR = Path(
        os.environ.get(
            "PDF_CLASSIFIER_BASE_DIR",
            Path(__file__).resolve().parent.parent / "workspace",
        )
    )

    INPUT_FOLDER = BASE_DIR / "unclassified"
    CLASSIFIED_FOLDER = BASE_DIR / "classified"
    UNCLASSIFIED_FOLDER = BASE_DIR / "not_classified"
    REPORTS_FOLDER = BASE_DIR / "reports"
    LOGS_FOLDER = BASE_DIR / "logs"

    REPORT_FILENAME = "classification_report.xlsx"
    LOG_FILENAME = "classification_process.log"

    # Path to Poppler binaries (used by pdf2image to convert PDF -> image).
    # On Windows under Anaconda this is typically the Library/bin folder.
    # Override via the POPPLER_PATH environment variable. None lets pdf2image
    # fall back to Poppler on the system PATH (common on Linux/macOS).
    POPPLER_PATH = os.environ.get("POPPLER_PATH") or None

    OCR_DPI = 300
    MIN_TEXT_LENGTH = 10
    WORKERS = 2

    @classmethod
    def setup_logging(cls):
        cls.LOGS_FOLDER.mkdir(parents=True, exist_ok=True)
        log_path = cls.LOGS_FOLDER / cls.LOG_FILENAME
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        if logger.hasHandlers():
            logger.handlers.clear()
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s - [%(levelname)s] - %(processName)s - %(message)s"
            )
        )
        logger.addHandler(file_handler)
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)
        logger.addHandler(console_handler)
        return logger

    @classmethod
    def initialize(cls):
        cls.CLASSIFIED_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.UNCLASSIFIED_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.REPORTS_FOLDER.mkdir(parents=True, exist_ok=True)
        logger = cls.setup_logging()
        logger.info("--- PROCESS START (LOCAL PADDLEOCR, SAFE ENVIRONMENT) ---")
        return True


# =======================================================
# 1. UTILITIES AND SPECIFIC EXTRACTION
# =======================================================
def unique_name(destination: Path) -> Path:
    if not destination.exists():
        return destination
    i = 1
    new = destination
    while new.exists():
        new_name = f"{destination.stem}_{i}{destination.suffix}"
        new = destination.with_name(new_name)
        i += 1
    return new


def sanitize_name(name: str) -> str:
    name = name.replace("\n", "").replace("\r", "")
    safe_name = re.sub(r"[^\w\-\.]", "", name)
    return safe_name.strip()


def clean_ocr_noise(text: str) -> str:
    text = re.sub(r"[|!§@#¢¬°]", " ", text)
    text = re.sub(r"([^\w\s])\1+", r"\1", text)
    return text


def extract_id_number(text: str) -> str:
    """Extract a national ID number that follows the word 'cédula'."""
    match = re.search(r"c[eé]dula[\s:]*([\d\.]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).replace(".", "").strip()
    return "NOID"


def extract_person_name(text: str) -> str:
    """Extract a person's name that follows the word 'Nombre'."""
    match = re.search(
        r"Nombre[\s:]*([A-Za-zÑñ\s]+?)(?:\s+Tipo|\s+Regional|\n|$)",
        text,
        re.IGNORECASE,
    )
    if match:
        clean_name = re.sub(r"\s+", "", match.group(1).title())
        return clean_name
    return "NONAME"


# =======================================================
# 2. CLASSIFICATION RULES
# =======================================================
# Document types are matched by regex patterns within a search zone.
# "HEADER" searches only the first part of the text; "ALL" searches everything.
# Company- or vendor-specific keywords have been replaced with generic
# placeholders. Adjust these patterns to match your own document set.
CLASSIFICATION_RULES = {
    "MEMO": {
        "patterns": [
            r"(?:^|\n)\s*[MN][EÉF]MORAND[O0]",
            r"M\s*E\s*M\s*O\s*R\s*A\s*N\s*D\s*[O0]",
            r"MEMORANDO\s+[\d\-\s_]+",
        ],
        "name_format": "DATE_MEMO_NUMBER",
        "zone": "HEADER",
    },
    "COMMISSION_STATEMENT": {
        "patterns": [
            r"LIQUIDACION\s+MENSUAL\s+DE\s+COMISIONES",
        ],
        "name_format": "FMT_COMMISSION_STATEMENT",
        "zone": "HEADER",
    },
    "HOTEL_RESERVATION": {
        "patterns": [
            r"CONFIRMANDO\s+RESERVA",
            r"HABITACI[OÓ]N\s+SENCILLA",
        ],
        "name_format": "FMT_HOTEL_RESERVATION",
        "zone": "ALL",
    },
    "CIRCULAR": {
        "patterns": [r"CIRCULAR"],
        "name_format": "DATE_CIRCULAR_NUMBER",
        "zone": "HEADER",
    },
    "PROMISSORY_NOTE": {
        "patterns": [r"PAGAR[EÉ]"],
        "name_format": "DATE_PROMISSORY_NOTE_NUMBER",
        "zone": "HEADER",
    },
    "CREDIT_NOTE": {
        "patterns": [r"NOTA\s+(?:DE\s+)?CR[EÉ]DITO"],
        "name_format": "FMT_CREDIT_NOTE",
        "zone": "HEADER",
    },
    "PAYMENT_REQUEST": {
        "patterns": [r"CUENTA\s+DE\s+COBRO"],
        "name_format": "FMT_PAYMENT_REQUEST",
        "zone": "ALL",
    },
    "INVOICE": {
        "patterns": [r"(?:^|\n)\s*(?:FACTURA ELECTRONICA|FACTURA DE VENTA)"],
        "name_format": "DATE_INVOICE_NUMBER",
        "zone": "ALL",
    },
    "CONTRACT": {
        "patterns": [r"CONTRATO|TERMINACI[OÓ]N.*ACUERDO|COMPRAVENTA"],
        "name_format": "DATE_CONTRACT_NUMBER",
        "zone": "HEADER",
    },
}

# Content-based fallback rules: a category is chosen when at least two of its
# keywords appear anywhere in the text.
SMART_RULES = {
    "LEGAL": {
        "keywords": [
            "JUZGADO", "FALLO", "SENTENCIA", "DEMANDA",
            "TUTELA", "ABOGADO", "NOTARIA",
        ],
        "format": "DATE_LEGAL_NUMBER",
    },
    "TAXES": {
        "keywords": [
            "IMPUESTO", "RETEFUENTE", "RETEICA",
            "DECLARACION", "DIAN", "BIMESTRE",
        ],
        "format": "DATE_TAXES_NUMBER",
    },
}

MONTHS_DICT = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", "MAY": "05",
    "JUN": "06", "JUL": "07", "AGO": "08", "SEP": "09", "OCT": "10",
    "NOV": "11", "DIC": "12", "ENERO": "01", "FEBRERO": "02",
    "MARZO": "03", "ABRIL": "04", "MAYO": "05", "JUNIO": "06",
    "JULIO": "07", "AGOSTO": "08", "SEPTIEMBRE": "09",
}

FORMAT_MAP = {
    "FMT_COMMISSION_STATEMENT": lambda p: f"CS_{p['ID']}_{p['PERSON_NAME']}",
    "FMT_HOTEL_RESERVATION": lambda p: f"{p['DATE']}_HOTEL_RESERVATION_"
        f"{p['NUMBER'] if p['NUMBER'] != 'NO_NUMBER' else 'RES'}",
    "FMT_PAYMENT_REQUEST": lambda p: f"{p['DATE']}_PAYMENT_REQUEST_{p['NUMBER']}",
    "FMT_CREDIT_NOTE": lambda p: f"{p['DATE']}_CREDIT_NOTE_{p['NUMBER']}",
    "DATE_MEMO_NUMBER": lambda p: f"{p['DATE']}_MEMO_{p['NUMBER']}_"
        f"{p['PLATE'] if p['PLATE'] != 'NO_PLATE' else (p['UNIT'] if p['UNIT'] != 'NO_UNIT' else 'NO_PLATE')}",
    "DATE_CIRCULAR_NUMBER": lambda p: f"{p['DATE']}_CIRCULAR_{p['NUMBER']}",
    "DATE_PROMISSORY_NOTE_NUMBER": lambda p: f"{p['DATE']}_PROMISSORY_NOTE_{p['NUMBER']}",
    "DATE_INVOICE_NUMBER": lambda p: f"{p['DATE']}_INVOICE_{p['NUMBER']}",
    "DATE_CONTRACT_NUMBER": lambda p: f"{p['DATE']}_CONTRACT_{p['NUMBER']}",
    "DATE_LEGAL_NUMBER": lambda p: f"{p['DATE']}_LEGAL_{p['NUMBER']}",
    "DATE_TAXES_NUMBER": lambda p: f"{p['DATE']}_TAXES_{p['NUMBER']}",
    "DATE_GENERIC_DATA": lambda p: f"{p['DATE']}_GENERIC_{p['NUMBER']}_{p['PLATE']}",
    "ORIGINAL_FALLBACK": lambda p: "UNKNOWN",
}


def classify_document(text: str) -> tuple:
    text = clean_ocr_noise(text)
    text_upper = text.upper()
    text_header = text_upper[:2000]
    high_priority = [
        "MEMO", "COMMISSION_STATEMENT", "HOTEL_RESERVATION", "CIRCULAR",
        "PROMISSORY_NOTE", "CONTRACT", "INVOICE", "CREDIT_NOTE",
        "PAYMENT_REQUEST",
    ]
    for doc_type in high_priority:
        if doc_type in CLASSIFICATION_RULES:
            data = CLASSIFICATION_RULES[doc_type]
            zone = data.get("zone", "ALL")
            search_text = text_header if zone == "HEADER" else text_upper
            if all(re.search(pattern, search_text) for pattern in data["patterns"]):
                return doc_type, data["name_format"]
    return "UNKNOWN", "ORIGINAL_FALLBACK"


def classify_by_content(text: str) -> tuple:
    text_upper = text.upper()
    scores = {}
    for category, data in SMART_RULES.items():
        score = sum(1 for keyword in data["keywords"] if keyword in text_upper)
        if score >= 2:
            scores[category] = score
    if scores:
        best = max(scores, key=scores.get)
        return best, SMART_RULES[best]["format"]
    return None, None


# =======================================================
# 4. IMAGE PROCESSING WITH OCR (SAFE PADDLEOCR INSTANCE)
# =======================================================
def ocr_image_text(pdf_path: Path) -> str:
    from paddleocr import PaddleOCR

    # Instantiated inside the worker to avoid math-library collisions on Windows.
    ocr_model = PaddleOCR(use_angle_cls=True, lang="es")
    try:
        images = convert_from_path(
            pdf_path,
            PDFConfig.OCR_DPI,
            first_page=1,
            last_page=1,
            thread_count=1,
            poppler_path=PDFConfig.POPPLER_PATH,
        )
    except Exception:
        return ""

    full_text = ""
    for image in images:
        img_np = np.array(image)
        results = ocr_model.ocr(img_np, cls=True)
        if not results or not results[0]:
            continue
        for line in results[0]:
            detected_text = line[1][0]
            full_text += f"{detected_text}\n "
    return full_text.strip()


# --- GENERAL EXTRACTORS ---
def _normalize_year(year_str: str) -> str:
    if len(year_str) == 4:
        return year_str if int(year_str) >= 1980 else None
    val = int(year_str)
    return f"19{year_str}" if val > 80 else f"20{year_str}"


def extract_date(text: str) -> str:
    text = text.upper()
    text = re.sub(r"\bDEL\b", "DE", text)
    text = re.sub(r"[(\[”\"'°º]", "", text)
    text = re.sub(r"(\d{1,2})DE", r"\1 DE", text)
    text = re.sub(r"DE(\d{4})", r"DE \1", text)

    header = text[:1200]
    candidates = []

    for match in re.finditer(
        r"\b(\d{1,2})\s+(?:DE\s+)?([A-Z]{3,10})\.?\s+(?:DE\s+)?(\d{2,4})", header
    ):
        day, month, year = match.groups()
        month_str = MONTHS_DICT.get(month[:3], "00")
        if month_str != "00":
            year_norm = _normalize_year(year)
            if year_norm:
                return f"{year_norm}{month_str}{day.zfill(2)}"

    for match in re.finditer(r"\b([A-Z]{3,10})\s+(\d{1,2})\b", header):
        month, day = match.groups()
        month_str = MONTHS_DICT.get(month[:3], "00")
        match_year = re.search(r"\b(20\d{2}|19\d{2})\b", header)
        year = match_year.group(1) if match_year else "2007"
        if month_str != "00":
            return f"{year}{month_str}{day.zfill(2)}"

    for match in re.finditer(r"(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{2,4})", text):
        day, month, year = match.groups()
        year_norm = _normalize_year(year)
        if year_norm:
            candidates.append(f"{year_norm}{month.zfill(2)}{day.zfill(2)}")

    if candidates:
        return sorted(candidates, reverse=True)[0]
    return "NO_DATE"


def extract_plate(text: str) -> str:
    pattern = r"\b([A-Z]{3})\s*[-._]?\s*(\d{3})\b|\b([A-Z]{2})\s*[-._]?\s*(\d{4})\b"
    all_matches = re.findall(pattern, text, re.IGNORECASE)
    found = []
    for p in all_matches:
        raw = "".join(p)
        if any(x in raw.upper() for x in ["FAX", "PBX", "NIT", "INT"]):
            continue
        plate_clean = raw.replace("-", "").replace(" ", "").replace(".", "").upper()
        if re.match(r"^(19|20)\d{2}$", plate_clean):
            continue
        found.append(plate_clean)
    unique = sorted(set(found))
    return unique[0] if unique else "NO_PLATE"


def extract_document_number(text: str) -> str:
    text = text.upper()
    if re.search(r"MEMORANDO", text) or re.search(r"M\s*E\s*M\s*O", text):
        match_memo = re.search(
            r"(?:MEMORANDO|M\s*E\s*M\s*O\s*R\s*A\s*N\s*D\s*O)\s*(?:N[°º.]?)?\s*[-]?\s*(\d+)",
            text,
        )
        if match_memo:
            return match_memo.group(1)

    if "HOTEL" in text and "FAX" in text:
        match_fax = re.search(r"FAX\s*[:.]?\s*(\d{6,})", text)
        if match_fax:
            return match_fax.group(1)

    match_doc = re.search(
        r"(?:MEMORANDO|CIRCULAR|FAX)\s*(?:No\.?)?[\s-]*([0-9]{3})?[\s\-\/]*([0-9\s\-\/]+)",
        text,
    )
    if match_doc:
        prefix = match_doc.group(1) if match_doc.group(1) else ""
        cleaned = prefix + re.sub(r"[^0-9]", "", match_doc.group(2))
        if 3 <= len(cleaned) <= 12:
            return cleaned
    return "NO_NUMBER"


def extract_vehicle_unit(text: str) -> str:
    text = text.upper()
    match_block = re.search(
        r"(?:NO\.?|NUMERO)\s*INT(?:ERNO)?\.?[\s.:]+([\d\s-]+)", text
    )
    if match_block:
        numbers = re.findall(r"\b(\d{3,5})\b", match_block.group(1))
        if numbers:
            return f"UNIT_{numbers[0]}"
    return "NO_UNIT"


def extract_generic_name(text: str) -> str:
    match = re.search(r"NOMBRE[:\s]*([A-Z\s]+)", text.upper())
    return "_".join(match.group(1).strip().split()[0:3]) if match else "NO_NAME"


# =======================================================
# 5. MAIN WORKER
# =======================================================
def process_one_pdf(pdf_file: Path) -> dict:
    logger = logging.getLogger()
    res = {
        "Original_Name": pdf_file.name,
        "Classified_Type": "ERROR",
        "Extracted_Date": "N/A",
        "New_Name": f"ERROR_{pdf_file.name}",
        "Process_Status": "OCR_FAILED",
    }
    try:
        text = ocr_image_text(pdf_file)
        if len(text) >= PDFConfig.MIN_TEXT_LENGTH:
            doc_type, fmt = classify_document(text)
            if doc_type == "UNKNOWN":
                smart_type, smart_fmt = classify_by_content(text)
                if smart_type:
                    doc_type, fmt, res["Process_Status"] = (
                        smart_type, smart_fmt, "OK_SMART"
                    )

            params = {
                "DATE": extract_date(text),
                "PLATE": extract_plate(text),
                "NUMBER": extract_document_number(text),
                "UNIT": extract_vehicle_unit(text),
                "NAME": extract_generic_name(text),
                "ID": extract_id_number(text),
                "PERSON_NAME": extract_person_name(text),
            }

            if (
                doc_type == "UNKNOWN"
                and params["DATE"] != "NO_DATE"
                and (params["PLATE"] != "NO_PLATE" or params["NUMBER"] != "NO_NUMBER")
            ):
                doc_type, fmt, res["Process_Status"] = (
                    "GENERIC_DOCUMENT", "DATE_GENERIC_DATA", "OK_GENERIC"
                )

            if doc_type != "UNKNOWN":
                base = FORMAT_MAP.get(fmt, lambda p: "ERROR")(params)
                clean = sanitize_name(base)
                new_name = f"{clean}.pdf"
                final = unique_name(PDFConfig.CLASSIFIED_FOLDER / new_name)
                shutil.copy(pdf_file, final)
                res["New_Name"] = final.name
                res["Process_Status"] = "OK"
                logger.info(f"OK {pdf_file.name} -> {doc_type} ({new_name})")
            else:
                final = unique_name(
                    PDFConfig.UNCLASSIFIED_FOLDER / f"UNKNOWN_{pdf_file.name}"
                )
                shutil.copy(pdf_file, final)
                res["Process_Status"] = "CLASSIFICATION_FAILED"
                res["New_Name"] = final.name

            res.update({
                "Classified_Type": doc_type,
                "Extracted_Date": params["DATE"],
                "Extracted_Plate": params["PLATE"],
                "Document_Number": params["NUMBER"],
                "Extracted_ID": params["ID"],
                "Extracted_Name": params["PERSON_NAME"],
            })
        else:
            final = unique_name(
                PDFConfig.UNCLASSIFIED_FOLDER / f"OCR_FAILED_{pdf_file.name}"
            )
            shutil.copy(pdf_file, final)
            res["Process_Status"] = "OCR_EMPTY"
            res["New_Name"] = final.name

    except Exception as e:
        res["Process_Status"] = f"CRASH: {str(e)}"
        logger.error(f"CRASH {pdf_file.name}: {e}")
    return res


def main():
    if not PDFConfig.initialize():
        return

    input_path = PDFConfig.INPUT_FOLDER
    print(f"\nChecking input folder: {input_path.resolve()}")

    if not input_path.exists():
        print(f"ERROR: Input folder '{input_path.name}' does not exist at the given path.")
        return

    # Robust scan covering both .pdf and .PDF extensions
    files = list(input_path.glob("*.pdf")) + list(input_path.glob("*.PDF"))
    print(f"PDF files detected for processing: {len(files)}")

    if not files:
        print(f"No .pdf or .PDF files found in '{input_path.name}'.")
        return

    results = []
    with ProcessPoolExecutor(max_workers=PDFConfig.WORKERS) as executor:
        futures = {executor.submit(process_one_pdf, pdf): pdf for pdf in files}
        for future in tqdm(as_completed(futures), total=len(files), unit="pdf"):
            results.append(future.result())

    if results:
        df = pd.DataFrame(results)
        report_path = PDFConfig.REPORTS_FOLDER / PDFConfig.REPORT_FILENAME
        ordered_columns = [
            "Original_Name", "New_Name", "Classified_Type", "Extracted_Date",
            "Extracted_Plate", "Document_Number", "Extracted_ID",
            "Extracted_Name", "Process_Status",
        ]
        final_columns = [c for c in ordered_columns if c in df.columns]

        # Guard against the report being locked by an open Excel window
        try:
            df[final_columns].to_excel(report_path, index=False)
            print(f"\nREPORT SAVED TO: {report_path}")
        except PermissionError:
            backup_path = PDFConfig.REPORTS_FOLDER / "classification_report_BACKUP.xlsx"
            final_backup = unique_name(backup_path)
            df[final_columns].to_excel(final_backup, index=False)
            print(f"\nMain report locked by Excel. Saved backup to: {final_backup}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
