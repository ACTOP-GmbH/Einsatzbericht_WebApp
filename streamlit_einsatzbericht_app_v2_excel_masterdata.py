from __future__ import annotations

import datetime as dt
import math
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
try:
    import streamlit as st
except ModuleNotFoundError:  # pragma: no cover
    st = None


APP_TITLE = "Einsatzbericht Manager (MVP)"
TAETIGKEITEN_SHEET = "Tätigkeiten"
HILFS_SHEET = "Hilfstabelle"
KODIERUNG_SHEET = "Kodierung Joyson"
RELEVANTE_KODIERUNG_SHEET = "relevante Kodierung"

# Original Excel report template area (detail rows)
REPORT_DETAIL_START_ROW = 17
REPORT_DETAIL_END_ROW = 35
REPORT_DETAIL_COL_START = 1  # A
REPORT_DETAIL_COL_END = 8    # H
REPORT_ROWS_PER_PAGE = REPORT_DETAIL_END_ROW - REPORT_DETAIL_START_ROW + 1

TAET_COLS = [
    "Datum",
    "Projekt",
    "Zeit von",
    "Zeit bis",
    "Pause",
    "Stunden",
    "Zahl",
    "km",
    "Tätigkeit",
    "Kodierung",
    "Interne Projekte",
    "Info",
    "Abgerechnet",
    "eingetragen",
]

KEY_COLS_FOR_EMPTY_CHECK = [1, 2, 3, 4, 9, 12, 13, 14]  # ignore formula columns F/G


@dataclass
class WorkbookData:
    path: Path
    taetigkeiten_df: pd.DataFrame
    lookups: Dict[str, Any]


# ------------------------- parsing helpers -------------------------

def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _safe_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _normalize_yes_no(value: Any) -> str:
    s = _safe_str(value).strip().lower()
    if s in {"ja", "yes", "y", "true", "1"}:
        return "ja"
    if s in {"nein", "no", "n", "false", "0"}:
        return "nein"
    return ""


def _to_date(value: Any) -> Optional[dt.date]:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        # try ISO first, then common DE format
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(value, fmt).date()
            except ValueError:
                pass
    return None


def _to_time(value: Any) -> Optional[dt.time]:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, dt.time):
        return value.replace(microsecond=0)
    if isinstance(value, dt.timedelta):
        total_seconds = int(value.total_seconds())
        total_seconds %= 24 * 3600
        return dt.time(total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60)
    if isinstance(value, (int, float)) and not pd.isna(value):
        # Excel serial fraction of a day
        frac = float(value) % 1
        total_seconds = int(round(frac * 24 * 3600))
        total_seconds %= 24 * 3600
        return dt.time(total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return dt.datetime.strptime(s, fmt).time()
            except ValueError:
                pass
    return None


def _time_to_minutes(value: Any) -> int:
    t = _to_time(value)
    if t is None:
        return 0
    return t.hour * 60 + t.minute + round(t.second / 60)


def _minutes_to_time(minutes: int) -> Optional[dt.time]:
    if minutes <= 0:
        return None
    minutes = int(minutes)
    minutes = max(0, minutes)
    h = (minutes // 60) % 24
    m = minutes % 60
    return dt.time(hour=h, minute=m)


def _compute_hours_decimal(start: Any, end: Any, pause_minutes: int) -> Optional[float]:
    t_start = _to_time(start)
    t_end = _to_time(end)
    if t_start is None or t_end is None:
        return None
    d0 = dt.datetime.combine(dt.date(2000, 1, 1), t_start)
    d1 = dt.datetime.combine(dt.date(2000, 1, 1), t_end)
    if d1 < d0:
        d1 += dt.timedelta(days=1)
    delta = d1 - d0 - dt.timedelta(minutes=max(pause_minutes, 0))
    hours = delta.total_seconds() / 3600
    if hours < 0:
        return None
    return round(hours, 4)


def _format_time(value: Any) -> str:
    t = _to_time(value)
    if t is None:
        return ""
    return t.strftime("%H:%M")


def _format_date(value: Any) -> str:
    d = _to_date(value)
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


def _display_row_label(row: pd.Series) -> str:
    d = _format_date(row.get("Datum"))
    zv = _format_time(row.get("Zeit von"))
    zb = _format_time(row.get("Zeit bis"))
    projekt = _safe_str(row.get("Projekt"))
    typ = _safe_str(row.get("Tätigkeit"))
    info = _safe_str(row.get("Info"))[:40]
    return f"Zeile {int(row['_excel_row'])} | {d} | {projekt} | {zv}-{zb} | {typ} | {info}"


# ------------------------- workbook reading -------------------------

def _read_list_column(ws, col_idx: int, start_row: int = 2) -> List[str]:
    values: List[str] = []
    seen = set()
    for r in range(start_row, ws.max_row + 1):
        v = ws.cell(r, col_idx).value
        if _is_blank(v):
            continue
        s = _safe_str(v).strip()
        if s not in seen:
            seen.add(s)
            values.append(s)
    return values


def _load_lookups(wb: openpyxl.Workbook) -> Dict[str, Any]:
    lookups: Dict[str, Any] = {}

    # Hilfstabelle
    if HILFS_SHEET not in wb.sheetnames:
        raise ValueError(f"Arbeitsblatt '{HILFS_SHEET}' nicht gefunden.")
    h = wb[HILFS_SHEET]

    lookups["ja_nein"] = _read_list_column(h, 1)
    if not lookups["ja_nein"]:
        lookups["ja_nein"] = ["ja", "nein"]

    lookups["taetigkeit_typen"] = _read_list_column(h, 3)
    if not lookups["taetigkeit_typen"]:
        lookups["taetigkeit_typen"] = ["F", "R", "I", "K"]

    lookups["interne_projekte"] = _read_list_column(h, 5)
    lookups["projekte"] = _read_list_column(h, 7)

    projekt_infos: Dict[str, Dict[str, Any]] = {}
    for r in range(2, h.max_row + 1):
        projekt = h.cell(r, 7).value
        if _is_blank(projekt):
            continue
        p = _safe_str(projekt).strip()
        projekt_infos[p] = {
            "Projekt": p,
            "Kunde": h.cell(r, 8).value,
            "Straße": h.cell(r, 9).value,
            "Ort": h.cell(r, 10).value,
            "Ansprechpartner": h.cell(r, 11).value,
            "Projektadresse Standard": h.cell(r, 12).value,
            "Projektadresse Alternativ": h.cell(r, 13).value,
        }
    lookups["projekt_infos"] = projekt_infos

    # Kodierung Joyson mapping: Aufgabe -> Kodierung EB / Kodierung intern
    kodierung_map_eb: Dict[str, str] = {}
    kodierung_map_intern: Dict[str, str] = {}
    kodierungen_aufgaben: List[str] = []
    if KODIERUNG_SHEET in wb.sheetnames:
        k = wb[KODIERUNG_SHEET]
        seen = set()
        for r in range(2, k.max_row + 1):
            aufgabe = k.cell(r, 2).value  # B
            if _is_blank(aufgabe):
                continue
            aufgabe_s = _safe_str(aufgabe).strip()
            if aufgabe_s not in seen:
                kodierungen_aufgaben.append(aufgabe_s)
                seen.add(aufgabe_s)
            kod_intern = k.cell(r, 3).value  # C
            kod_eb = k.cell(r, 4).value  # D
            if not _is_blank(kod_eb):
                kodierung_map_eb[aufgabe_s] = _safe_str(kod_eb).strip()
            if not _is_blank(kod_intern):
                kodierung_map_intern[aufgabe_s] = _safe_str(kod_intern).strip()
    lookups["kodierung_aufgaben"] = kodierungen_aufgaben
    lookups["kodierung_map_eb"] = kodierung_map_eb
    lookups["kodierung_map_intern"] = kodierung_map_intern

    # Relevante Kodierungen: Quelle der Wahrheit = Marker in 'Kodierung Joyson' (Spalte A)
    # (Das Sheet 'relevante Kodierung' kann Formeln enthalten, deren Cache ohne Excel-Neuberechnung veraltet sein kann.)
    relevante: List[str] = []
    if KODIERUNG_SHEET in wb.sheetnames:
        k = wb[KODIERUNG_SHEET]
        for r in range(2, k.max_row + 1):
            marker = _safe_str(k.cell(r, 1).value).strip().lower()
            aufgabe = k.cell(r, 2).value
            if marker == "x" and not _is_blank(aufgabe):
                relevante.append(_safe_str(aufgabe).strip())
    if not relevante and RELEVANTE_KODIERUNG_SHEET in wb.sheetnames:
        rk = wb[RELEVANTE_KODIERUNG_SHEET]
        relevante = _read_list_column(rk, 1)
    lookups["relevante_kodierungen"] = list(dict.fromkeys(relevante))

    return lookups


def _read_taetigkeiten_df(wb: openpyxl.Workbook) -> pd.DataFrame:
    if TAETIGKEITEN_SHEET not in wb.sheetnames:
        raise ValueError(f"Arbeitsblatt '{TAETIGKEITEN_SHEET}' nicht gefunden.")
    ws = wb[TAETIGKEITEN_SHEET]

    rows: List[Dict[str, Any]] = []
    for r in range(2, ws.max_row + 1):
        raw = [ws.cell(r, c).value for c in range(1, 15)]
        is_empty = all(_is_blank(ws.cell(r, c).value) for c in KEY_COLS_FOR_EMPTY_CHECK)
        if is_empty:
            continue

        rec = dict(zip(TAET_COLS, raw))
        rec["_excel_row"] = r
        rec["Datum"] = _to_date(rec.get("Datum"))
        rec["Zeit von"] = _to_time(rec.get("Zeit von"))
        rec["Zeit bis"] = _to_time(rec.get("Zeit bis"))
        rec["Pause"] = _to_time(rec.get("Pause"))
        rec["Pause_Min"] = _time_to_minutes(rec.get("Pause"))

        zahl = rec.get("Zahl")
        if isinstance(zahl, (int, float)) and not (isinstance(zahl, float) and math.isnan(zahl)):
            rec["Zahl"] = float(zahl)
        else:
            rec["Zahl"] = _compute_hours_decimal(rec.get("Zeit von"), rec.get("Zeit bis"), rec.get("Pause_Min", 0))

        rec["Stunden_Anzeige"] = None
        if rec.get("Zahl") is not None:
            total_minutes = int(round(float(rec["Zahl"]) * 60))
            rec["Stunden_Anzeige"] = f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"

        rec["Abgerechnet"] = _normalize_yes_no(rec.get("Abgerechnet")) or _safe_str(rec.get("Abgerechnet"))
        rec["eingetragen"] = _normalize_yes_no(rec.get("eingetragen")) or _safe_str(rec.get("eingetragen"))

        rows.append(rec)

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=TAET_COLS + ["_excel_row", "Pause_Min", "Stunden_Anzeige"])
        return df

    sort_date = pd.to_datetime(df["Datum"], errors="coerce")
    sort_start = df["Zeit von"].apply(lambda x: _time_to_minutes(x) if x else -1)
    df = df.assign(_sort_date=sort_date, _sort_start=sort_start).sort_values(
        ["_sort_date", "_sort_start", "_excel_row"], ascending=[False, False, False]
    )
    df = df.drop(columns=["_sort_date", "_sort_start"]).reset_index(drop=True)
    return df



def _default_excel_candidates() -> List[Path]:
    script_dir = Path(__file__).resolve().parent
    cwd = Path.cwd()
    names = [
        "Tätigkeiten_Überblick.xlsx",
        "__Tätigkeiten_Überblick - Kopie.xlsx",
    ]
    cands: List[Path] = []
    for base in [script_dir, script_dir / "data", cwd, cwd / "data"]:
        for n in names:
            cands.append(base / n)
        for p in sorted(base.glob("*.xlsx")):
            cands.append(p)
    # Dev fallback for this environment
    cands.append(Path("/mnt/data/__Tätigkeiten_Überblick - Kopie.xlsx"))
    # de-duplicate preserving order
    out: List[Path] = []
    seen = set()
    for c in cands:
        key = str(c)
        if key not in seen:
            seen.add(key)
            out.append(c)
    return out


def _resolve_excel_path(path_str: str) -> Path:
    raw = (path_str or "").strip()
    if raw:
        p = Path(raw).expanduser()
        try:
            if p.exists():
                return p.resolve()
        except Exception:
            pass
        # Relative paths: prefer script dir, then cwd
        candidates = [
            Path(__file__).resolve().parent / p,
            Path.cwd() / p,
        ]
        for c in candidates:
            if c.exists():
                return c.resolve()
        # As last resort return normalized absolute for clean error
        return p.resolve()
    for c in _default_excel_candidates():
        if c.exists():
            return c.resolve()
    # default target in ./data for helpful message
    return (Path(__file__).resolve().parent / "data" / "Tätigkeiten_Überblick.xlsx").resolve()

def _store_uploaded_excel(uploaded_file) -> Path:
    """
    Speichert eine hochgeladene Excel-Datei als lokale Arbeitskopie und gibt den Pfad zurück.
    Wichtig: Für Excel-COM (PDF/Druck) brauchen wir eine echte Datei auf Disk.
    """
    if uploaded_file is None:
        raise ValueError("Keine Datei hochgeladen.")

    original_name = Path(getattr(uploaded_file, "name", "upload.xlsx")).name
    suffix = Path(original_name).suffix.lower()

    if suffix != ".xlsx":
        raise ValueError("Bitte eine .xlsx-Datei hochladen.")

    script_dir = Path(__file__).resolve().parent
    imports_dir = script_dir / "imports"
    imports_dir.mkdir(parents=True, exist_ok=True)

    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    stem = Path(original_name).stem
    target = imports_dir / f"{stem}_import_{ts}.xlsx"

    # uploaded_file ist ein Streamlit UploadedFile
    data = uploaded_file.getvalue()
    target.write_bytes(data)

    return target.resolve()

def _prepare_report_formula_in_excel_sheet(ws) -> None:
    # Fallback (legacy): repair original FILTER spill in A17 using Excel itself.
    # NOTE: The app now prefers direct page filling for robust multi-page export/print.
    try:
        ws.Range(f"A{REPORT_DETAIL_START_ROW}:H200").ClearContents()
    except Exception:
        ws.Range(f"A{REPORT_DETAIL_START_ROW}:H{REPORT_DETAIL_END_ROW}").ClearContents()
    # Correct project filter (column B = project / C12)
    formula_en = '=FILTER(Berechnung!C2:J444,Berechnung!B2:B444=Einsatzbericht!C12,"")'
    try:
        ws.Range(f"A{REPORT_DETAIL_START_ROW}").Formula2 = formula_en
    except Exception:
        # Fallback for localized Excel
        ws.Range(f"A{REPORT_DETAIL_START_ROW}").FormulaLocal = '=FILTER(Berechnung!C2:J444;Berechnung!B2:B444=Einsatzbericht!C12;"")'


def _clear_original_report_detail_area(ws) -> None:
    ws.Range(f"A{REPORT_DETAIL_START_ROW}:H{REPORT_DETAIL_END_ROW}").ClearContents()


def _report_row_to_excel_values(row: pd.Series) -> List[Any]:
    """Convert a web report row into values for template columns A:H.

    IMPORTANT (pywin32/Excel COM):
    Writing ``datetime.date`` / ``datetime.time`` directly may raise
    ``must be a pywintypes time object`` on some systems. We therefore write
    display strings for date/time columns and keep only the decimal hours numeric.
    """
    datum = _format_date(row.get("Datum")) or _safe_str(row.get("Datum"))
    beginn = _format_time(row.get("Beginn")) or _safe_str(row.get("Beginn"))
    ende = _format_time(row.get("Ende")) or _safe_str(row.get("Ende"))
    pause = _format_time(row.get("Pause")) or _safe_str(row.get("Pause"))

    zeit_h = row.get("Zeit (h)")
    if zeit_h is None or (isinstance(zeit_h, float) and math.isnan(zeit_h)):
        zeit_h = None
    else:
        try:
            zeit_h = float(zeit_h)
        except Exception:
            zeit_h = None

    art = _safe_str(row.get("Art"))
    kod_eb = row.get("Kodierung EB")
    if pd.isna(kod_eb):
        kod_eb = None
    else:
        kod_eb = _safe_str(kod_eb) or None
    leistung = _safe_str(row.get("Leistungsbeschreibung"))
    return [datum, beginn, ende, pause, zeit_h, art, kod_eb, leistung]


def _write_original_report_page(ws, page_df: pd.DataFrame, year: int, month: int, project: str) -> None:
    """
    Fill the ORIGINAL 'Einsatzbericht' template directly (A17:H35) for one page.
    This avoids fragile spill formulas and allows deterministic multi-page export.
    """
    ws.Range("K2").Value = int(year)
    ws.Range("K3").Value = int(month)
    ws.Range("C12").Value = str(project)

    _clear_original_report_detail_area(ws)

    if page_df is None or page_df.empty:
        return

    # Write rows directly into template area (A:H)
    page_df = page_df.reset_index(drop=True)
    max_rows = min(len(page_df), REPORT_ROWS_PER_PAGE)
    rows_2d = []
    for idx in range(max_rows):
        rows_2d.append(_report_row_to_excel_values(page_df.iloc[idx]))

    if rows_2d:
        start = REPORT_DETAIL_START_ROW
        end = REPORT_DETAIL_START_ROW + len(rows_2d) - 1
        ws.Range(f"A{start}:H{end}").Value = tuple(tuple(r) for r in rows_2d)


def _split_report_pages(report_df: Optional[pd.DataFrame]) -> List[pd.DataFrame]:
    if report_df is None:
        return []
    if report_df.empty:
        return [report_df.copy()]
    pages: List[pd.DataFrame] = []
    total = len(report_df)
    for start in range(0, total, REPORT_ROWS_PER_PAGE):
        pages.append(report_df.iloc[start:start + REPORT_ROWS_PER_PAGE].copy().reset_index(drop=True))
    return pages


def _excel_original_report_action(
    xlsx_path: Path,
    year: int,
    month: int,
    project: str,
    action: str,
    pdf_output_path: Optional[Path] = None,
    xlsx_output_path: Optional[Path] = None,
    report_df: Optional[pd.DataFrame] = None,
) -> Tuple[bool, str]:
    """
    action: 'pdf' | 'xlsx' | 'print' | 'open'
    Uses Microsoft Excel via pywin32. Preferably fills the original sheet template directly
    from report_df (robust, paginated). Falls back to repairing Excel FILTER if needed.
    """
    try:
        import pythoncom  # type: ignore
        import win32com.client as win32  # type: ignore
    except Exception as e:
        return False, f"Excel-Automation nicht verfügbar (pywin32 / Excel): {e}"

    if not Path(xlsx_path).exists():
        return False, f"Excel-Datei nicht gefunden: {xlsx_path}"

    excel = None
    wb = None
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = bool(action == "open")
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        # Optional optimizations (dürfen nicht fatal sein)
        try:
            excel.EnableEvents = False
        except Exception:
            pass

        manual_calc_enabled = False
        try:
            # Manche Excel-Dateien/Setups erlauben das nicht -> dann einfach weiter mit AutoCalc
            excel.Calculation = -4135  # xlCalculationManual
            manual_calc_enabled = True
        except Exception:
            manual_calc_enabled = False
        wb = excel.Workbooks.Open(str(Path(xlsx_path).resolve()), UpdateLinks=False, ReadOnly=False)
        ws = wb.Worksheets("Einsatzbericht")

        pages = _split_report_pages(report_df)
        if not pages:
            pages = [pd.DataFrame()]
        page_count = len(pages)

        def _recalc() -> None:
            # Direktes Befüllen des Formulars braucht meist nur Sheet/normalen Recalc,
            # kein FullRebuild (sehr langsam).
            try:
                ws.Calculate()
                return
            except Exception:
                pass
            try:
                excel.Calculate()
                return
            except Exception:
                pass
            # Nur als allerletzter Notnagel:
            try:
                excel.CalculateFull()
            except Exception:
                pass

        if action == "pdf":
            if pdf_output_path is None:
                pdf_output_path = Path(xlsx_path).with_name(f"Einsatzbericht_{project}_{year}-{int(month):02d}.pdf")
            pdf_output_path = Path(pdf_output_path)
            pdf_output_path.parent.mkdir(parents=True, exist_ok=True)

            exported_files: List[Path] = []
            for page_idx, page_df in enumerate(pages, start=1):
                if report_df is not None:
                    _write_original_report_page(ws, page_df, int(year), int(month), str(project))
                else:
                    ws.Range("K2").Value = int(year)
                    ws.Range("K3").Value = int(month)
                    ws.Range("C12").Value = str(project)
                    _prepare_report_formula_in_excel_sheet(ws)
                _recalc()

                if page_count > 1:
                    out_path = pdf_output_path.with_name(f"{pdf_output_path.stem}_{page_idx:02d}{pdf_output_path.suffix}")
                else:
                    out_path = pdf_output_path
                ws.ExportAsFixedFormat(0, str(out_path))  # 0 = xlTypePDF
                exported_files.append(out_path)

            wb.Close(SaveChanges=False)
            try:
                if manual_calc_enabled:
                    excel.Calculation = -4105  # xlCalculationAutomatic
            except Exception:
                pass
            try:
                excel.EnableEvents = True
            except Exception:
                pass
            excel.Quit()
            if page_count > 1:
                return True, f"{page_count} PDFs exportiert ({REPORT_ROWS_PER_PAGE} Positionen pro Formularseite): {exported_files[0].name} ... {exported_files[-1].name}"
            return True, f"PDF exportiert: {exported_files[0]}"
        if action == "xlsx":
            if xlsx_output_path is None:
                xlsx_output_path = Path(xlsx_path).with_name(f"Einsatzbericht_{project}_{year}-{int(month):02d}.xlsx")
            xlsx_output_path = Path(xlsx_output_path)
            xlsx_output_path.parent.mkdir(parents=True, exist_ok=True)

            exported_files: List[Path] = []
            for page_idx, page_df in enumerate(pages, start=1):
                if report_df is not None:
                    _write_original_report_page(ws, page_df, int(year), int(month), str(project))
                else:
                    ws.Range("K2").Value = int(year)
                    ws.Range("K3").Value = int(month)
                    ws.Range("C12").Value = str(project)
                    _prepare_report_formula_in_excel_sheet(ws)
                _recalc()

                if page_count > 1:
                    out_path = xlsx_output_path.with_name(f"{xlsx_output_path.stem}_{page_idx:02d}{xlsx_output_path.suffix}")
                else:
                    out_path = xlsx_output_path

                # Speichert eine vorbereitete Kopie des Workbooks (Original bleibt unverändert)
                wb.SaveCopyAs(str(out_path))
                exported_files.append(out_path)

            wb.Close(SaveChanges=False)
            try:
                if manual_calc_enabled:
                    excel.Calculation = -4105  # xlCalculationAutomatic
            except Exception:
                pass
            try:
                excel.EnableEvents = True
            except Exception:
                pass
            excel.Quit()
            if page_count > 1:
                return True, f"{page_count} Excel-Kopien exportiert ({REPORT_ROWS_PER_PAGE} Positionen pro Formularseite): {exported_files[0].name} ... {exported_files[-1].name}"
            return True, f"Excel-Kopie exportiert: {exported_files[0]}"
        if action == "print":
            for page_df in pages:
                if report_df is not None:
                    _write_original_report_page(ws, page_df, int(year), int(month), str(project))
                else:
                    ws.Range("K2").Value = int(year)
                    ws.Range("K3").Value = int(month)
                    ws.Range("C12").Value = str(project)
                    _prepare_report_formula_in_excel_sheet(ws)
                _recalc()
                ws.PrintOut()
            wb.Close(SaveChanges=False)
            try:
                if manual_calc_enabled:
                    excel.Calculation = -4105  # xlCalculationAutomatic
            except Exception:
                pass
            try:
                excel.EnableEvents = True
            except Exception:
                pass
            excel.Quit()
            if page_count > 1:
                return True, f"{page_count} Druckaufträge gesendet ({REPORT_ROWS_PER_PAGE} Positionen pro Formularseite)."
            return True, "Druckauftrag an Standarddrucker gesendet."

        if action == "open":
            # Prepare first page for manual inspection in Excel. (PDF/Druck paginates automatically.)
            if report_df is not None:
                _write_original_report_page(ws, pages[0], int(year), int(month), str(project))
            else:
                ws.Range("K2").Value = int(year)
                ws.Range("K3").Value = int(month)
                ws.Range("C12").Value = str(project)
                _prepare_report_formula_in_excel_sheet(ws)
            _recalc()
            excel.Visible = True
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            if page_count > 1:
                return True, f"Originaldatei in Excel geöffnet (Seite 1/{page_count} vorbereitet). PDF/Druck erstellt weitere Seiten automatisch."
            return True, "Originaldatei in Excel geöffnet (Bericht vorbereitet)."

        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            try:
                if manual_calc_enabled:
                    excel.Calculation = -4105  # xlCalculationAutomatic
            except Exception:
                pass
            try:
                excel.EnableEvents = True
            except Exception:
                pass
            excel.Quit()
        return False, f"Unbekannte Aktion: {action} (erwartet: pdf | xlsx | print | open)"

    except Exception as e:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                try:
                    if manual_calc_enabled:
                        excel.Calculation = -4105  # xlCalculationAutomatic
                except Exception:
                    pass
                try:
                    excel.EnableEvents = True
                except Exception:
                    pass
                excel.Quit()
        except Exception:
            pass
        return False, f"Fehler bei Excel-Druck/PDF: {e}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def load_workbook_data(path_str: str) -> WorkbookData:
    path = _resolve_excel_path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {path}")
    wb = openpyxl.load_workbook(path)
    lookups = _load_lookups(wb)
    taetigkeiten_df = _read_taetigkeiten_df(wb)
    return WorkbookData(path=path, taetigkeiten_df=taetigkeiten_df, lookups=lookups)



# ------------------------- workbook writing -------------------------

def _find_next_write_row(ws) -> int:
    for r in range(2, ws.max_row + 2):
        if all(_is_blank(ws.cell(r, c).value) for c in KEY_COLS_FOR_EMPTY_CHECK):
            return r
    return ws.max_row + 1


def _write_taetigkeit_row(ws, row_idx: int, record: Dict[str, Any]) -> None:
    datum = _to_date(record.get("Datum"))
    projekt = _safe_str(record.get("Projekt")).strip()
    zeit_von = _to_time(record.get("Zeit von"))
    zeit_bis = _to_time(record.get("Zeit bis"))
    pause_min = int(record.get("Pause_Min") or 0)
    pause_time = _minutes_to_time(pause_min)
    km = record.get("km")
    try:
        km_val = None if _is_blank(km) else int(float(km))
    except Exception:
        km_val = None
    taet_typ = _safe_str(record.get("Tätigkeit")).strip()
    kodierung = _safe_str(record.get("Kodierung")).strip() or None
    interne = _safe_str(record.get("Interne Projekte")).strip() or None
    info = _safe_str(record.get("Info"))
    abgerechnet = _safe_str(record.get("Abgerechnet")).strip() or None
    eingetragen = _safe_str(record.get("eingetragen")).strip() or None

    ws.cell(row_idx, 1).value = datum
    ws.cell(row_idx, 2).value = projekt or None
    ws.cell(row_idx, 3).value = zeit_von
    ws.cell(row_idx, 4).value = zeit_bis
    ws.cell(row_idx, 5).value = pause_time
    ws.cell(row_idx, 6).value = f'=IF(A{row_idx}="","",D{row_idx}-C{row_idx}-E{row_idx})'
    ws.cell(row_idx, 7).value = f'=IF(A{row_idx}="","",F{row_idx}*24)'
    ws.cell(row_idx, 8).value = km_val
    ws.cell(row_idx, 9).value = taet_typ or None
    ws.cell(row_idx, 10).value = kodierung
    ws.cell(row_idx, 11).value = interne
    ws.cell(row_idx, 12).value = info or None
    ws.cell(row_idx, 13).value = abgerechnet
    ws.cell(row_idx, 14).value = eingetragen

    # number formats for safety
    ws.cell(row_idx, 1).number_format = "DD.MM.YYYY"
    for c in (3, 4, 5, 6):
        ws.cell(row_idx, c).number_format = "hh:mm"
    ws.cell(row_idx, 7).number_format = "0.00"


def _clear_taetigkeit_row(ws, row_idx: int) -> None:
    for c in range(1, 15):
        ws.cell(row_idx, c).value = None


def _find_hilfstabelle_project_row(ws, project_name: str) -> Optional[int]:
    target = _safe_str(project_name).strip()
    if not target:
        return None
    for r in range(2, ws.max_row + 1):
        p = _safe_str(ws.cell(r, 7).value).strip()
        if p and p == target:
            return r
    return None


def _find_next_hilfstabelle_project_row(ws) -> int:
    for r in range(2, ws.max_row + 2):
        if _is_blank(ws.cell(r, 7).value):
            return r
    return ws.max_row + 1


def _upsert_project_stammdaten(
    wb: openpyxl.Workbook,
    project_data: Dict[str, Any],
    original_project: Optional[str] = None,
    rename_taetigkeiten: bool = False,
) -> Tuple[bool, str]:
    if HILFS_SHEET not in wb.sheetnames:
        return False, f"Arbeitsblatt '{HILFS_SHEET}' nicht gefunden."
    ws = wb[HILFS_SHEET]

    new_project = _safe_str(project_data.get("Projekt")).strip()
    if not new_project:
        return False, "Projektname/-kürzel darf nicht leer sein."

    orig = _safe_str(original_project).strip() if original_project else ""
    row_idx = _find_hilfstabelle_project_row(ws, orig or new_project)

    if row_idx is None:
        existing = _find_hilfstabelle_project_row(ws, new_project)
        if existing is not None:
            row_idx = existing
        else:
            row_idx = _find_next_hilfstabelle_project_row(ws)
    else:
        if orig and orig != new_project:
            conflict = _find_hilfstabelle_project_row(ws, new_project)
            if conflict is not None and conflict != row_idx:
                return False, f"Projekt '{new_project}' existiert bereits in der Hilfstabelle (Zeile {conflict})."

    ws.cell(row_idx, 7).value = new_project
    ws.cell(row_idx, 8).value = _safe_str(project_data.get("Kunde")).strip() or None
    ws.cell(row_idx, 9).value = _safe_str(project_data.get("Straße")).strip() or None
    ws.cell(row_idx, 10).value = _safe_str(project_data.get("Ort")).strip() or None
    ws.cell(row_idx, 11).value = _safe_str(project_data.get("Ansprechpartner")).strip() or None
    ws.cell(row_idx, 12).value = _safe_str(project_data.get("Projektadresse Standard")).strip() or None
    ws.cell(row_idx, 13).value = _safe_str(project_data.get("Projektadresse Alternativ")).strip() or None
    ws.cell(row_idx, 14).value = f'=H{row_idx}&","&" "&I{row_idx}&","&" "&J{row_idx}'

    renamed_count = 0
    if rename_taetigkeiten and orig and orig != new_project and TAETIGKEITEN_SHEET in wb.sheetnames:
        tws = wb[TAETIGKEITEN_SHEET]
        for r in range(2, tws.max_row + 1):
            if _safe_str(tws.cell(r, 2).value).strip() == orig:
                tws.cell(r, 2).value = new_project
                renamed_count += 1

    msg = f"Projektstammdaten gespeichert (Hilfstabelle Zeile {row_idx})."
    if renamed_count:
        msg += f" Tätigkeiten umbenannt: {renamed_count}."
    return True, msg


def _set_relevante_kodierungen(wb: openpyxl.Workbook, selected_aufgaben: List[str]) -> Tuple[bool, str]:
    if KODIERUNG_SHEET not in wb.sheetnames:
        return False, f"Arbeitsblatt '{KODIERUNG_SHEET}' nicht gefunden."
    ws = wb[KODIERUNG_SHEET]
    selected = {_safe_str(x).strip() for x in (selected_aufgaben or []) if _safe_str(x).strip()}
    changed = 0
    total = 0
    for r in range(2, ws.max_row + 1):
        aufgabe = _safe_str(ws.cell(r, 2).value).strip()
        if not aufgabe:
            continue
        total += 1
        new_marker = "x" if aufgabe in selected else None
        old_marker = _safe_str(ws.cell(r, 1).value).strip().lower()
        normalized_old = "x" if old_marker == "x" else ""
        normalized_new = "x" if new_marker == "x" else ""
        if normalized_old != normalized_new:
            ws.cell(r, 1).value = new_marker
            changed += 1
    return True, f"Relevante Kodierungen gespeichert: {len(selected)} ausgewählt ({changed} Änderungen in {total} Kodierungen)."


def _save_workbook(path: Path, mutator) -> Tuple[bool, str]:
    """Load workbook, apply mutator(wb), save with backup. Returns (ok, message)."""
    try:
        backup_path = path.with_suffix(path.suffix + f".bak-{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}")
        shutil.copy2(path, backup_path)

        wb = openpyxl.load_workbook(path)
        try:
            if hasattr(wb, "calculation") and wb.calculation is not None:
                wb.calculation.fullCalcOnLoad = True
        except Exception:
            pass
        mutator(wb)
        wb.save(path)
        return True, f"Gespeichert. Backup erstellt: {backup_path.name}"
    except Exception as e:
        return False, f"Fehler beim Speichern: {e}"


# ------------------------- reporting logic -------------------------

def _build_report(df: pd.DataFrame, lookups: Dict[str, Any], year: int, month: int, project: str, include_abgerechnet: bool) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    x = df.copy()
    x["Datum_dt"] = pd.to_datetime(x["Datum"], errors="coerce")
    x = x[x["Datum_dt"].notna()]
    x = x[x["Datum_dt"].dt.year == int(year)]
    x = x[x["Datum_dt"].dt.month == int(month)]
    x = x[x["Projekt"].astype(str) == str(project)]

    # Excel logic: interne Tätigkeiten raus
    x = x[x["Tätigkeit"].astype(str).str.upper() != "I"]

    if not include_abgerechnet:
        x = x[x["Abgerechnet"].fillna("").astype(str).str.strip().str.lower() != "ja"]

    x = x.copy()
    x["Kodierung EB"] = x["Kodierung"].map(lookups.get("kodierung_map_eb", {}))
    x["Datum"] = x["Datum"].apply(_format_date)
    x["Beginn"] = x["Zeit von"].apply(_format_time)
    x["Ende"] = x["Zeit bis"].apply(_format_time)
    x["Pause"] = x["Pause"].apply(_format_time)
    x["Zeit (h)"] = x["Zahl"].apply(lambda v: round(float(v), 2) if v is not None and not pd.isna(v) else None)
    x["Art"] = x["Tätigkeit"]
    x["Leistungsbeschreibung"] = x["Info"].fillna("")

    # Sort ascending like report
    x["_sort_date"] = pd.to_datetime(x["Datum"], format="%d.%m.%Y", errors="coerce")
    x["_sort_start"] = x["Zeit von"].apply(lambda t: _time_to_minutes(t) if t else -1)
    x = x.sort_values(["_sort_date", "_sort_start", "_excel_row"]).drop(columns=["_sort_date", "_sort_start"])

    return x[["_excel_row", "Datum", "Beginn", "Ende", "Pause", "Zeit (h)", "Art", "Kodierung EB", "Leistungsbeschreibung", "Abgerechnet"]].reset_index(drop=True)


def _summaries_from_report(report_df: pd.DataFrame) -> Dict[str, float]:
    if report_df.empty:
        return {"F": 0.0, "R": 0.0, "K": 0.0, "gesamt": 0.0}
    sums = report_df.groupby("Art", dropna=False)["Zeit (h)"].sum(min_count=1).to_dict()
    out = {
        "F": round(float(sums.get("F", 0.0) or 0.0), 2),
        "R": round(float(sums.get("R", 0.0) or 0.0), 2),
        "K": round(float(sums.get("K", 0.0) or 0.0), 2),
    }
    out["gesamt"] = round(out["F"] + out["R"] + out["K"], 2)
    return out


# ------------------------- UI helpers -------------------------

def _project_defaults(df: pd.DataFrame) -> Tuple[int, int]:
    if df.empty or "Datum" not in df.columns:
        today = dt.date.today()
        return today.year, today.month
    dates = [d for d in df["Datum"].tolist() if isinstance(d, dt.date)]
    if not dates:
        today = dt.date.today()
        return today.year, today.month
    latest = max(dates)
    return latest.year, latest.month


def _render_taetigkeit_form(prefix: str, lookups: Dict[str, Any], defaults: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    defaults = defaults or {}
    projekte = list(dict.fromkeys([*(lookups.get("projekte") or []), _safe_str(defaults.get("Projekt"))]))
    projekte = [p for p in projekte if p]
    typen = list(dict.fromkeys([*(lookups.get("taetigkeit_typen") or []), _safe_str(defaults.get("Tätigkeit"))]))
    ja_nein = list(dict.fromkeys([*(lookups.get("ja_nein") or ["ja", "nein"]), _safe_str(defaults.get("Abgerechnet")), _safe_str(defaults.get("eingetragen"))]))
    kod_options = list(
        dict.fromkeys(
            [""]
            + (lookups.get("relevante_kodierungen") or [])
            + (lookups.get("kodierung_aufgaben") or [])
            + [_safe_str(defaults.get("Kodierung"))]
        )
    )
    interne_options = list(dict.fromkeys([""] + (lookups.get("interne_projekte") or []) + [_safe_str(defaults.get("Interne Projekte"))]))

    d_default = _to_date(defaults.get("Datum")) or dt.date.today()
    zv_default = _to_time(defaults.get("Zeit von")) or dt.time(8, 0)
    zb_default = _to_time(defaults.get("Zeit bis")) or dt.time(9, 0)
    pause_min_default = int(defaults.get("Pause_Min") or _time_to_minutes(defaults.get("Pause")) or 0)
    km_default = 0
    try:
        if not _is_blank(defaults.get("km")):
            km_default = int(float(defaults.get("km")))
    except Exception:
        km_default = 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        datum = st.date_input("Datum", value=d_default, key=f"{prefix}_datum")
        projekt = st.selectbox(
            "Projekt",
            options=projekte if projekte else [""],
            index=(projekte.index(_safe_str(defaults.get("Projekt"))) if _safe_str(defaults.get("Projekt")) in projekte else 0),
            key=f"{prefix}_projekt",
        )
    with col2:
        zeit_von = st.time_input("Zeit von", value=zv_default, key=f"{prefix}_zv")
        zeit_bis = st.time_input("Zeit bis", value=zb_default, key=f"{prefix}_zb")
    with col3:
        pause_min = st.number_input("Pause (Minuten)", min_value=0, max_value=600, value=pause_min_default, step=5, key=f"{prefix}_pause")
        km = st.number_input("km", min_value=0, value=km_default, step=1, key=f"{prefix}_km")
    with col4:
        taet_typ = st.selectbox(
            "Tätigkeit (Typ)",
            options=typen if typen else ["F", "R", "I"],
            index=(typen.index(_safe_str(defaults.get("Tätigkeit"))) if _safe_str(defaults.get("Tätigkeit")) in typen else 0),
            key=f"{prefix}_typ",
        )
        kodierung = st.selectbox(
            "Kodierung (Aufgabe)",
            options=kod_options,
            index=(kod_options.index(_safe_str(defaults.get("Kodierung"))) if _safe_str(defaults.get("Kodierung")) in kod_options else 0),
            key=f"{prefix}_kod",
        )

    col5, col6, col7 = st.columns([2, 1, 1])
    with col5:
        info = st.text_area("Info / Leistungsbeschreibung", value=_safe_str(defaults.get("Info")), height=80, key=f"{prefix}_info")
    with col6:
        interne = st.selectbox(
            "Interne Projekte",
            options=interne_options,
            index=(interne_options.index(_safe_str(defaults.get("Interne Projekte"))) if _safe_str(defaults.get("Interne Projekte")) in interne_options else 0),
            key=f"{prefix}_intern",
        )
    with col7:
        abgerechnet = st.selectbox(
            "Abgerechnet",
            options=ja_nein if ja_nein else ["ja", "nein"],
            index=((ja_nein.index(_safe_str(defaults.get("Abgerechnet"))) if _safe_str(defaults.get("Abgerechnet")) in ja_nein else (ja_nein.index("nein") if "nein" in ja_nein else 0))),
            key=f"{prefix}_abg",
        )
        eingetragen = st.selectbox(
            "eingetragen",
            options=[""] + (ja_nein if ja_nein else ["ja", "nein"]),
            index=([""] + (ja_nein if ja_nein else ["ja", "nein"])).index(_safe_str(defaults.get("eingetragen"))) if _safe_str(defaults.get("eingetragen")) in ([""] + (ja_nein if ja_nein else ["ja", "nein"])) else 0,
            key=f"{prefix}_eing",
        )

    hours = _compute_hours_decimal(zeit_von, zeit_bis, int(pause_min))
    if hours is None:
        st.warning("Zeitangaben ergeben keine gültige Dauer.")
    else:
        st.caption(f"Berechnete Dauer: **{hours:.2f} h** ({int(round(hours*60))//60:02d}:{int(round(hours*60))%60:02d})")

    return {
        "Datum": datum,
        "Projekt": projekt,
        "Zeit von": zeit_von,
        "Zeit bis": zeit_bis,
        "Pause_Min": int(pause_min),
        "km": int(km),
        "Tätigkeit": taet_typ,
        "Kodierung": kodierung,
        "Interne Projekte": interne,
        "Info": info,
        "Abgerechnet": abgerechnet,
        "eingetragen": eingetragen,
    }


def _filtered_taetigkeiten(df: pd.DataFrame, year: Optional[int], month: Optional[int], project: str, include_abgerechnet: bool) -> pd.DataFrame:
    x = df.copy()
    if x.empty:
        return x
    if year:
        x = x[x["Datum"].apply(lambda d: isinstance(d, dt.date) and d.year == year)]
    if month:
        x = x[x["Datum"].apply(lambda d: isinstance(d, dt.date) and d.month == month)]
    if project:
        x = x[x["Projekt"].astype(str) == project]
    if not include_abgerechnet:
        x = x[x["Abgerechnet"].fillna("").astype(str).str.strip().str.lower() != "ja"]
    return x.reset_index(drop=True)
def _row_key_from_series(s: pd.Series) -> tuple:
    """Key used to detect duplicates across workbooks."""
    return (
        _format_date(s.get("Datum")),
        _safe_str(s.get("Projekt")).strip(),
        _format_time(s.get("Zeit von")),
        _format_time(s.get("Zeit bis")),
        int(s.get("Pause_Min") or _time_to_minutes(s.get("Pause")) or 0),
        int(float(s.get("km") or 0) or 0),
        _safe_str(s.get("Tätigkeit")).strip(),
        _safe_str(s.get("Kodierung")).strip(),
        _safe_str(s.get("Interne Projekte")).strip(),
        _safe_str(s.get("Info")).strip(),
    )


def _existing_keys(df: pd.DataFrame) -> set:
    if df is None or df.empty:
        return set()
    keys = set()
    for _, r in df.iterrows():
        keys.add(_row_key_from_series(r))
    return keys


def _import_df_from_xlsx(path: Path) -> pd.DataFrame:
    """Load source workbook and read the 'Tätigkeiten' sheet using the same parser as the app."""
    wb = openpyxl.load_workbook(path)
    if TAETIGKEITEN_SHEET not in wb.sheetnames:
        raise ValueError(f"Quelle hat kein Sheet '{TAETIGKEITEN_SHEET}'.")
    return _read_taetigkeiten_df(wb)


def _store_uploaded_import_file(uploaded_file) -> Path:
    """Store uploaded import source file to disk (separate from 'working copy' logic)."""
    if uploaded_file is None:
        raise ValueError("Keine Datei hochgeladen.")
    name = Path(getattr(uploaded_file, "name", "import.xlsx")).name
    if Path(name).suffix.lower() != ".xlsx":
        raise ValueError("Bitte eine .xlsx-Datei hochladen.")

    script_dir = Path(__file__).resolve().parent
    imports_dir = script_dir / "imports_lines"
    imports_dir.mkdir(parents=True, exist_ok=True)

    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = imports_dir / f"{Path(name).stem}_lines_{ts}.xlsx"
    target.write_bytes(uploaded_file.getvalue())
    return target.resolve()


def _series_to_write_record(s: pd.Series, project_value: str) -> Dict[str, Any]:
    """Convert a parsed Tätigkeiten row into a write payload for _write_taetigkeit_row."""
    pause_min = int(s.get("Pause_Min") or _time_to_minutes(s.get("Pause")) or 0)
    return {
        "Datum": _to_date(s.get("Datum")),
        "Projekt": project_value,
        "Zeit von": _to_time(s.get("Zeit von")),
        "Zeit bis": _to_time(s.get("Zeit bis")),
        "Pause_Min": pause_min,
        "km": int(float(s.get("km") or 0) or 0),
        "Tätigkeit": _safe_str(s.get("Tätigkeit")).strip(),
        "Kodierung": _safe_str(s.get("Kodierung")).strip(),
        "Interne Projekte": _safe_str(s.get("Interne Projekte")).strip(),
        "Info": _safe_str(s.get("Info")),
        "Abgerechnet": _safe_str(s.get("Abgerechnet")).strip() or None,
        "eingetragen": _safe_str(s.get("eingetragen")).strip() or None,
    }
def _store_uploaded_report_file(uploaded_file) -> Path:
    if uploaded_file is None:
        raise ValueError("Keine Datei hochgeladen.")
    name = Path(getattr(uploaded_file, "name", "report.xlsx")).name
    if Path(name).suffix.lower() != ".xlsx":
        raise ValueError("Bitte eine .xlsx-Datei hochladen.")

    script_dir = Path(__file__).resolve().parent
    reports_dir = script_dir / "imports_reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = reports_dir / f"{Path(name).stem}_report_{ts}.xlsx"
    target.write_bytes(uploaded_file.getvalue())
    return target.resolve()


def _build_reverse_kod_map_eb_to_aufgabe(lookups: Dict[str, Any]) -> Dict[str, str]:
    # Aufgabe -> EB  ==>  EB -> Aufgabe (nur wenn eindeutig)
    fwd = lookups.get("kodierung_map_eb", {}) or {}
    rev: Dict[str, str] = {}
    collisions = set()
    for aufgabe, eb in fwd.items():
        eb_s = _safe_str(eb).strip()
        aufg_s = _safe_str(aufgabe).strip()
        if not eb_s or not aufg_s:
            continue
        if eb_s in rev and rev[eb_s] != aufg_s:
            collisions.add(eb_s)
        else:
            rev[eb_s] = aufg_s
    for c in collisions:
        rev.pop(c, None)
    return rev

import re

GER_MONTHS = {
    "januar": 1, "februar": 2, "märz": 3, "maerz": 3, "april": 4, "mai": 5, "juni": 6,
    "juli": 7, "august": 8, "september": 9, "oktober": 10, "november": 11, "dezember": 12,
}

def _parse_month_year_from_filename(p: Path) -> Tuple[Optional[int], Optional[int]]:
    s = p.stem.lower()
    y = None
    m = None

    my = re.search(r"\b(20\d{2})\b", s)
    if my:
        y = int(my.group(1))

    mm = re.search(r"\b(januar|februar|märz|maerz|april|mai|juni|juli|august|september|oktober|november|dezember)\b", s)
    if mm:
        m = GER_MONTHS.get(mm.group(1))

    # optional: auch "02 februar" o.ä. enthalten -> reicht meist mit Monatswort
    return y, m

def _read_project_from_sheet(ws) -> str:
    # 1) Template-style
    c12 = _safe_str(ws["C12"].value).strip()
    if c12:
        return c12

    # 2) Alt/Export-style: "Firma:" rechts daneben (z.B. A6/B6)
    for r in range(1, 50):
        for c in range(1, 12):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() == "firma:":
                right = _safe_str(ws.cell(r, c + 1).value).strip()
                if right:
                    return right
    return ""

def _find_header(ws, max_rows: int = 80, max_cols: int = 30) -> Tuple[Optional[int], Dict[str, int]]:
    required = {"datum", "beginn", "ende", "art"}
    for r in range(1, max_rows + 1):
        mapping: Dict[str, int] = {}
        for c in range(1, max_cols + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                mapping[v.strip().lower()] = c
        if required.issubset(mapping.keys()):
            return r, mapping
    return None, {}

def _read_einsatzbericht_xlsx(path: Path) -> Tuple[Dict[str, Any], pd.DataFrame]:
    wb = openpyxl.load_workbook(path, data_only=True)

    # Sheet fallback
    ws = None
    for sheet_name in ["Einsatzbericht", "Tabelle1"]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            break
    if ws is None:
        # notfalls erstes Sheet
        ws = wb[wb.sheetnames[0]]

    header_row, col = _find_header(ws)
    if header_row is None:
        meta = {"project": "", "year": None, "month": None, "source_path": str(path)}
        return meta, pd.DataFrame(columns=["Datum","Beginn","Ende","Pause_Min","Zeit_h","Art","Kodierung_EB","Leistungsbeschreibung"])

    # Meta robust
    project = _read_project_from_sheet(ws)

    year = ws["K2"].value
    month = ws["K3"].value
    if _is_blank(year) or _is_blank(month):
        y2, m2 = _parse_month_year_from_filename(path)
        year = year if not _is_blank(year) else y2
        month = month if not _is_blank(month) else m2

    def c(name: str) -> Optional[int]:
        return col.get(name.lower())

    c_datum = c("datum")
    c_beginn = c("beginn")
    c_ende = c("ende")
    c_pause = c("pause (h)") or c("pause")
    c_zeit = c("zeit  (h)") or c("zeit (h)") or c("zeit")
    c_art = c("art")
    c_text = c("leistungsbeschreibung") or c("beschreibung")
    c_kod = c("kodierung") or c("kodierung eb")  # optional

    rows = []

    empty_streak = 0
    start = header_row + 1

    for r in range(start, start + 400):
        a = ws.cell(r, c_datum).value if c_datum else None
        b = ws.cell(r, c_beginn).value if c_beginn else None
        e = ws.cell(r, c_ende).value if c_ende else None
        t = ws.cell(r, c_text).value if c_text else None

        if _is_blank(a) and _is_blank(b) and _is_blank(e) and _is_blank(t):
            empty_streak += 1
            if empty_streak >= 5:
                break
            continue
        empty_streak = 0

        datum = _to_date(a)
        beginn = _to_time(b)
        ende = _to_time(e)

        pause_min = 0
        if c_pause:
            pause = _to_time(ws.cell(r, c_pause).value)
            pause_min = _time_to_minutes(pause)

        zeit_h = None
        if c_zeit:
            v = ws.cell(r, c_zeit).value
            try:
                zeit_h = None if _is_blank(v) else float(v)
            except Exception:
                zeit_h = None

        # nur rechnen, wenn Beginn+Ende existieren
        if zeit_h is None and beginn is not None and ende is not None:
            zeit_h = _compute_hours_decimal(beginn, ende, pause_min)

        art = _safe_str(ws.cell(r, c_art).value).strip() if c_art else ""
        kod_eb = _safe_str(ws.cell(r, c_kod).value).strip() if c_kod else ""
        text = _safe_str(t).strip()

        rows.append({
            "Datum": datum,
            "Beginn": beginn,
            "Ende": ende,
            "Pause_Min": pause_min,
            "Zeit_h": zeit_h,
            "Art": art,
            "Kodierung_EB": kod_eb,
            "Leistungsbeschreibung": text,
        })

    df_lines = pd.DataFrame(rows)

    # final meta fallback: wenn year/month immer noch fehlen -> aus erstem Datum
    if (year is None or month is None) and not df_lines.empty:
        first_date = df_lines["Datum"].dropna().iloc[0] if df_lines["Datum"].notna().any() else None
        if isinstance(first_date, dt.date):
            year = year or first_date.year
            month = month or first_date.month

    meta = {
        "project": project,
        "year": year,
        "month": month,
        "source_path": str(path),
    }
    return meta, df_lines

def _key_for_import(rec: Dict[str, Any]) -> tuple:
    d = _format_date(rec.get("Datum"))
    zv = _format_time(rec.get("Zeit von"))
    zb = _format_time(rec.get("Zeit bis"))
    info = _safe_str(rec.get("Info")).strip()
    typ = _safe_str(rec.get("Tätigkeit")).strip()
    pause = int(rec.get("Pause_Min") or 0)

    # Wenn keine Zeiten vorhanden sind, nimm Info+Zeit_h in den Key,
    # sonst kollidiert alles bei ("", "", "", ...)
    if not zv and not zb:
        zeit_h = rec.get("Zeit_h") or rec.get("Zeit (h)") or ""
        return (d, "<NO_TIME>", typ, str(zeit_h), info)

    return (d, zv, zb, pause, typ, info)


def _existing_keys_for_master(df_master: pd.DataFrame) -> set:
    keys = set()
    if df_master is None or df_master.empty:
        return keys
    for _, r in df_master.iterrows():
        rec = {
            "Datum": r.get("Datum"),
            "Projekt": r.get("Projekt"),
            "Zeit von": r.get("Zeit von"),
            "Zeit bis": r.get("Zeit bis"),
            "Pause_Min": int(r.get("Pause_Min") or _time_to_minutes(r.get("Pause")) or 0),
            "km": r.get("km") or 0,
            "Tätigkeit": r.get("Tätigkeit"),
            "Kodierung": r.get("Kodierung") or "",
            "Info": r.get("Info") or "",
        }
        keys.add(_key_for_import(rec))
    return keys
# ------------------------- Streamlit UI -------------------------

def main() -> None:
    if st is None:
        raise RuntimeError("Streamlit ist nicht installiert. Bitte `pip install streamlit` ausführen.")
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Lokale Webapp für Tätigkeiten-Erfassung und Einsatzbericht-Auswertung (Excel-basiert, MVP)")

    # Robust default path handling: if a remembered path no longer exists (e.g. moved folder),
    # fall back to auto-discovery instead of showing a broken default forever.
    remembered_path = _safe_str(st.session_state.get("excel_path", "")).strip()
    default_path = ""
    if remembered_path:
        try:
            remembered_resolved = _resolve_excel_path(remembered_path)
            if remembered_resolved.exists():
                default_path = remembered_path
        except Exception:
            default_path = ""
    if not default_path:
        for cand in _default_excel_candidates():
            if cand.exists():
                try:
                    script_dir = Path(__file__).resolve().parent
                    default_path = str(cand.resolve().relative_to(script_dir))
                except Exception:
                    default_path = str(cand.resolve())
                break
    if not default_path:
        default_path = "data/Tätigkeiten_Überblick.xlsx"
    pending_import_path = _safe_str(st.session_state.get("_pending_import_excel_path", "")).strip()
    if pending_import_path:
        st.session_state["excel_path_input"] = pending_import_path
        st.session_state["excel_path"] = pending_import_path
        st.session_state["_pending_import_excel_path"] = ""
    # Keep the visible sidebar widget in sync with the remembered path.
    if _safe_str(st.session_state.get("excel_path_input", "")).strip():
        # If the visible widget points to a non-existent path, reset it too.
        try:
            _tmp_vis = _resolve_excel_path(_safe_str(st.session_state.get("excel_path_input", "")).strip())
            if not _tmp_vis.exists():
                st.session_state["excel_path_input"] = default_path
        except Exception:
            st.session_state["excel_path_input"] = default_path
    else:
        st.session_state["excel_path_input"] = default_path

    with st.sidebar:
        st.header("Datei")
        excel_path = st.text_input("Pfad zur Excel-Datei", key="excel_path_input")
        uploaded_excel = st.file_uploader(
            "Excel-Datei importieren (.xlsx)",
            type=["xlsx"],
            key="excel_upload_file",
            help="Lädt eine Excel-Datei hoch und speichert sie als lokale Arbeitskopie (für Bearbeitung + Excel-Druck/PDF).",
        )

        if uploaded_excel is not None:
            if st.button("Upload als Arbeitskopie übernehmen", key="import_uploaded_excel_btn"):
                try:
                    imported_path = _store_uploaded_excel(uploaded_excel)
                    st.session_state["_pending_import_excel_path"] = str(imported_path)
                    st.session_state["excel_path"] = str(imported_path)  # optional
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Import fehlgeschlagen: {e}")
        st.session_state["excel_path"] = excel_path
        reload_clicked = st.button("Neu laden")
        st.info(
            "Hinweis: Relative Pfade sind erlaubt (z. B. `Tätigkeiten_Überblick.xlsx` oder `data/Tätigkeiten_Überblick.xlsx`).\n\n"
            "Beim Speichern wird eine Backup-Datei erstellt. openpyxl kann eingebettete WMF-Bilder nicht vollständig erhalten; arbeite deshalb am besten auf einer Kopie der Datei."
        )

    if reload_clicked:
        st.cache_data.clear()

    try:
        data = load_workbook_data(excel_path)
    except Exception as e:
        st.error(f"Datei konnte nicht geladen werden: {e}")
        st.stop()

    df = data.taetigkeiten_df.copy()
    lookups = data.lookups

    tab1, tab2, tab3 = st.tabs(["Tätigkeiten", "Einsatzbericht", "Stammdaten / Debug"])

    with tab1:
        st.subheader("Tätigkeiten erfassen und bearbeiten")

        y_default, m_default = _project_defaults(df)
        projekte_available = sorted(list(dict.fromkeys([p for p in (lookups.get("projekte") or []) if p] + [p for p in df.get("Projekt", pd.Series(dtype=str)).dropna().astype(str).tolist() if p])))
        filt_cols = st.columns([1, 1, 2, 1])
        with filt_cols[0]:
            f_year = st.number_input("Jahr-Filter", min_value=2000, max_value=2100, value=int(y_default), step=1)
        with filt_cols[1]:
            f_month = st.selectbox("Monat-Filter", options=list(range(1, 13)), index=max(0, min(11, int(m_default) - 1)))
        with filt_cols[2]:
            f_project = st.selectbox("Projekt-Filter", options=[""] + projekte_available, index=0)
        with filt_cols[3]:
            f_include_abg = st.checkbox("abgerechnete zeigen", value=True)

        # Optionen für Editor (immer definiert -> IDE meckert nicht)
        projekte_opts = projekte_available if projekte_available else [""]

        typen_opts = list(dict.fromkeys(lookups.get("taetigkeit_typen") or ["F", "R", "I", "K"]))
        ja_nein_opts = list(dict.fromkeys(lookups.get("ja_nein") or ["ja", "nein"]))

        kod_opts = list(dict.fromkeys(
            [""] + (lookups.get("relevante_kodierungen") or []) + (lookups.get("kodierung_aufgaben") or [])
        ))

        interne_opts = list(dict.fromkeys([""] + (lookups.get("interne_projekte") or [])))

        filtered = _filtered_taetigkeiten(df, int(f_year), int(f_month), f_project, include_abgerechnet=f_include_abg)
        # --- Inline Edit Table (statt nur Anzeige) ---
        st.markdown("### Tätigkeiten (Inline bearbeiten)")

        if filtered.empty:
            st.info("Keine Tätigkeiten für den aktuellen Filter gefunden.")
        else:
            # Spalten, die im Grid editierbar sein sollen
            editor_cols = [
                "_excel_row",
                "Datum",
                "Projekt",
                "Zeit von",
                "Zeit bis",
                "Pause_Min",
                "Dauer",
                "km",
                "Tätigkeit",
                "Kodierung",
                "Interne Projekte",
                "Info",
                "Abgerechnet",
                "eingetragen",
            ]

            editor_df = filtered.copy()

            # Ensure editor columns exist
            for c in editor_cols:
                if c not in editor_df.columns:
                    editor_df[c] = None

            # Optional: Delete marker column
            editor_df["Löschen"] = False


            # Streamlit compatibility (falls du eine ältere Version hast)
            data_editor_fn = getattr(st, "data_editor", None) or getattr(st, "experimental_data_editor")

            def _calc_dauer_str(zv, zb, pause_min) -> str:
                h = _compute_hours_decimal(zv, zb, int(pause_min or 0))
                if h is None:
                    return ""
                mins = int(round(h * 60))
                return f"{mins // 60:02d}:{mins % 60:02d}"

            editor_df["Dauer"] = editor_df.apply(
                lambda r: _calc_dauer_str(r.get("Zeit von"), r.get("Zeit bis"), r.get("Pause_Min")),
                axis=1
            )


            edited_df = data_editor_fn(
                editor_df[editor_cols + ["Löschen"]],
                key="taetigkeiten_inline_editor",
                use_container_width=True,
                height=420,
                num_rows="dynamic",  # <-- erlaubt neue Zeilen unten im Grid
                hide_index=True,
                disabled=["_excel_row", "Dauer"],  # Excel-Row ist die technische ID
                column_config={
                    "_excel_row": st.column_config.NumberColumn("Excel-Zeile",
                                                                help="Technische Zeilen-ID (nicht ändern)"),
                    "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY"),
                    "Projekt": st.column_config.SelectboxColumn("Projekt", options=projekte_opts),
                    "Zeit von": st.column_config.TimeColumn("Zeit von", format="HH:mm"),
                    "Zeit bis": st.column_config.TimeColumn("Zeit bis", format="HH:mm"),
                    "Pause_Min": st.column_config.NumberColumn("Pause (Min)", min_value=0, max_value=600, step=5),
                    "Dauer": st.column_config.TextColumn("Dauer", help="Berechnet aus Zeit von/bis und Pause", width="small"),
                    "km": st.column_config.NumberColumn("km", min_value=0, step=1),
                    "Tätigkeit": st.column_config.SelectboxColumn("Tätigkeit", options=typen_opts),
                    "Kodierung": st.column_config.SelectboxColumn("Kodierung (Aufgabe)", options=kod_opts),
                    "Interne Projekte": st.column_config.SelectboxColumn("Interne Projekte", options=interne_opts),
                    "Info": st.column_config.TextColumn("Leistungsbeschreibung", width="large"),
                    "Abgerechnet": st.column_config.SelectboxColumn("Abgerechnet", options=[""] + ja_nein_opts),
                    "eingetragen": st.column_config.SelectboxColumn("eingetragen", options=[""] + ja_nein_opts),
                    "Löschen": st.column_config.CheckboxColumn("Löschen"),
                },
            )

            st.caption(
                "✔ Du kannst direkt in der Tabelle editieren. "
                "➕ Neue Zeilen unten hinzufügen. "
                "🗑️ Zum Löschen Checkbox 'Löschen' setzen und speichern."
            )

            def _editor_row_to_record(r: pd.Series) -> Dict[str, Any]:
                pause_min = 0
                try:
                    pause_min = int(r.get("Pause_Min") or 0)
                except Exception:
                    pause_min = 0

                km_val = 0
                try:
                    km_val = int(float(r.get("km") or 0) or 0)
                except Exception:
                    km_val = 0

                return {
                    "Datum": _to_date(r.get("Datum")),
                    "Projekt": _safe_str(r.get("Projekt")).strip(),
                    "Zeit von": _to_time(r.get("Zeit von")),
                    "Zeit bis": _to_time(r.get("Zeit bis")),
                    "Pause_Min": pause_min,
                    "km": km_val,
                    "Tätigkeit": _safe_str(r.get("Tätigkeit")).strip(),
                    "Kodierung": _safe_str(r.get("Kodierung")).strip(),
                    "Interne Projekte": _safe_str(r.get("Interne Projekte")).strip(),
                    "Info": _safe_str(r.get("Info")),
                    "Abgerechnet": _normalize_yes_no(r.get("Abgerechnet")) or _safe_str(r.get("Abgerechnet")).strip(),
                    "eingetragen": _normalize_yes_no(r.get("eingetragen")) or _safe_str(r.get("eingetragen")).strip(),
                }

            def _is_editor_row_blank(r: pd.Series) -> bool:
                # Nutze deine KEY_COLS_FOR_EMPTY_CHECK Logik sinngemäß (Datum/Projekt/Zeit/Typ/Info/Flags)
                return (
                        _is_blank(r.get("Datum")) and
                        _is_blank(r.get("Projekt")) and
                        _is_blank(r.get("Zeit von")) and
                        _is_blank(r.get("Zeit bis")) and
                        _is_blank(r.get("Tätigkeit")) and
                        _is_blank(r.get("Info")) and
                        _is_blank(r.get("Abgerechnet")) and
                        _is_blank(r.get("eingetragen"))
                )

            if st.button("Tabellenänderungen speichern", key="save_inline_table"):
                # Original-Signaturen nach Excel-Zeile, um Änderungen zu erkennen
                orig_by_row = {}
                for _, r in editor_df.iterrows():
                    exr = r.get("_excel_row")
                    if exr is None or (isinstance(exr, float) and pd.isna(exr)):
                        continue
                    try:
                        orig_by_row[int(exr)] = _editor_row_to_record(r)
                    except Exception:
                        continue

                updates: List[Tuple[int, Dict[str, Any]]] = []
                inserts: List[Dict[str, Any]] = []
                deletes: List[int] = []

                for _, r in edited_df.iterrows():
                    exr = r.get("_excel_row")
                    mark_delete = bool(r.get("Löschen", False))

                    # neue Zeile?
                    is_new = (exr is None) or (isinstance(exr, float) and pd.isna(exr))

                    if is_new:
                        if _is_editor_row_blank(r):
                            continue
                        rec = _editor_row_to_record(r)
                        # minimale Validierung: Projekt + Datum sollte da sein
                        if rec.get("Projekt") and rec.get("Datum"):
                            inserts.append(rec)
                        continue

                    row_excel = int(exr)

                    if mark_delete:
                        deletes.append(row_excel)
                        continue

                    new_rec = _editor_row_to_record(r)
                    old_rec = orig_by_row.get(row_excel)

                    # Wenn wir keinen "alt"-Datensatz haben, behandeln wir es als Update
                    if old_rec is None or _key_for_import(new_rec) != _key_for_import(old_rec):
                        updates.append((row_excel, new_rec))

                def _mutator_inline(wb):
                    ws = wb[TAETIGKEITEN_SHEET]

                    # Deletes
                    for r in deletes:
                        _clear_taetigkeit_row(ws, r)

                    # Updates
                    for r, rec in updates:
                        _write_taetigkeit_row(ws, r, rec)

                    # Inserts
                    row_idx = _find_next_write_row(ws)
                    for rec in inserts:
                        _write_taetigkeit_row(ws, row_idx, rec)
                        row_idx += 1

                ok, msg = _save_workbook(data.path, _mutator_inline)
                if ok:
                    st.success(
                        f"Gespeichert. Updates: {len(updates)}, Neu: {len(inserts)}, Gelöscht: {len(deletes)}. {msg}"
                    )
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(msg)

        st.markdown("---")
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("### Neue Tätigkeit anlegen")
            with st.form("add_form", clear_on_submit=False):
                defaults_add = {
                    "Datum": dt.date.today(),
                    "Projekt": f_project or (projekte_available[0] if projekte_available else ""),
                    "Tätigkeit": "F" if "F" in (lookups.get("taetigkeit_typen") or []) else (lookups.get("taetigkeit_typen") or [""])[0],
                    "Abgerechnet": "nein" if "nein" in (lookups.get("ja_nein") or []) else "",
                }
                rec_add = _render_taetigkeit_form("add", lookups, defaults=defaults_add)
                submit_add = st.form_submit_button("Eintrag speichern")
            if submit_add:
                def _mutator_add(wb):
                    ws = wb[TAETIGKEITEN_SHEET]
                    row_idx = _find_next_write_row(ws)
                    _write_taetigkeit_row(ws, row_idx, rec_add)
                ok, msg = _save_workbook(data.path, _mutator_add)
                if ok:
                    st.success(msg)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(msg)

        with c2:
            st.markdown("### Bestehenden Eintrag bearbeiten")
            if filtered.empty:
                st.info("Zum Bearbeiten zuerst oben einen Filter mit Treffern wählen.")
            else:
                options = filtered.to_dict(orient="records")
                labels = [_display_row_label(pd.Series(o)) for o in options]
                idx = st.selectbox("Eintrag auswählen", options=list(range(len(options))), format_func=lambda i: labels[i], key="edit_row_selector")
                selected = options[int(idx)]

                with st.form("edit_form", clear_on_submit=False):
                    rec_edit = _render_taetigkeit_form("edit", lookups, defaults=selected)
                    save_edit = st.form_submit_button("Änderungen speichern")
                    delete_row = st.form_submit_button("Zeile leeren (vorsichtig)")

                row_excel = int(selected["_excel_row"])
                if save_edit:
                    def _mutator_edit(wb):
                        ws = wb[TAETIGKEITEN_SHEET]
                        _write_taetigkeit_row(ws, row_excel, rec_edit)
                    ok, msg = _save_workbook(data.path, _mutator_edit)
                    if ok:
                        st.success(f"Zeile {row_excel} aktualisiert. {msg}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(msg)
                if delete_row:
                    def _mutator_delete(wb):
                        ws = wb[TAETIGKEITEN_SHEET]
                        _clear_taetigkeit_row(ws, row_excel)
                    ok, msg = _save_workbook(data.path, _mutator_delete)
                    if ok:
                        st.success(f"Zeile {row_excel} geleert. {msg}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(msg)

        st.markdown("---")
        st.markdown("### Schnellaktion")
        if not filtered.empty:
            if st.button("Gefilterte Treffer als 'abgerechnet = ja' markieren"):
                rows_to_mark = [int(r) for r in filtered["_excel_row"].tolist()]
                def _mutator_mark(wb):
                    ws = wb[TAETIGKEITEN_SHEET]
                    for r in rows_to_mark:
                        ws.cell(r, 13).value = "ja"
                ok, msg = _save_workbook(data.path, _mutator_mark)
                if ok:
                    st.success(f"{len(rows_to_mark)} Einträge markiert. {msg}")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(msg)
        else:
            st.caption("Keine Treffer für Schnellaktion.")
        st.markdown("---")
        st.markdown("### Migration: Einsatzbericht importieren")

        with st.expander("Einsatzbericht(e) (Excel-Layout) hochladen und in Tätigkeiten übernehmen"):
            report_files = st.file_uploader(
                "Einsatzbericht-Excel auswählen (.xlsx) – gerne mehrere Dateien (z. B. _01, _02 ...)",
                type=["xlsx"],
                accept_multiple_files=True,
                key="upload_reports",
            )

            default_km = st.number_input("km (Default für importierte Zeilen)", min_value=0, value=0, step=1,
                                         key="import_km_default")
            set_eingetragen = st.checkbox("eingetragen automatisch = ja", value=True, key="import_set_eingetragen")
            set_abgerechnet = st.checkbox("abgerechnet automatisch = nein", value=True, key="import_set_abgerechnet")
            tag_info = st.checkbox("Info markieren mit ''", value=True, key="import_tag_info")

            if report_files:
                rev_kod_map = _build_reverse_kod_map_eb_to_aufgabe(lookups)
                existing_keys = _existing_keys_for_master(df)

                all_records: List[Dict[str, Any]] = []
                meta_list = []

                # Projekte im Ziel
                target_projects = sorted(list(dict.fromkeys(
                    [p for p in (lookups.get("projekte") or []) if p] +
                    [p for p in df.get("Projekt", pd.Series(dtype=str)).dropna().astype(str).tolist() if p]
                )))

                for i, uf in enumerate(report_files, start=1):
                    try:
                        p = _store_uploaded_report_file(uf)
                        meta, lines = _read_einsatzbericht_xlsx(p)
                        st.write("Import-Zeilen erkannt:", len(lines))
                        st.dataframe(lines.head())
                        meta_list.append(meta)

                        detected_project = meta.get("project") or ""
                        st.write(
                            f"""**Datei {i}:** `{Path(meta['source_path']).name}`
                        Projekt erkannt: `{detected_project or '-'}`
                        Monat/Jahr: {meta.get('month')}/{meta.get('year')}"""
                        )

                        # pro Datei: Mapping detected -> target
                        default_target = detected_project if detected_project in target_projects else (
                            target_projects[0] if target_projects else "")
                        target_project = st.selectbox(
                            f"Ziel-Projekt für Datei {i}",
                            options=target_projects if target_projects else [""],
                            index=(target_projects.index(default_target) if default_target in target_projects else 0),
                            key=f"import_report_target_{i}",
                        )

                        if lines.empty:
                            st.info("Keine Zeilen im Einsatzbericht gefunden.")
                            continue

                        for _, row in lines.iterrows():
                            datum = row.get("Datum")
                            beginn = row.get("Beginn")
                            ende = row.get("Ende")
                            pause_min = int(row.get("Pause_Min") or 0)
                            art = _safe_str(row.get("Art")).strip()
                            kod_eb = _safe_str(row.get("Kodierung_EB")).strip()
                            text = _safe_str(row.get("Leistungsbeschreibung")).strip()

                            # Reverse mapping EB -> Aufgabe (falls möglich)
                            aufgabe = rev_kod_map.get(kod_eb, "")
                            info = text
                            if kod_eb and not aufgabe:
                                # EB-Code mitführen, falls kein Reverse-Mapping möglich
                                info = (info + f"  [EB:{kod_eb}]").strip()

                            rec = {
                                "Datum": datum,
                                "Projekt": target_project,
                                "Zeit von": beginn,
                                "Zeit bis": ende,
                                "Pause_Min": pause_min,
                                "km": int(default_km),
                                "Tätigkeit": art,
                                "Kodierung": aufgabe or "",  # Aufgabe (wenn gefunden)
                                "Interne Projekte": "",
                                "Info": info,
                                "Abgerechnet": "nein" if set_abgerechnet else "",
                                "eingetragen": "ja" if set_eingetragen else "",
                                "Zeit_h": row.get("Zeit_h"),
                            }


                            # If Info contains "organi" then

                            k = _key_for_import(rec)
                            if k in existing_keys:
                                continue
                            existing_keys.add(k)
                            all_records.append(rec)

                    except Exception as e:
                        st.error(f"Fehler beim Lesen der Datei {i}: {e}")

                st.metric("Neu zu importierende Zeilen", len(all_records))

                if all_records:
                    preview = pd.DataFrame([{
                        "Datum": _format_date(r["Datum"]),
                        "Projekt": r["Projekt"],
                        "Zeit von": _format_time(r["Zeit von"]),
                        "Zeit bis": _format_time(r["Zeit bis"]),
                        "Pause (Min)": r["Pause_Min"],
                        "Typ": r["Tätigkeit"],
                        "Kodierung (Aufgabe)": r["Kodierung"],
                        "Info": _safe_str(r["Info"])[:80],
                    } for r in all_records[:50]])
                    st.dataframe(preview, use_container_width=True, height=260)

                    if st.button("Import durchführen (in Tätigkeiten schreiben)", key="commit_report_import"):
                        def _mutator_import_reports(wb):
                            ws = wb[TAETIGKEITEN_SHEET]
                            row_idx = _find_next_write_row(ws)
                            for rec in all_records:
                                _write_taetigkeit_row(ws, row_idx, rec)
                                row_idx += 1

                        ok, msg = _save_workbook(data.path, _mutator_import_reports)
                        if ok:
                            st.success(f"Import OK: {len(all_records)} Zeilen übernommen. {msg}")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error(msg)
    with tab2:
        st.subheader("Einsatzbericht (Web-Ansicht)")

        rep_col1, rep_col2, rep_col3, rep_col4 = st.columns([1, 1, 2, 1])
        rep_projects = sorted(list(dict.fromkeys([p for p in (lookups.get("projekte") or []) if p] + [p for p in df.get("Projekt", pd.Series(dtype=str)).dropna().astype(str).tolist() if p])))
        y_default, m_default = _project_defaults(df)
        with rep_col1:
            r_year = st.number_input("Jahr", min_value=2000, max_value=2100, value=int(y_default), step=1, key="rep_year")
        with rep_col2:
            r_month = st.selectbox("Monat", options=list(range(1, 13)), index=max(0, min(11, int(m_default)-1)), key="rep_month")
        with rep_col3:
            # Prefer ABS if available, otherwise first project
            default_project = "ABS" if "ABS" in rep_projects else (rep_projects[0] if rep_projects else "")
            r_project = st.selectbox("Projekt", options=rep_projects if rep_projects else [""], index=(rep_projects.index(default_project) if default_project in rep_projects else 0), key="rep_project")
        with rep_col4:
            include_abgerechnet = st.checkbox("abgerechnete einschließen", value=False)

        report_df = _build_report(df, lookups, int(r_year), int(r_month), r_project, include_abgerechnet)
        sums = _summaries_from_report(report_df)

        info = lookups.get("projekt_infos", {}).get(r_project, {})
        p1, p2 = st.columns(2)
        with p1:
            st.markdown("#### Leistungsnehmer")
            st.write(f"**Firma:** {_safe_str(info.get('Kunde')) or '-'}")
            st.write(f"**Straße:** {_safe_str(info.get('Straße')) or '-'}")
            st.write(f"**Ort:** {_safe_str(info.get('Ort')) or '-'}")
            st.write(f"**Kontakt:** {_safe_str(info.get('Ansprechpartner')) or '-'}")
        with p2:
            st.markdown("#### Report-Kontext")
            st.write(f"**Projekt:** {r_project}")
            st.write(f"**Monat/Jahr:** {int(r_month):02d}/{int(r_year)}")
            st.write(f"**Projektadresse (Standard):** {_safe_str(info.get('Projektadresse Standard')) or '-'}")
            alt = _safe_str(info.get('Projektadresse Alternativ'))
            if alt:
                st.write(f"**Projektadresse (Alternativ):** {alt}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Beratung (F)", f"{sums['F']:.2f} h")
        m2.metric("Reisezeit (R)", f"{sums['R']:.2f} h")
        m3.metric("Kulanz (K)", f"{sums['K']:.2f} h")
        m4.metric("Gesamt", f"{sums['gesamt']:.2f} h")

        if len(report_df) > REPORT_ROWS_PER_PAGE:
            st.info(f"Es werden automatisch mehrere Formularseiten erzeugt: {len(report_df)} Positionen = {math.ceil(len(report_df)/REPORT_ROWS_PER_PAGE)} Seiten à {REPORT_ROWS_PER_PAGE} Positionen.")

        st.markdown("### Original-Einsatzbericht (Excel-Layout)")
        st.caption("Verwendet Microsoft Excel (pywin32) und befüllt das Original-Formular direkt aus der App-Logik (robust, ohne fragile Excel-FILTER-Abhängigkeit). Bei mehr als 19 Positionen werden automatisch mehrere Seiten/PDFs/Excel-Kopien erzeugt.")
        export_dir = data.path.parent / "exports"
        base_export_name = f"Einsatzbericht_{r_project}_{int(r_year)}-{int(r_month):02d}"
        default_pdf_name = f"{base_export_name}.pdf"
        default_xlsx_name = f"{base_export_name}.xlsx"

        pdf_rel_or_abs = st.text_input(
            "PDF-Ausgabepfad",
            value=str((export_dir / default_pdf_name)),
            help="Relativer oder absoluter Pfad. Bei relativem Pfad wird vom aktuellen Startordner ausgegangen.",
            key="orig_pdf_path",
        )

        xlsx_rel_or_abs = st.text_input(
            "Excel-Kopie-Ausgabepfad (Original-Layout)",
            value=str((export_dir / default_xlsx_name)),
            help="Speichert eine vorbereitete Kopie der Excel mit befülltem Original-Einsatzbericht. Bei mehreren Seiten werden _01, _02, ... angehängt.",
            key="orig_xlsx_copy_path",
        )

        b1, b2, b3, b4 = st.columns(4)

        with b1:
            if st.button("Original in Excel öffnen", key="open_original_excel"):
                ok, msg = _excel_original_report_action(
                    data.path,
                    int(r_year),
                    int(r_month),
                    r_project,
                    action="open",
                    report_df=report_df,
                )
                (st.success if ok else st.error)(msg)

        with b2:
            if st.button("Original als PDF (Excel)", key="export_original_pdf"):
                pdf_target = Path(pdf_rel_or_abs).expanduser()
                if not pdf_target.is_absolute():
                    pdf_target = (Path.cwd() / pdf_target).resolve()
                ok, msg = _excel_original_report_action(
                    data.path,
                    int(r_year),
                    int(r_month),
                    r_project,
                    action="pdf",
                    pdf_output_path=pdf_target,
                    report_df=report_df,
                )
                if ok:
                    st.success(msg)
                    try:
                        page_count = max(1, math.ceil(len(report_df) / REPORT_ROWS_PER_PAGE)) if report_df is not None else 1
                        if page_count == 1 and pdf_target.exists():
                            st.download_button(
                                "Exportierte PDF herunterladen",
                                data=Path(pdf_target).read_bytes(),
                                file_name=pdf_target.name,
                                mime="application/pdf",
                                key="download_exported_pdf",
                            )
                        elif page_count > 1:
                            st.caption("Mehrere PDFs wurden erzeugt und im Exportordner durchnummeriert abgelegt (z. B. _01, _02, ...).")
                    except Exception:
                        pass
                else:
                    st.error(msg)

        with b3:
            if st.button("Original als Excel-Kopie", key="export_original_xlsx_copy"):
                xlsx_target = Path(xlsx_rel_or_abs).expanduser()
                if not xlsx_target.is_absolute():
                    xlsx_target = (Path.cwd() / xlsx_target).resolve()
                ok, msg = _excel_original_report_action(
                    data.path,
                    int(r_year),
                    int(r_month),
                    r_project,
                    action="xlsx",
                    xlsx_output_path=xlsx_target,
                    report_df=report_df,
                )
                if ok:
                    st.success(msg)
                    try:
                        page_count = max(1, math.ceil(len(report_df) / REPORT_ROWS_PER_PAGE)) if report_df is not None else 1
                        if page_count == 1 and xlsx_target.exists():
                            st.download_button(
                                "Exportierte Excel-Kopie herunterladen",
                                data=Path(xlsx_target).read_bytes(),
                                file_name=xlsx_target.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_exported_original_xlsx_copy",
                            )
                        elif page_count > 1:
                            st.caption("Mehrere Excel-Kopien wurden erzeugt und im Exportordner durchnummeriert abgelegt (z. B. _01, _02, ...).")
                    except Exception:
                        pass
                else:
                    st.error(msg)

        with b4:
            if st.button("Original direkt drucken (Excel)", key="print_original_excel"):
                ok, msg = _excel_original_report_action(
                    data.path,
                    int(r_year),
                    int(r_month),
                    r_project,
                    action="print",
                    report_df=report_df,
                )
                (st.success if ok else st.error)(msg)
        if report_df.empty:
            st.warning("Keine Einträge für diesen Einsatzbericht gefunden.")
            st.caption("Tipp: 'abgerechnete einschließen' aktivieren, falls die Beispiel-Datei nur bereits markierte Einträge enthält.")
        else:
            st.dataframe(report_df.drop(columns=["_excel_row"]), use_container_width=True, height=360)
            csv_bytes = report_df.drop(columns=["_excel_row"]).to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "CSV exportieren",
                data=csv_bytes,
                file_name=f"Einsatzbericht_{r_project}_{r_year}-{int(r_month):02d}.csv",
                mime="text/csv",
            )
            with st.expander("Treffer als abgerechnet markieren"):
                st.caption("Schreibt 'ja' in Spalte 'Abgerechnet' für die aktuell im Report enthaltenen Tätigkeiten.")
                if st.button("Report-Treffer markieren", key="mark_report_done"):
                    rows_to_mark = [int(r) for r in report_df["_excel_row"].tolist()]

                    def _mutator_mark_report(wb):
                        ws = wb[TAETIGKEITEN_SHEET]
                        for r in rows_to_mark:
                            ws.cell(r, 13).value = "ja"

                    ok, msg = _save_workbook(data.path, _mutator_mark_report)
                    if ok:
                        st.success(f"{len(rows_to_mark)} Einträge markiert. {msg}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(msg)
    with tab3:
        st.subheader("Stammdaten / Debug")
        st.write(f"**Datei:** `{data.path}`")
        st.write(f"**Tätigkeiten (erkannte Datensätze):** {len(df)}")

        c1, c2, c3 = st.columns(3)
        c1.write("**Projekte**")
        c1.dataframe(pd.DataFrame({"Projekt": lookups.get("projekte", [])}), use_container_width=True, height=200)
        c2.write("**Tätigkeitstypen**")
        c2.dataframe(pd.DataFrame({"Typ": lookups.get("taetigkeit_typen", [])}), use_container_width=True, height=200)
        c3.write("**Relevante Kodierungen (aktiv)**")
        c3.dataframe(pd.DataFrame({"Kodierung": lookups.get("relevante_kodierungen", [])}), use_container_width=True, height=200)

        st.markdown("---")
        st.markdown("### Relevante Kodierungen pflegen")
        all_kod_aufgaben = lookups.get("kodierung_aufgaben", []) or []
        if all_kod_aufgaben:
            with st.form("relevante_kodierungen_form", clear_on_submit=False):
                selected_relevante = st.multiselect(
                    "Wähle die abrechnungs-/auswahltrelevanten Kodierungen (Aufgabe)",
                    options=all_kod_aufgaben,
                    default=[k for k in (lookups.get("relevante_kodierungen", []) or []) if k in all_kod_aufgaben],
                    help="Speichert die Auswahl als 'x' in 'Kodierung Joyson' Spalte A.",
                )
                save_relevante = st.form_submit_button("Relevante Kodierungen speichern")
            if save_relevante:
                def _mutator_relevante(wb):
                    ok_inner, msg_inner = _set_relevante_kodierungen(wb, selected_relevante)
                    if not ok_inner:
                        raise ValueError(msg_inner)
                ok, msg = _save_workbook(data.path, _mutator_relevante)
                if ok:
                    st.success(msg)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(msg)
        else:
            st.info("Keine Kodierungen aus 'Kodierung Joyson' erkannt.")

        st.markdown("---")
        st.markdown("### Projekt-Stammdaten pflegen")
        projekt_infos = lookups.get("projekt_infos", {}) or {}
        projekte_sortiert = sorted(projekt_infos.keys())
        proj_select = st.selectbox(
            "Projekt auswählen",
            options=["<Neues Projekt>"] + projekte_sortiert,
            index=0,
            key="proj_master_select",
        )

        if proj_select == "<Neues Projekt>":
            proj_defaults = {
                "Projekt": "",
                "Kunde": "",
                "Straße": "",
                "Ort": "",
                "Ansprechpartner": "",
                "Projektadresse Standard": "",
                "Projektadresse Alternativ": "",
            }
            original_project_name = ""
        else:
            proj_defaults = dict(projekt_infos.get(proj_select, {}))
            original_project_name = proj_select

        with st.form("projekt_stammdaten_form", clear_on_submit=False):
            pc1, pc2 = st.columns(2)
            with pc1:
                p_name = st.text_input("Projekt (Kürzel/Name)", value=_safe_str(proj_defaults.get("Projekt")))
                p_kunde = st.text_input("Kunde", value=_safe_str(proj_defaults.get("Kunde")))
                p_ansp = st.text_input("Ansprechpartner", value=_safe_str(proj_defaults.get("Ansprechpartner")))
                p_addr_std = st.text_input("Projektadresse Standard", value=_safe_str(proj_defaults.get("Projektadresse Standard")))
            with pc2:
                p_str = st.text_input("Straße", value=_safe_str(proj_defaults.get("Straße")))
                p_ort = st.text_input("Ort", value=_safe_str(proj_defaults.get("Ort")))
                p_addr_alt = st.text_input("Projektadresse Alternativ", value=_safe_str(proj_defaults.get("Projektadresse Alternativ")))
                rename_taet = st.checkbox(
                    "Bei Projekt-Umbenennung auch Tätigkeiten aktualisieren",
                    value=False,
                    disabled=(not original_project_name),
                    help="Ändert Spalte 'Projekt' in 'Tätigkeiten' von altem auf neuen Projektnamen.",
                )

            save_proj = st.form_submit_button("Projekt-Stammdaten speichern")

        if save_proj:
            payload = {
                "Projekt": p_name,
                "Kunde": p_kunde,
                "Straße": p_str,
                "Ort": p_ort,
                "Ansprechpartner": p_ansp,
                "Projektadresse Standard": p_addr_std,
                "Projektadresse Alternativ": p_addr_alt,
            }

            def _mutator_proj(wb):
                ok_inner, msg_inner = _upsert_project_stammdaten(
                    wb,
                    payload,
                    original_project=original_project_name or None,
                    rename_taetigkeiten=bool(rename_taet),
                )
                if not ok_inner:
                    raise ValueError(msg_inner)

            ok, msg = _save_workbook(data.path, _mutator_proj)
            if ok:
                st.success(msg)
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(msg)

        with st.expander("Projekt-Stammdaten (Tabelle)"):
            proj_df = pd.DataFrame.from_dict(lookups.get("projekt_infos", {}), orient="index").reset_index(drop=True)
            st.dataframe(proj_df, use_container_width=True, height=300)

        with st.expander("Kodierung-Mapping (Aufgabe -> Kodierung EB)"):
            km = lookups.get("kodierung_map_eb", {})
            km_df = pd.DataFrame([{"Aufgabe": k, "Kodierung EB": v} for k, v in km.items()])
            st.dataframe(km_df, use_container_width=True, height=300)


if __name__ == "__main__":
    main()
