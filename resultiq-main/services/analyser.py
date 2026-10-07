"""
services/analyser.py
====================
Intelligent, content-driven workbook analyser for RESULTIQ — Module 5.

Pure Python / pandas only.  Zero Flask or SQLAlchemy imports so the module
can be unit-tested in complete isolation and reused by Module 6 without any
framework coupling.

Public API
----------
analyse_workbook(filepath)  ->  list[dict]
    Open an Excel workbook and return a fully-enriched list of sheet dicts.
    Each column dict is extended with 'classification' and 'quality' keys.
    The returned structure is a drop-in replacement for the 'worksheets' JSON
    already stored in UploadedFile.worksheets by Module 4.
"""

from __future__ import annotations

import re
import pandas as pd
import openpyxl


# ------------------------------------------------------------------ #
# Column-name patterns that should always be classified as 'identity'
# regardless of data content (serial numbers, roll numbers, names, IDs).
# ------------------------------------------------------------------ #
_IDENTITY_NAME_PATTERNS = re.compile(
    r'^\s*('
    r'sl\.?\s*no\.?|s\.?no\.?|sr\.?\s*no\.?|serial(\s*(no\.?|number))?|index'
    r'|roll(\s*(no\.?|number))?|roll_no|rollno|rno'
    r'|reg\.?\s*no\.?|reg(ister)?\s*(no\.?|number)|registration\s*(no\.?|number)?|regiter\s*number|regd\.?\s*no\.?'
    r'|usn|urn|prn|ht\.?\s*no\.?|hall\s*ticket(\s*(no\.?|number))?|seat(\s*(no\.?|number))?'
    r'|admission\s*(no\.?|number)?|enroll(ment)?\s*(no\.?|number)?'
    r'|((student|candidate|user|respondent|employee|emp|applicant)\s*)?id(\s*(no\.?|number))?'
    r'|((student|candidate|pupil)\s*)?name(\s*of\s*(the\s*)?(student|candidate))?|first\s*name|last\s*name|full\s*name|student_name'
    r'|email(\s*id)?|phone(\s*(no\.?|number))?|mobile(\s*(no\.?|number))?|contact(\s*(no\.?|number))?'
    r'|aadhaar(\s*(no\.?|number))?'
    r')\s*$',
    re.IGNORECASE,
)

# ------------------------------------------------------------------ #
# Summary, computed, feedback, or administrative column names that
# must NEVER be classified as marks (e.g. totals, averages, CGPA, rank).
# ------------------------------------------------------------------ #
_EXCLUDE_FROM_MARKS_PATTERNS = re.compile(
    r'^\s*('
    r'total(\s*marks?)?|grand\s*total|tot|max(\.?|\s*)marks?|maximum\s*marks?'
    r'|percentage|percent|%|avg|average|overall\s*%?'
    r'|rank|grade|gpa|cgpa|sgpa|credits?|credit\s*points?'
    r'|attendance|attendance\s*%?|attended|present|absent\s*(count|days)?'
    r'|result|status|remarks?|pass\s*\/?\s*fail|outcome|division'
    r'|rating(\s*\(.*\))?|feedback|response(\s*text)?|comments?|review'
    r'|section|sec|branch|dept|department|degree|semester|sem|year|batch|group|category|age\s*group'
    r')\s*$',
    re.IGNORECASE,
)

_CATEGORY_NAME_HINTS = re.compile(
    r'^\s*(grade|result|status|remarks?|pass\s*\/?\s*fail|outcome|division|'
    r'section|sec|branch|dept|department|degree|semester|sem|year|batch|group|category|gender|sex|age\s*group)\s*$',
    re.IGNORECASE,
)

# ------------------------------------------------------------------ #
# Subject / Exam / Test name patterns that strongly signal a marks column
# ------------------------------------------------------------------ #
_SUBJECT_NAME_PATTERNS = re.compile(
    r'^\s*('
    r'math|maths|mathematics|physics|chemistry|biology|science|social|history|geography'
    r'|english|tamil|hindi|sanskrit|french|german|language'
    r'|computer|programming|coding|python|java|c\+\+|dbms|oops?|os|operating\s*system|network|dsa|data\s*structures'
    r'|economics|commerce|account(s|ing|ancy)?|business|statistics|stats'
    r'|drawing|art|music|pe|physical\s*education'
    r'|cit(-|\s*)?[0-9]*|cat(-|\s*)?[0-9]*|internal(-|\s*)?[0-9]*'
    r'|mid(-|\s*)?(term|sem)[0-9]*|end(-|\s*)?(term|sem)[0-9]*|unit(-|\s*)?test[0-9]*|test(-|\s*)?[0-9]*'
    r'|quiz(-|\s*)?[0-9]*|assignment(-|\s*)?[0-9]*|theory|practical|lab|viva'
    r'|ia(-|\s*)?[0-9]*|mse(-|\s*)?[0-9]*|ese(-|\s*)?[0-9]*|paper(-|\s*)?[0-9]*'
    r'|sub(-|\s*)?[0-9]*|subject(-|\s*)?[0-9]*|score[0-9]*|mark(s)?[0-9]*'
    r'|[0-9]{1,2}[a-z]{2,4}[0-9]{2,4}[a-z]?\.?'    # e.g. 24CS301, 24BS303.
    r'|[a-z]{2,4}[0-9]{2,4}[a-z]?\.?'              # e.g. CS101, IT301, MA8351
    r')\s*$',
    re.IGNORECASE,
)

# Footer / Signature / Non-data row pattern
FOOTER_REGEX = re.compile(
    r'(faculty|incharge|in-charge|h\.?o\.?d|head of|signature|staff|principal|dean|'
    r'prepared by|verified by|advisor|coordinator|controller|result analysis|total|average|passed|appeared)',
    re.IGNORECASE,
)

# ------------------------------------------------------------------ #
# Absent / non-numeric tokens that commonly appear in marks sheets
# (case-insensitive, stripped).  These are *not* treated as errors.
# ------------------------------------------------------------------ #
ABSENT_TOKENS = frozenset({
    'aa', 'ab', 'abs', 'absent', 'a', 'na', 'n/a', 'n.a', 'n.a.', '-', '--', '---',
    'nil', 'null', 'none', 'not appeared', 'withheld', 'w/h', 'wh',
    'medical', 'detained', 'ex', 'exempted', 'od', 'on duty', 'leave', 'ml', 'ne',
})


# ------------------------------------------------------------------ #
# Helpers
# ------------------------------------------------------------------ #

def _is_absent_token(value: str) -> bool:
    """Return True if a string represents a known absent/non-participation marker."""
    return value.strip().lower() in ABSENT_TOKENS


def _clean_numeric_str(value) -> str | None:
    """
    Clean and normalize potential numeric strings.
    Strips commas, whitespace, and common academic grace/note characters (*, #, @, $, %).
    Returns cleaned string, or None if empty/NaN.
    """
    if value is None or pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None
    # If it is a known absent token, do not strip it
    if _is_absent_token(s):
        return None
    # Strip commas and trailing academic symbols like 45*, 60#, 50@
    s = s.replace(',', '')
    s = re.sub(r'[*#@$%\s]', '', s)
    return s if s else None


def _try_numeric(value) -> bool:
    """Return True if *value* can be interpreted as a number."""
    s = _clean_numeric_str(value)
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def _coerce_numeric_series(series: pd.Series):
    """
    Coerce a Series to numeric, treating absent tokens as NaN.
    Returns numeric_series with floats or NaN.
    """
    def _convert(v):
        s = _clean_numeric_str(v)
        if not s:
            return float('nan')
        try:
            return float(s)
        except ValueError:
            return float('nan')

    return series.map(_convert)


# ------------------------------------------------------------------ #
# Column classification
# ------------------------------------------------------------------ #

_MARKS_NUMERIC_THRESHOLD = 0.50
_MARKS_COVERAGE_THRESHOLD = 0.75
_IDENTITY_UNIQUENESS_THRESHOLD = 0.85
_CATEGORY_MAX_UNIQUE = 20
_EMPTY_THRESHOLD = 0.05


def _build_quality(series: pd.Series) -> dict:
    """
    Compute data-quality metrics for a single column Series
    (the full column including blanks).
    """
    total = len(series)
    null_mask = series.isna() | (series.astype(str).str.strip() == '')
    filled = int((~null_mask).sum())
    missing = int(null_mask.sum())

    non_null = series[~null_mask]

    numeric_count = int(non_null.apply(_try_numeric).sum())
    absent_count  = int(non_null.apply(
        lambda v: _is_absent_token(str(v))
    ).sum())
    text_count    = filled - numeric_count - absent_count

    try:
        unique_count = int(non_null.nunique())
    except Exception:
        unique_count = 0

    return {
        'total':               total,
        'filled':              filled,
        'missing':             missing,
        'numeric_count':       numeric_count,
        'absent_token_count':  absent_count,
        'text_count':          text_count,
        'unique_count':        unique_count,
    }


def _infer_dtype_label(series: pd.Series) -> str:
    """Human-readable dtype label (mirrors Module 4 logic for consistency)."""
    dtype = series.dtype
    if pd.api.types.is_integer_dtype(dtype):
        return 'Integer'
    if pd.api.types.is_float_dtype(dtype):
        return 'Decimal'
    if pd.api.types.is_bool_dtype(dtype):
        return 'Boolean'
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return 'Date/Time'
    return 'Text'


def classify_column(series: pd.Series, col_name: str) -> tuple[str, dict]:
    """
    Classify a single column with high precision by inspecting both its name
    and actual value distribution.

    Parameters
    ----------
    series   : full column including NaN / blanks
    col_name : column name

    Returns
    -------
    (classification_label, quality_dict)
    """
    quality = _build_quality(series)
    col_str = str(col_name).strip()

    # 1. Datetime / Date of Birth check -> metadata
    if pd.api.types.is_datetime64_any_dtype(series.dtype) or re.match(r'^\s*(dob|date(\s*of\s*birth)?|timestamp|time)\s*$', col_str, re.IGNORECASE):
        quality['suitable_for_analysis'] = False
        return 'metadata', quality

    # 2. Identity name guard
    # Columns matching roll numbers, registration numbers, student names, IDs, etc.
    if _IDENTITY_NAME_PATTERNS.match(col_str):
        quality['suitable_for_analysis'] = False
        return 'identity', quality

    # 2. Exclude non-subject / computed columns from marks
    # (Total Marks, Percentage, CGPA, Rank, Rating, Result, Grade, Section, etc.)
    if _EXCLUDE_FROM_MARKS_PATTERNS.match(col_str):
        quality['suitable_for_analysis'] = False
        if _CATEGORY_NAME_HINTS.match(col_str):
            return 'category', quality
        return 'metadata', quality

    # 3. Unnamed placeholder columns from multi-row headers
    if col_str.startswith('Unnamed:'):
        quality['suitable_for_analysis'] = False
        return 'metadata', quality

    # 4. Long descriptive column names are never marks
    if len(col_str) > 40:
        quality['suitable_for_analysis'] = False
        return 'metadata', quality

    filled = quality['filled']
    total  = quality['total']

    # 5. Empty check
    if total == 0 or (filled / total) < _EMPTY_THRESHOLD:
        quality['suitable_for_analysis'] = False
        return 'empty', quality

    numeric_count = quality['numeric_count']
    absent_count  = quality['absent_token_count']
    unique_count  = quality['unique_count']

    numeric_ratio  = numeric_count / filled if filled else 0
    coverage_ratio = (numeric_count + absent_count) / filled if filled else 0

    # 6. Extract clean numeric values for distribution inspection
    numeric_vals = None
    median_val = None
    if numeric_count > 0:
        cleaned_series = series.dropna().map(_clean_numeric_str)
        numeric_vals = pd.to_numeric(cleaned_series, errors='coerce').dropna()
        if len(numeric_vals) > 0:
            median_val = float(numeric_vals.median())

    # 7. Large-number guard (e.g. 9-digit registration numbers, phone numbers)
    if median_val is not None and median_val > 999:
        quality['suitable_for_analysis'] = False
        return 'identity', quality

    # 8. Rating / Likert guard (survey ratings strictly 1..5, or boolean 0..1)
    is_subject_name = bool(_SUBJECT_NAME_PATTERNS.match(col_str))
    if numeric_vals is not None and len(numeric_vals) > 0 and not is_subject_name:
        unique_vals = set(numeric_vals.unique())
        # Ratings strictly 1 to 5
        if unique_vals.issubset({1, 2, 3, 4, 5}) and 'rate' in col_str.lower():
            quality['suitable_for_analysis'] = False
            return 'metadata', quality
        # Years (e.g. 2020..2030)
        if all(1990 <= v <= 2040 for v in unique_vals) and ('year' in col_str.lower() or 'batch' in col_str.lower()):
            quality['suitable_for_analysis'] = False
            return 'metadata', quality

    # 9. Marks classification
    # High-confidence subject name with valid numeric and absent values
    if is_subject_name and coverage_ratio >= 0.65 and numeric_count >= 1:
        quality['suitable_for_analysis'] = True
        return 'marks', quality

    # Content-driven marks classification
    if (numeric_ratio >= _MARKS_NUMERIC_THRESHOLD and
            coverage_ratio >= _MARKS_COVERAGE_THRESHOLD):
        # Academic marks are within a realistic range [0, 200]
        if median_val is not None and median_val <= 200:
            quality['suitable_for_analysis'] = True
            return 'marks', quality

    # 10. Datetime check
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        quality['suitable_for_analysis'] = False
        return 'metadata', quality

    # 11. Identity: high uniqueness non-numeric
    if unique_count > 0 and filled > 0:
        uniqueness_ratio = unique_count / filled
        if uniqueness_ratio >= _IDENTITY_UNIQUENESS_THRESHOLD and numeric_ratio < 0.5:
            # If dataset is small (< 10 rows), don't classify arbitrary text as identity unless name hints it
            if filled >= 10 or 'name' in col_str.lower() or 'id' in col_str.lower():
                quality['suitable_for_analysis'] = False
                return 'identity', quality

    # 12. Category: low-cardinality text
    if unique_count <= _CATEGORY_MAX_UNIQUE and numeric_ratio < 0.5:
        quality['suitable_for_analysis'] = False
        return 'category', quality

    # 13. Default: text
    quality['suitable_for_analysis'] = False
    return 'text', quality


# ------------------------------------------------------------------ #
# Sheet & workbook analysis
# ------------------------------------------------------------------ #

def analyse_sheet(df: pd.DataFrame, sheet_name: str) -> dict:
    """
    Analyse every column in a DataFrame.

    Returns a sheet dict that is backwards-compatible with the Module 4
    format but enriches each column dict with 'classification' and 'quality'.
    """
    if df is None or df.empty or len(df.columns) == 0:
        return {
            'name':       sheet_name,
            'is_empty':   True,
            'row_count':  0,
            'columns':    [],
        }

    # Count real data rows by excluding trailing empty or footer rows
    real_data_rows = len(df)
    for idx, row in df.iterrows():
        row_vals = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip() != '']
        if row_vals and bool(FOOTER_REGEX.search(' '.join(row_vals))):
            real_data_rows = idx
            break

    columns = []
    for col in df.columns:
        col_str = str(col).strip()
        series = df[col]
        non_blank = series.dropna()
        if col_str.startswith('Unnamed') and non_blank.empty:
            continue

        classification, quality = classify_column(series, col_str)
        samples = [str(v) for v in non_blank.head(3).tolist()]

        columns.append({
            'name':           col_str,
            'dtype':          _infer_dtype_label(series),
            'samples':        samples,
            'classification': classification,
            'quality':        quality,
        })

    return {
        'name':      sheet_name,
        'is_empty':  len(columns) == 0,
        'row_count': real_data_rows,
        'columns':   columns,
    }


def _detect_header_row(filepath: str, sheet_name: str) -> int:
    """
    Detect the actual header row in a worksheet with high accuracy.
    Scores each candidate row across the first 35 rows based on header keywords,
    distinct short non-numeric column labels, and the presence of data in subsequent rows.
    """
    ext = filepath.rsplit('.', 1)[-1].lower()
    raw_rows = []

    if ext == 'xls':
        try:
            import xlrd
            rb = xlrd.open_workbook(filepath, formatting_info=False)
            sheet = rb.sheet_by_name(sheet_name)
            for row_idx in range(min(sheet.nrows, 35)):
                row_vals = []
                for c in range(sheet.ncols):
                    val = sheet.cell_value(row_idx, c)
                    row_vals.append(None if val == '' else val)
                raw_rows.append(row_vals)
        except Exception:
            return 0
    else:
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb[sheet_name]
            for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                raw_rows.append(list(row))
                if row_idx >= 35:
                    break
            wb.close()
        except Exception:
            return 0

    best_row = 0
    best_score = -1

    for row_idx, row in enumerate(raw_rows):
        text_cell_count = 0
        keyword_matches = 0
        total_non_blank = 0

        for val in row:
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            total_non_blank += 1
            if len(s) < 50 and not s.endswith(':'):
                try:
                    float(s.replace(',', ''))
                except ValueError:
                    text_cell_count += 1
                    if (_IDENTITY_NAME_PATTERNS.match(s) or
                            _SUBJECT_NAME_PATTERNS.match(s) or
                            any(k in s.lower() for k in ['name', 'roll', 'reg', 'sl', 'no', 'mark', 'score', 'sub', 'cit'])):
                        keyword_matches += 1

        if text_cell_count >= 2:
            # Score this row
            score = text_cell_count + (keyword_matches * 3)
            # Bonus if next row has data
            if row_idx + 1 < len(raw_rows):
                next_row = raw_rows[row_idx + 1]
                has_next_data = any(v is not None and str(v).strip() != '' for v in next_row)
                if has_next_data:
                    score += 2

            if score > best_score:
                best_score = score
                best_row = row_idx

    return best_row


def _read_sub_header_names(filepath: str, sheet_name: str,
                           header_row_0based: int, engine: str) -> dict[int, str]:
    """Read the row immediately after *header_row_0based* for sub-header names."""
    sub_row = header_row_0based + 1
    result: dict[int, str] = {}

    try:
        if engine == 'xlrd':
            import xlrd
            rb = xlrd.open_workbook(filepath, formatting_info=False)
            sheet = rb.sheet_by_name(sheet_name)
            if sub_row >= sheet.nrows:
                return result
            for c in range(sheet.ncols):
                val = sheet.cell_value(sub_row, c)
                if val is None or val == '':
                    continue
                s = str(val).strip()
                if not s or len(s) >= 40:
                    continue
                try:
                    float(s.replace(',', ''))
                except ValueError:
                    result[c] = s
        else:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb[sheet_name]
            ws_row_1based = sub_row + 1
            if ws_row_1based <= ws.max_row:
                for c_idx, cell in enumerate(list(ws.iter_rows(
                        min_row=ws_row_1based, max_row=ws_row_1based,
                        values_only=True))[0]):
                    if cell is None:
                        continue
                    s = str(cell).strip()
                    if not s or len(s) >= 40:
                        continue
                    try:
                        float(s.replace(',', ''))
                    except ValueError:
                        result[c_idx] = s
            wb.close()
    except Exception:
        return {}

    return result


def analyse_workbook(filepath: str) -> list[dict]:
    """
    Open an Excel workbook and return a fully-enriched list of sheet dicts.
    """
    ext = filepath.rsplit('.', 1)[-1].lower()
    try:
        engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'
        xf = pd.ExcelFile(filepath, engine=engine)
    except Exception as exc:
        raise ValueError(f"Cannot open workbook: {exc}") from exc

    sheets = []
    try:
        for sheet_name in xf.sheet_names:
            try:
                header_row = _detect_header_row(filepath, sheet_name)
                df = xf.parse(sheet_name, header=header_row)

                # Multi-row header resolution
                unnamed_cols = [c for c in df.columns if str(c).startswith('Unnamed:')]
                if unnamed_cols:
                    sub_header_names = _read_sub_header_names(
                        filepath, sheet_name, header_row, engine
                    )
                    if sub_header_names:
                        new_cols = list(df.columns)
                        for i, col_name in enumerate(new_cols):
                            if i in sub_header_names:
                                col_str = str(col_name)
                                if (col_str.startswith('Unnamed:')
                                        or len(col_str.strip()) > 40):
                                    new_cols[i] = sub_header_names[i]
                        df.columns = new_cols
                        if len(df) > 0:
                            first_row = df.iloc[0]
                            is_subheader_row = all(
                                str(first_row.iloc[i]).strip() == sub_header_names[i]
                                for i in sub_header_names
                                if i < len(first_row)
                            )
                            if is_subheader_row:
                                df = df.iloc[1:].reset_index(drop=True)

            except Exception:
                sheets.append({
                    'name':      sheet_name,
                    'is_empty':  True,
                    'row_count': 0,
                    'columns':   [],
                    'header_row_index': 0,
                })
                continue
            sheet_dict = analyse_sheet(df, sheet_name)
            sheet_dict['header_row_index'] = header_row
            sheets.append(sheet_dict)
    finally:
        xf.close()

    return sheets
