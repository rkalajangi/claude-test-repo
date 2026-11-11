#!/usr/bin/env python3
"""
INT5016 ORP/FRS fixed-width file -> validation + Excel workbook.

Reads the text file produced by INT5016_DT_v2.xsl, checks every record
against the layout in that stylesheet, and writes an .xlsx with:

    Summary      - overall PASS/FAIL, counts, control totals
    Detail       - one row per detail record, decoded (amounts, dates)
    Validation   - every error/warning with line, field and position
    Header_Trailer, Layout, Raw

Usage:
    python int5016_to_excel.py INPUT [INPUT ...] [-o OUTPUT_DIR] [--pattern "*.txt"]

INPUT can be a file or a folder. For a folder every data file is processed
(any extension, including none; .xlsx/.py/.bat/.xml etc. are skipped).

Exit codes: 0 = no errors (warnings allowed), 2 = validation errors found,
            1 = usage problem / file could not be processed.

Requires: openpyxl  (pip install openpyxl)
"""

import argparse
import datetime as dt
import glob
import os
import re
import sys
from collections import Counter

try:
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("ERROR: openpyxl is not installed. Run:  python -m pip install openpyxl\n")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Layout (mirrors INT5016_DT_v2.xsl exactly)
# ---------------------------------------------------------------------------
# kind:
#   text   - left aligned, space padded            (mf:pad-right)
#   zcode  - right aligned, zero padded code       (mf:pad-left '0')
#   amount - implied 2 decimals, zero padded       (mf:format-currency)
#   date8  - MMDDYYYY or spaces                    (mf:format-date-8)
#   date6  - MMYYYY or spaces                      (mf:format-date-6)
#   ssn / state / zip / email / flag - text with extra checks
# required: value must not be blank

HEADER_LAYOUT = [
    ("Agency_Number", 5, "text", True),
    ("Record_Type", 9, "text", True),        # '      NHD'
    ("Report_Period", 6, "date6", True),
]

TRAILER_LAYOUT = [
    ("Agency_Number", 5, "text", True),
    ("Record_Type", 9, "text", True),        # '       TR'
    ("Report_Period", 6, "date6", True),
    ("Total_Records", 6, "count", True),
]

DETAIL_LAYOUT = [
    ("Agency_Number", 5, "text", True),
    ("Social_Security_Number", 9, "ssn", True),
    ("Check_Date", 6, "date6", True),
    ("Name", 31, "text", True),
    ("Plan", 2, "text", False),
    ("Work_Period", 2, "text", False),
    ("Adjustment_Code", 1, "flag", False),
    ("Salary", 11, "amount", False),
    ("Employee_Contributions", 11, "amount", False),
    ("Employer_Contributions", 11, "amount", False),
    ("Annual_Leave_Hours", 6, "amount", False),
    ("Department_Number", 12, "zcode", False),
    ("Sex_Code", 1, "flag", False),
    ("Date_of_Birth", 8, "date8", False),
    ("Educational_Employee_Indicator", 1, "flag", False),
    ("Class_Code", 5, "zcode", False),
    ("Position_Number", 10, "zcode", False),
    ("Address_Line_1", 31, "text", False),
    ("Address_Line_2", 31, "text", False),
    ("Out_of_Country_Address_Line", 31, "text", False),
    ("City", 25, "text", False),
    ("State", 2, "state", False),
    ("Zip", 9, "zip", False),
    ("Date_of_Hire", 8, "date8", False),
    ("Date_of_Termination", 8, "date8", False),
    ("Annualized_Salary", 11, "amount", False),
    ("Eligible_Compensation_415", 11, "amount", False),
    ("Plan_Contribution_457", 11, "amount", False),
    ("Catch_Up_Flag_457", 1, "flag", False),
    ("Unused_457_Catch_Up_Contributions", 11, "amount", False),
    ("Includible_Compensation_403b", 11, "amount", False),
    ("Employee_Contribution_403b", 11, "amount", False),
    ("Employer_Contribution_403b", 11, "amount", False),
    ("Catch_Up_Flag_403b", 1, "flag", False),
    ("Catch_Up_Contributions_403b", 11, "amount", False),
    ("Employee_Contributions_401k", 11, "amount", False),
    ("Employer_Contributions_401k", 11, "amount", False),
    ("Employer_Contributions_401a", 11, "amount", False),
    ("Public_Records_Exemption_Indicator", 1, "flag", False),
    ("Mandatory_ORP_Indicator", 1, "flag", False),
    ("Email_Address", 60, "email", False),
    ("Fund_Code", 15, "text", False),
    ("Mobile_Phone_Number", 9, "text", False),
]


def with_positions(layout):
    """Add 1-based start/end positions to a layout."""
    out, pos = [], 1
    for name, length, kind, required in layout:
        out.append({"name": name, "len": length, "kind": kind, "required": required,
                    "start": pos, "end": pos + length - 1})
        pos += length
    return out, pos - 1


HEADER, HEADER_LEN = with_positions(HEADER_LAYOUT)      # 20
TRAILER, TRAILER_LEN = with_positions(TRAILER_LAYOUT)   # 26
DETAIL, DETAIL_LEN = with_positions(DETAIL_LAYOUT)      # 486

ERROR, WARNING = "ERROR", "WARNING"
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ---------------------------------------------------------------------------
# Field decoding / validation
# ---------------------------------------------------------------------------
def parse_field(fld, raw):
    """Return (value_for_excel, number_format, [(severity, message), ...])."""
    kind, issues = fld["kind"], []
    s = raw.strip()

    if fld["required"] and s == "":
        issues.append((ERROR, "Required field is blank"))
        return "", None, issues

    if kind == "amount":
        if raw.strip("0") == "" or raw.isdigit():
            return int(raw) / 100.0, "#,##0.00", issues
        if "-" in raw:
            # mf:pad-left puts the zeros in front of the sign: 00000-12345
            digits = raw.replace("-", "")
            if digits.isdigit() and raw.count("-") == 1:
                issues.append((ERROR, "Negative amount - XSLT zero-pads in front of the '-' sign "
                                      "(e.g. 00000-12345); receiving system will likely reject it"))
                return -int(digits) / 100.0, "#,##0.00", issues
        issues.append((ERROR, "Amount must be all digits (implied 2 decimals)"))
        return raw, None, issues

    if kind == "count":
        if not raw.isdigit():
            issues.append((ERROR, "Record count must be all digits"))
            return raw, None, issues
        return int(raw), "0", issues

    if kind in ("date8", "date6"):
        if s == "":
            return "", None, issues
        try:
            if kind == "date8":
                if not raw.isdigit():
                    raise ValueError
                return dt.datetime.strptime(raw, "%m%d%Y"), "mm/dd/yyyy", issues
            if not raw.isdigit():
                raise ValueError
            return dt.datetime.strptime(raw, "%m%Y"), "mm/yyyy", issues
        except ValueError:
            fmt = "MMDDYYYY" if kind == "date8" else "MMYYYY"
            issues.append((ERROR, "Invalid date, expected %s or spaces" % fmt))
            return raw, None, issues

    if kind == "zcode":
        if " " in raw:
            issues.append((ERROR, "Zero-padded code contains spaces"))
        elif not raw.isdigit():
            issues.append((WARNING, "Code is not numeric"))
        return raw, "@", issues

    if kind == "ssn":
        if not raw.isdigit():
            issues.append((ERROR, "SSN must be exactly 9 digits"))
        elif raw in ("000000000", "111111111", "123456789", "999999999") or raw[:3] in ("000", "666") \
                or raw[3:5] == "00" or raw[5:] == "0000":
            issues.append((WARNING, "SSN looks invalid/placeholder"))
        return raw, "@", issues

    if kind == "state":
        if s and not (len(s) == 2 and s.isalpha() and s.isupper()):
            issues.append((WARNING, "State should be a 2-letter upper-case code"))
        return s, None, issues

    if kind == "zip":
        if s and not (s.isdigit() and len(s) in (5, 9)):
            issues.append((WARNING, "Zip should be 5 or 9 digits"))
        return s, "@", issues

    if kind == "email":
        if s and not EMAIL_RE.match(s):
            issues.append((WARNING, "Email address format looks invalid"))
        return s, None, issues

    if kind == "text" and len(raw) > 1 and raw[-1] != " " and len(s) == fld["len"] \
            and fld["name"] in ("Name", "Address_Line_1", "Address_Line_2", "City", "Email_Address"):
        issues.append((WARNING, "Field is completely filled - value may have been truncated"))

    # text / flag
    return s, "@" if s.isdigit() else None, issues


def slice_record(line, layout):
    return [(f, line[f["start"] - 1:f["end"]].ljust(f["len"])) for f in layout]


# ---------------------------------------------------------------------------
# File processing
# ---------------------------------------------------------------------------
def read_lines(path, findings):
    with open(path, "rb") as fh:
        data = fh.read()
    if not data.strip():
        findings.append((0, "FILE", "", None, None, "", ERROR, "File is empty"))
        return []
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252", errors="replace")
        findings.append((0, "FILE", "", None, None, "", WARNING,
                         "File is not valid UTF-8 (the XSLT writes UTF-8); decoded as Windows-1252"))

    if "\r\n" not in text and "\n" in text:
        findings.append((0, "FILE", "", None, None, "", WARNING,
                         "Line endings are LF only; the XSLT writes CRLF"))
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    while lines and lines[-1] == "":
        lines.pop()
    return lines


def classify(line):
    rec = line[5:14].strip()
    if rec == "NHD" and len(line.rstrip()) <= HEADER_LEN:
        return "HEADER"
    if rec == "TR" and len(line.rstrip()) <= TRAILER_LEN:
        return "TRAILER"
    return "DETAIL"


def process_file(path):
    findings = []   # (line_no, rec_type, field, start, end, raw, severity, message)
    lines = read_lines(path, findings)

    def add(line_no, rtype, fld, raw, sev, msg):
        findings.append((line_no, rtype, fld["name"] if fld else "",
                         fld["start"] if fld else None, fld["end"] if fld else None,
                         raw, sev, msg))

    header = trailer = None
    details = []          # dict: line_no, raw, values{name: (value, fmt, sev)}, status
    headers_seen = trailers_seen = 0

    for idx, line in enumerate(lines, start=1):
        if line.strip() == "":
            add(idx, "BLANK", None, "", ERROR, "Blank line inside file")
            continue
        rtype = classify(line)
        layout, exp_len = {"HEADER": (HEADER, HEADER_LEN), "TRAILER": (TRAILER, TRAILER_LEN),
                           "DETAIL": (DETAIL, DETAIL_LEN)}[rtype]

        if len(line) != exp_len:
            add(idx, rtype, None, "", ERROR,
                "Record length is %d, expected %d" % (len(line), exp_len))
        if any(ord(c) > 127 for c in line):
            add(idx, rtype, None, "", WARNING,
                "Non-ASCII characters: line is %d chars but %d bytes in UTF-8 - "
                "byte-position readers will see shifted fields"
                % (len(line), len(line.encode("utf-8"))))

        values, worst = {}, None
        for fld, raw in slice_record(line, layout):
            val, fmt, issues = parse_field(fld, raw)
            sev = None
            for s, m in issues:
                add(idx, rtype, fld, raw, s, m)
                sev = ERROR if ERROR in (s, sev) else s
            values[fld["name"]] = (val, fmt, sev)
            if sev == ERROR or (sev == WARNING and worst is None):
                worst = sev
        rec = {"line_no": idx, "raw": line, "values": values, "type": rtype}

        if rtype == "HEADER":
            headers_seen += 1
            if idx != 1:
                add(idx, rtype, None, "", ERROR, "Header record must be the first line")
            if line[5:14] != "      NHD":
                add(idx, rtype, HEADER[1], line[5:14], ERROR, "Record type must be 'NHD' right-aligned in 9")
            header = header or rec
        elif rtype == "TRAILER":
            trailers_seen += 1
            if idx != len(lines):
                add(idx, rtype, None, "", ERROR, "Trailer record must be the last line")
            if line[5:14] != "       TR":
                add(idx, rtype, TRAILER[1], line[5:14], ERROR, "Record type must be 'TR' right-aligned in 9")
            trailer = trailer or rec
        else:
            details.append(rec)
        rec["worst"] = worst

    # ---- file level checks ---------------------------------------------
    if headers_seen == 0:
        add(0, "FILE", None, "", ERROR, "No header (NHD) record found")
    elif headers_seen > 1:
        add(0, "FILE", None, "", ERROR, "%d header records found, expected 1" % headers_seen)
    if trailers_seen == 0:
        add(0, "FILE", None, "", ERROR, "No trailer (TR) record found")
    elif trailers_seen > 1:
        add(0, "FILE", None, "", ERROR, "%d trailer records found, expected 1" % trailers_seen)
    if not details:
        add(0, "FILE", None, "", WARNING, "File contains no detail records")

    h_agency = header["values"]["Agency_Number"][0] if header else None
    h_period = header["values"]["Report_Period"][0] if header else None

    if trailer:
        t = trailer["values"]
        if header and t["Agency_Number"][0] != h_agency:
            add(trailer["line_no"], "TRAILER", TRAILER[0], trailer["raw"][:5], ERROR,
                "Trailer agency does not match header agency '%s'" % h_agency)
        if header and t["Report_Period"][0] != h_period:
            add(trailer["line_no"], "TRAILER", TRAILER[2], trailer["raw"][14:20], ERROR,
                "Trailer report period does not match header")
        cnt = t["Total_Records"][0]
        if isinstance(cnt, int) and cnt != len(details):
            add(trailer["line_no"], "TRAILER", TRAILER[3], trailer["raw"][20:26], ERROR,
                "Trailer count %d does not equal number of detail records %d. Note: the XSLT "
                "adds one per EE_ORP_OP result line but never writes a separate ORP record"
                % (cnt, len(details)))

    # ---- detail cross-record / business checks --------------------------
    ssn_count = Counter(d["values"]["Social_Security_Number"][0] for d in details)
    pos = {f["name"]: f for f in DETAIL}
    today = dt.datetime.now()

    def flag(d, name, sev, msg):
        f = pos[name]
        add(d["line_no"], "DETAIL", f, d["raw"][f["start"] - 1:f["end"]], sev, msg)
        val, fmt, old = d["values"][name]
        d["values"][name] = (val, fmt, ERROR if ERROR in (sev, old) else sev)
        if sev == ERROR or d["worst"] is None:
            d["worst"] = sev

    for d in details:
        v = {k: x[0] for k, x in d["values"].items()}
        if h_agency is not None and v["Agency_Number"] != h_agency:
            flag(d, "Agency_Number", ERROR, "Agency does not match header agency '%s'" % h_agency)
        if ssn_count.get(v["Social_Security_Number"], 0) > 1 and v["Social_Security_Number"]:
            flag(d, "Social_Security_Number", WARNING,
                 "SSN appears %d times in file" % ssn_count[v["Social_Security_Number"]])
        if isinstance(h_period, dt.datetime) and isinstance(v["Check_Date"], dt.datetime) \
                and v["Check_Date"] != h_period:
            flag(d, "Check_Date", WARNING, "Check date period differs from header report period")
        dob, hire, term = v["Date_of_Birth"], v["Date_of_Hire"], v["Date_of_Termination"]
        if isinstance(dob, dt.datetime):
            if dob > today:
                flag(d, "Date_of_Birth", ERROR, "Date of birth is in the future")
            elif (today - dob).days < 14 * 365:
                flag(d, "Date_of_Birth", WARNING, "Worker would be younger than 14")
        if isinstance(hire, dt.datetime) and isinstance(dob, dt.datetime) and hire <= dob:
            flag(d, "Date_of_Hire", ERROR, "Hire date is on/before date of birth")
        if isinstance(hire, dt.datetime) and isinstance(term, dt.datetime) and term < hire:
            flag(d, "Date_of_Termination", ERROR, "Termination date is before hire date")
        if v["Out_of_Country_Address_Line"] == "" and v["Address_Line_1"] == "":
            flag(d, "Address_Line_1", WARNING, "No address supplied")

    return {"path": path, "lines": lines, "header": header, "trailer": trailer,
            "details": details, "findings": findings}


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="305496")
ERR_FILL = PatternFill("solid", fgColor="F8CBAD")
WRN_FILL = PatternFill("solid", fgColor="FFE699")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
MONO = Font(name="Consolas", size=10)


def style_header_row(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def autosize(ws, max_width=60):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            v = cell.value
            n = 10 if isinstance(v, (dt.datetime, dt.date)) else len(str(v))
            widths[cell.column] = max(widths.get(cell.column, 0), n)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(w + 2, 8), max_width)


def write_record_cells(ws, row, start_col, values, layout):
    for i, f in enumerate(layout):
        val, fmt, sev = values[f["name"]]
        cell = ws.cell(row=row, column=start_col + i, value=val if val != "" else None)
        if fmt:
            cell.number_format = fmt
        if sev == ERROR:
            cell.fill = ERR_FILL
        elif sev == WARNING:
            cell.fill = WRN_FILL


def build_workbook(result, out_path):
    wb = Workbook()
    findings = sorted(result["findings"], key=lambda x: (x[0], x[3] or 0))
    n_err = sum(1 for f in findings if f[6] == ERROR)
    n_wrn = sum(1 for f in findings if f[6] == WARNING)
    details = result["details"]

    # -- Summary ------------------------------------------------------------
    ws = wb.active
    ws.title = "Summary"
    h, t = result["header"], result["trailer"]
    rows = [
        ("Result", "PASS" if n_err == 0 else "FAIL"),
        ("Input file", os.path.abspath(result["path"])),
        ("Processed at", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Total lines", len(result["lines"])),
        ("Header agency", h["values"]["Agency_Number"][0] if h else "(missing)"),
        ("Header report period", h["values"]["Report_Period"][0] if h else "(missing)"),
        ("Detail records", len(details)),
        ("Trailer record count", t["values"]["Total_Records"][0] if t else "(missing)"),
        ("Unique SSNs", len({d["values"]["Social_Security_Number"][0] for d in details})),
        ("Errors", n_err),
        ("Warnings", n_wrn),
        ("Records with errors", sum(1 for d in details if d["worst"] == ERROR)),
        ("", ""),
        ("Control totals", ""),
    ]
    for f in DETAIL:
        if f["kind"] == "amount":
            total = sum(d["values"][f["name"]][0] for d in details
                        if isinstance(d["values"][f["name"]][0], float))
            rows.append((f["name"], round(total, 2)))
    for r, (k, v) in enumerate(rows, start=1):
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=v)
        if isinstance(v, dt.datetime):
            c.number_format = "mm/yyyy"
        if isinstance(v, float):
            c.number_format = "#,##0.00"
    res = ws["B1"]
    res.font = Font(bold=True, size=14)
    res.fill = OK_FILL if n_err == 0 else ERR_FILL
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 70

    # -- Detail ---------------------------------------------------------------
    ws = wb.create_sheet("Detail")
    cols = ["Line", "Status"] + [f["name"] for f in DETAIL]
    ws.append(cols)
    style_header_row(ws, len(cols))
    for f_i, f in enumerate(DETAIL, start=3):
        ws.cell(row=1, column=f_i).comment = Comment(
            "Pos %d-%d, len %d, %s" % (f["start"], f["end"], f["len"], f["kind"]), "INT5016")
    for r, d in enumerate(details, start=2):
        ws.cell(row=r, column=1, value=d["line_no"])
        st = ws.cell(row=r, column=2, value=d["worst"] or "OK")
        st.fill = {ERROR: ERR_FILL, WARNING: WRN_FILL}.get(d["worst"], OK_FILL)
        write_record_cells(ws, r, 3, d["values"], DETAIL)
    if details:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), len(details) + 1)
    ws.freeze_panes = "F2"   # keep Line/Status/Agency/SSN visible
    autosize(ws, 40)

    # -- Validation -----------------------------------------------------------
    ws = wb.create_sheet("Validation")
    cols = ["Line", "Record Type", "Field", "Start", "End", "Raw Value", "Severity", "Message"]
    ws.append(cols)
    style_header_row(ws, len(cols))
    if not findings:
        ws.append(["", "", "", "", "", "", "OK", "No issues found"])
    for f in findings:
        ws.append([f[0] or None, f[1], f[2], f[3], f[4], f[5], f[6], f[7]])
        ws.cell(row=ws.max_row, column=6).font = MONO
        ws.cell(row=ws.max_row, column=7).fill = ERR_FILL if f[6] == ERROR else WRN_FILL
    ws.auto_filter.ref = "A1:H%d" % ws.max_row
    autosize(ws, 90)

    # -- Header / Trailer -----------------------------------------------------
    ws = wb.create_sheet("Header_Trailer")
    ws.append(["Record", "Line"] + [f["name"] for f in TRAILER])
    style_header_row(ws, len(TRAILER) + 2)
    for label, rec, layout in (("Header", h, HEADER), ("Trailer", t, TRAILER)):
        if not rec:
            ws.append([label, "(missing)"])
            continue
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=rec["line_no"])
        write_record_cells(ws, r, 3, rec["values"], layout)
    autosize(ws)

    # -- Layout ---------------------------------------------------------------
    ws = wb.create_sheet("Layout")
    ws.append(["Record", "#", "Field", "Start", "End", "Length", "Type", "Required"])
    style_header_row(ws, 8)
    for label, layout in (("Header", HEADER), ("Detail", DETAIL), ("Trailer", TRAILER)):
        for i, f in enumerate(layout, start=1):
            ws.append([label, i, f["name"], f["start"], f["end"], f["len"], f["kind"],
                       "Y" if f["required"] else ""])
    autosize(ws)

    # -- Raw ------------------------------------------------------------------
    ws = wb.create_sheet("Raw")
    ws.append(["Line", "Length", "Record"])
    style_header_row(ws, 3)
    for i, line in enumerate(result["lines"], start=1):
        ws.append([i, len(line), line])
        ws.cell(row=i + 1, column=3).font = MONO
    ws.column_dimensions["C"].width = 120

    wb.save(out_path)
    return n_err, n_wrn


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
# Files in a folder that are never INT5016 data files
SKIP_EXT = {".xlsx", ".xls", ".xlsm", ".py", ".bat", ".cmd", ".xsl", ".xslt", ".xml",
            ".zip", ".pdf", ".doc", ".docx", ".log", ".ini", ".lnk"}


def collect_inputs(inputs, pattern):
    files = []
    for item in inputs:
        item = item.strip().strip('"')
        if os.path.isdir(item):
            found = sorted(p for p in glob.glob(os.path.join(item, pattern))
                           if os.path.isfile(p) and not os.path.basename(p).startswith(("~$", "."))
                           and os.path.splitext(p)[1].lower() not in SKIP_EXT)
            if not found:
                others = [n for n in os.listdir(item) if os.path.isfile(os.path.join(item, n))]
                sys.stderr.write("WARNING: no data files matching '%s' in %s\n" % (pattern, item))
                if others:
                    sys.stderr.write("         files present (skipped): %s\n" % ", ".join(others[:20]))
                else:
                    sys.stderr.write("         the folder is empty\n")
            files.extend(found)
        elif os.path.isfile(item):
            files.append(item)
        else:
            sys.stderr.write("WARNING: not found: %s\n" % item)
    return files


def main(argv=None):
    ap = argparse.ArgumentParser(description="Validate an INT5016 fixed-width file and convert it to Excel.")
    ap.add_argument("inputs", nargs="+", help="Input file(s) or folder(s)")
    ap.add_argument("-o", "--output-dir", help="Folder for .xlsx output (default: next to each input)")
    ap.add_argument("--pattern", default="*",
                    help="File pattern for folder inputs (default: every file except "
                         "Excel/script/XML etc., so files with no extension or .dat are included)")
    args = ap.parse_args(argv)

    files = collect_inputs(args.inputs, args.pattern)
    if not files:
        sys.stderr.write("ERROR: no input files found\n")
        return 1
    if args.output_dir:
        os.makedirs(args.output_dir, exist_ok=True)

    worst_rc = 0
    for path in files:
        base = os.path.splitext(os.path.basename(path))[0]
        out_dir = args.output_dir or os.path.dirname(os.path.abspath(path))
        out_path = os.path.join(out_dir, base + "_validated.xlsx")
        print("Processing %s" % path)
        try:
            result = process_file(path)
            n_err, n_wrn = build_workbook(result, out_path)
        except PermissionError:
            sys.stderr.write("  ERROR: cannot write %s - is it open in Excel?\n" % out_path)
            worst_rc = max(worst_rc, 1)
            continue
        except Exception as exc:  # keep going with the other files
            sys.stderr.write("  ERROR: failed to process file: %s\n" % exc)
            worst_rc = max(worst_rc, 1)
            continue

        status = "PASS" if n_err == 0 else "FAIL"
        print("  %s  detail records=%d  errors=%d  warnings=%d"
              % (status, len(result["details"]), n_err, n_wrn))
        for f in [f for f in result["findings"] if f[6] == ERROR][:10]:
            print("    line %s %s: %s" % (f[0] or "-", f[2] or f[1], f[7]))
        if n_err > 10:
            print("    ... %d more errors, see Validation sheet" % (n_err - 10))
        print("  -> %s" % out_path)
        if n_err:
            worst_rc = 2 if worst_rc != 1 else 1
    return worst_rc


if __name__ == "__main__":
    sys.exit(main())
