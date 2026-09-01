# plot_il_sweep.py — optical wavelength sweep plotter for the web dashboard.
# Paste into the code editor and run. Pyodide-safe: matplotlib + stdlib only
# (xlsx is read with zipfile + xml.etree, NOT openpyxl/pandas -- neither is
# assumed to be installable in this sandbox).
#
# Accepts .xlsx, .csv, or .txt for this file. For .xlsx: the workbook may have
# several sheets; the sheet actually plotted is whichever one is named "IL" or
# "RXTXAvgIL" (case-insensitive) -- see SHEET_NAME_CANDIDATES. Within that
# sheet (or within a .csv/.txt file), the X column is "Wavelength_nm" or
# "wavelength", and the Y column is "RXTXAvgIL" or "IL" -- see
# X_COLUMN_CANDIDATES / Y_COLUMN_CANDIDATES. Matching is case/space/underscore
# -insensitive. Which parser runs is chosen from the file's own extension
# (.xlsx / .csv / .txt) -- the filename itself is never restricted or
# rewritten, it's only inspected for its extension.
#
# Axis titles are fixed ("Wavelength [nm]" / "Loss [dB]") regardless of the
# source column names, so output is visually consistent across differently
# -named source files.
#
# Default Y axis is transmission dB with 0 dB at the top. Insertion-loss
# columns (positive = lossy) are negated automatically, so either sign
# convention plots the right way up. Change Y_MODE below to switch convention.
#
# This script emits a finished PNG, so its orientation is fixed once drawn --
# nothing downstream re-flips it. The dashboard's separate Spectrum window
# draws its own plot from the parsed numbers and is not affected by this file.
#
# If this file's raw content can't be located on the dashboard's `test`
# object (field names vary by dashboard build), this falls back to the
# dashboard's own generic test["parsed"]["sweep"/"traces"] structure, so
# whatever worked before continues to work even if the new direct-parsing
# path can't find anything.
#
# Dashboard globals: test (this file), tests (all files in the folder)

import io
import base64
import re
import zipfile
from xml.etree import ElementTree as ET

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ----------------------------------------------------------------- config ---
Y_MODE     = "transmission"  # "transmission" -> negative dB, 0 at top (default)
                             # "loss"         -> positive dB, axis inverted, 0 at top
                             # "raw"          -> values as supplied, axis untouched
ANNOTATE   = True            # peak / min callouts on the first trace
XLIM       = None            # zoom, e.g. (1540, 1560)
YLIM       = None            # fixed scale, e.g. (-40, 0); None = auto
FIGSIZE    = (9, 4.6)
DEBUG_TINT = False           # True = pink background, to confirm which panel
                             # you are looking at; set back to False afterwards

# Which xlsx sheet holds the data -- first sheet whose name matches ANY of
# these (case/space/underscore-insensitive) is used. Extend if a new sheet
# naming convention shows up.
SHEET_NAME_CANDIDATES = ["IL", "RXTXAvgIL"]

# Column headers -- same matching rules as above. Extend similarly.
X_COLUMN_CANDIDATES = ["Wavelength_nm", "wavelength"]
Y_COLUMN_CANDIDATES = ["RXTXAvgIL", "IL"]

# Fixed axis titles -- always these, regardless of the source column names.
X_AXIS_LABEL = "Wavelength [nm]"
Y_AXIS_LABEL = "Loss [dB]"
# -----------------------------------------------------------------------------

NAN = float("nan")


def emit(fig):
    """Hand the figure to the dashboard as a base64 PNG."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, bbox_inches="tight")
    print("DASH_PLOT_PNG:" + base64.b64encode(buf.getvalue()).decode())
    plt.close(fig)


def meta_get(meta, *candidates, default=None):
    """Case/space-insensitive metadata lookup with prefix fallback."""
    if not meta:
        return default
    norm = {str(k).strip().lower().replace(" ", ""): v for k, v in meta.items()}
    for cand in candidates:
        key = cand.strip().lower().replace(" ", "")
        if key in norm:
            return norm[key]
        for k, v in norm.items():
            if k.startswith(key):
                return v
    return default


def to_float(v):
    """None / non-numeric -> nan, so matplotlib gaps the line instead of raising."""
    try:
        return NAN if v is None else float(v)
    except (TypeError, ValueError):
        return NAN


def extremes(xs, ys):
    """Single pass -> ((max_y, max_x), (min_y, min_x)), or None if all nan."""
    hi = lo = None
    for x, y in zip(xs, ys):
        if y != y:
            continue
        if hi is None or y > hi[0]:
            hi = (y, x)
        if lo is None or y < lo[0]:
            lo = (y, x)
    return None if hi is None else (hi, lo)


def _norm_name(s):
    """Case/space/underscore-insensitive normalisation for sheet/column matching."""
    return str(s).strip().lower().replace(" ", "").replace("_", "")


def _name_matches(candidates, name):
    if name is None:
        return False
    n = _norm_name(name)
    return any(n == _norm_name(c) for c in candidates)


def get_filename(test):
    return test.get("name") or test.get("filename") or ""


def detect_extension(filename):
    """-> '.xlsx' / '.csv' / '.txt' / None. Only the extension is inspected --
    the filename itself is never validated, restricted, or rewritten."""
    lower = filename.lower()
    for ext in (".xlsx", ".csv", ".txt"):
        if lower.endswith(ext):
            return ext
    return None


def collect_candidates(obj, out=None):
    """Recursively collect every string/bytes value anywhere in obj that could
    plausibly be file content (skips short strings such as names/units).
    Field names for raw content vary across dashboard builds, so this doesn't
    guess a specific key -- it searches everything and lets the parsers below
    decide what's actually usable."""
    if out is None:
        out = []
    if isinstance(obj, (bytes, bytearray)):
        if len(obj) > 20:
            out.append(bytes(obj))
    elif isinstance(obj, str):
        if len(obj) > 20:
            out.append(obj)
    elif isinstance(obj, dict):
        for v in obj.values():
            collect_candidates(v, out)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            collect_candidates(v, out)
    return out


# ── XLSX parsing: pure stdlib (zipfile + xml.etree), no openpyxl/pandas ──────
# An .xlsx is a ZIP archive of XML parts. The pieces needed here:
#   xl/workbook.xml            -- sheet names -> relationship IDs
#   xl/_rels/workbook.xml.rels -- relationship IDs -> the actual sheetN.xml path
#   xl/sharedStrings.xml       -- text cells are stored as indices into this
#   xl/worksheets/sheetN.xml   -- the actual cell grid for one sheet

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_ns = {"m": _NS_MAIN}


def _col_letters_to_index(cell_ref):
    """'B3' -> 1 (zero-based column index)."""
    letters = re.match(r"[A-Za-z]+", cell_ref).group().upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _parse_shared_strings(zf):
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    out = []
    for si in root.findall("m:si", _ns):
        texts = si.findall(".//m:t", _ns)
        out.append("".join(t.text or "" for t in texts))
    return out


def _parse_workbook_sheets(zf):
    """-> {sheet_name: sheet_path_in_zip}, in workbook-declared order."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_target = {}
    for rel in rels_root:
        rid_to_target[rel.get("Id")] = rel.get("Target")

    sheets = {}
    for sheet_el in wb_root.findall("m:sheets/m:sheet", _ns):
        name = sheet_el.get("name")
        rid = sheet_el.get(f"{{{_NS_REL}}}id")
        target = rid_to_target.get(rid)
        if not target:
            continue
        # Target is either absolute ("/xl/worksheets/sheet1.xml") or relative
        # to the xl/ folder ("worksheets/sheet1.xml") -- per the OOXML spec,
        # both forms are valid and different writers use different ones.
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = f"xl/{target}"
        sheets[name] = path
    return sheets


def _parse_sheet_rows(zf, sheet_path, shared_strings):
    """-> list of rows, each a {col_index: value} dict. Empty cells are simply
    absent from the dict (xlsx omits them from the XML entirely).

    Handles the three cell-type variants actually seen in real xlsx output:
      t="s"         -> shared string, value is an index into sharedStrings.xml, in <v>
      t="inlineStr" -> string stored directly in the cell as <is><t>text</t></is>
                       (no <v> at all -- this is what openpyxl writes by default)
      t="str" / absent (numeric) -> value is plain text inside <v>
    """
    root = ET.fromstring(zf.read(sheet_path))
    rows_out = []
    for row_el in root.findall(".//m:sheetData/m:row", _ns):
        row_vals = {}
        for cell_el in row_el.findall("m:c", _ns):
            ref = cell_el.get("r")
            if not ref:
                continue
            col_idx = _col_letters_to_index(ref)
            ctype = cell_el.get("t")

            if ctype == "inlineStr":
                is_el = cell_el.find("m:is", _ns)
                t_el = is_el.find("m:t", _ns) if is_el is not None else None
                val = t_el.text if t_el is not None else None
            else:
                v_el = cell_el.find("m:v", _ns)
                if v_el is None or v_el.text is None:
                    val = None
                elif ctype == "s":              # shared string -> index lookup
                    val = shared_strings[int(v_el.text)]
                else:                            # "str" or numeric (absent/"n")
                    val = v_el.text

            row_vals[col_idx] = val
        rows_out.append(row_vals)
    return rows_out


def parse_xlsx(raw_bytes):
    """-> (sheet_name, x_header, y_header, xs, ys) from the first sheet whose
    name matches SHEET_NAME_CANDIDATES, using the first header row that has
    both an X and a Y column match. None if nothing matches."""
    with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
        shared = _parse_shared_strings(zf)
        sheets = _parse_workbook_sheets(zf)

        for name, path in sheets.items():
            if not _name_matches(SHEET_NAME_CANDIDATES, name):
                continue

            rows = _parse_sheet_rows(zf, path, shared)
            if not rows:
                continue

            header = rows[0]   # assumes header is the first row in the sheet
            x_col = y_col = None
            x_header_text = y_header_text = None
            for idx, text in header.items():
                if x_col is None and _name_matches(X_COLUMN_CANDIDATES, text):
                    x_col, x_header_text = idx, text
                if y_col is None and _name_matches(Y_COLUMN_CANDIDATES, text):
                    y_col, y_header_text = idx, text
            if x_col is None or y_col is None:
                continue   # this sheet matched by name but not by columns -- keep looking

            xs = [to_float(r.get(x_col)) for r in rows[1:]]
            ys = [to_float(r.get(y_col)) for r in rows[1:]]
            return name, x_header_text, y_header_text, xs, ys

    return None


# ── CSV / TXT parsing (stdlib csv module, delimiter auto-detected) ──────────

def parse_delimited_text(text):
    """-> (x_header, y_header, xs, ys) from CSV/TXT. Delimiter is guessed from
    the header line (comma, then tab, else whitespace-split). None if no
    header row matches both X and Y column candidates."""
    import csv as csv_mod

    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        return None

    first = lines[0]
    if "," in first:
        rows = list(csv_mod.reader(lines, delimiter=","))
    elif "\t" in first:
        rows = list(csv_mod.reader(lines, delimiter="\t"))
    else:
        rows = [ln.split() for ln in lines]

    header = rows[0]
    x_col = y_col = None
    x_header_text = y_header_text = None
    for idx, text_h in enumerate(header):
        if x_col is None and _name_matches(X_COLUMN_CANDIDATES, text_h):
            x_col, x_header_text = idx, text_h
        if y_col is None and _name_matches(Y_COLUMN_CANDIDATES, text_h):
            y_col, y_header_text = idx, text_h
    if x_col is None or y_col is None:
        return None

    xs, ys = [], []
    for row in rows[1:]:
        if len(row) <= max(x_col, y_col):
            continue
        xs.append(to_float(row[x_col]))
        ys.append(to_float(row[y_col]))
    return x_header_text, y_header_text, xs, ys


# ── Locating raw content on the dashboard's test object ─────────────────────

def find_direct_parse(test):
    """Try to parse this file's own raw content directly (xlsx/csv/txt),
    dispatching on the file's extension. Returns (xname, yname, xs, ys) or
    None if nothing usable was found."""
    filename = get_filename(test)
    ext = detect_extension(filename)
    candidates = collect_candidates(test)

    if ext in (".xlsx", None):
        for c in candidates:
            raw = None
            if isinstance(c, (bytes, bytearray)) and bytes(c[:2]) == b"PK":
                raw = bytes(c)
            elif isinstance(c, str):
                # binary content sometimes crosses the JS/Python bridge as base64
                try:
                    decoded = base64.b64decode(c, validate=True)
                    if decoded[:2] == b"PK":
                        raw = decoded
                except Exception:
                    pass
            if raw is not None:
                result = parse_xlsx(raw)
                if result:
                    _, xh, yh, xs, ys = result
                    if any(y == y for y in ys):
                        return xh, yh, xs, ys

    if ext in (".csv", ".txt", None):
        for c in candidates:
            if isinstance(c, str):
                result = parse_delimited_text(c)
                if result:
                    xh, yh, xs, ys = result
                    if any(y == y for y in ys):
                        return xh, yh, xs, ys

    return None


def find_generic_fallback(test):
    """Fall back to the dashboard's own generic parser
    (test["parsed"]["sweep"/"traces"]), for whatever this already handled
    correctly before the direct-parsing path above existed."""
    parsed = test.get("parsed")
    if not parsed or not parsed.get("traces") or not parsed.get("sweep"):
        return None
    xs = [to_float(v) for v in parsed["sweep"]["values"]]
    traces = parsed["traces"]
    if not traces:
        return None
    first = traces[0]
    ys = [to_float(v) for v in first["values"]]
    if not any(y == y for y in ys):
        return None
    return parsed["sweep"].get("name"), first.get("name"), xs, ys


def read_test(test):
    """-> (sweep, series=[(name, values)], meta), sorted by X ascending.
    Tries direct parsing (xlsx/csv/txt) first; falls back to the dashboard's
    own generic parser if direct parsing finds nothing."""
    result = find_direct_parse(test)
    if result is None:
        result = find_generic_fallback(test)
    if result is None:
        return None

    _xname, yname, sweep, values = result

    keep = [i for i, x in enumerate(sweep) if x == x]
    keep.sort(key=lambda i: sweep[i])
    sweep = [sweep[i] for i in keep]
    values = [values[i] if i < len(values) else NAN for i in keep]

    if not sweep or not any(y == y for y in values):
        return None

    series = [(yname or "Loss", values)]
    return sweep, series, (test.get("metadata") or {})


def build_title(name, meta):
    bits = []
    device = meta_get(meta, "device_id", "device")
    if device:
        bits.append(str(device))
    bits.append(str(name))
    tls = meta_get(meta, "TLSPower (dBm)", "TLSPower")
    if tls is not None:
        bits.append(f"TLS {tls} dBm")
    scans = meta_get(meta, "NumberOfScans")
    if scans is not None:
        bits.append(f"{scans} scans")
    date = meta_get(meta, "date")
    if date:
        bits.append(str(date))
    return "  \u00b7  ".join(bits)


def orient(vals, is_loss):
    """Convert source values to whatever Y_MODE asks for."""
    if Y_MODE == "raw":
        return vals
    want_loss = (Y_MODE == "loss")
    if is_loss == want_loss:
        return vals
    return [-y if y == y else y for y in vals]


def plot(test):
    data = read_test(test)
    name = test.get("name", "?")
    if not data:
        print(f"No plottable data on '{name}'.")
        print(f"  Detected extension: {detect_extension(get_filename(test))!r}")
        print(f"  Sheet name candidates tried:  {SHEET_NAME_CANDIDATES}")
        print(f"  X column candidates tried:    {X_COLUMN_CANDIDATES}")
        print(f"  Y column candidates tried:    {Y_COLUMN_CANDIDATES}")
        print("  Top-level keys on this file:", sorted(test.keys()))
        return

    sweep, series, meta = data

    finite = [y for _, vals in series for y in vals if y == y]
    is_loss = (sum(finite) / len(finite)) > 0

    fig, ax = plt.subplots(figsize=FIGSIZE)
    if DEBUG_TINT:
        ax.set_facecolor("#ffe9e9")
    drawn = []

    for idx, (tname, vals) in enumerate(series):
        ys = orient(vals, is_loss)
        ax.plot(sweep, ys, lw=0.9, label=tname)
        drawn.append((tname, ys))

        if ANNOTATE and idx == 0:
            ends = extremes(sweep, ys)
            if ends:
                inverted = (Y_MODE == "loss")
                top, bottom = (ends[1], ends[0]) if inverted else (ends[0], ends[1])
                for (val, at), tag, dy in ((top, "peak", -30),
                                           (bottom, "min", 22)):
                    ax.annotate(
                        f"{tag} {val:.2f} dB\n@ {at:.3f} nm",
                        xy=(at, val), xytext=(0, dy),
                        textcoords="offset points", ha="center", fontsize=8,
                        arrowprops=dict(arrowstyle="-", lw=0.6, alpha=0.6),
                    )

    # Axis titles are always fixed, regardless of the source column names.
    ax.set_xlabel(X_AXIS_LABEL)
    ax.set_ylabel(Y_AXIS_LABEL)
    ax.set_title(build_title(name, meta), fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.margins(x=0, y=0.10)
    if XLIM:
        ax.set_xlim(*XLIM)

    if YLIM:
        ax.set_ylim(*YLIM)
    elif Y_MODE == "transmission":
        # 0 dB is the physical maximum for transmission (ratio <= 1) -- always the
        # top of the axis. The bottom auto-scales per file, since how deep a given
        # device's loss goes genuinely differs from file to file.
        bottom, top = ax.get_ylim()
        ax.set_ylim(bottom=min(bottom, top), top=0)
    elif Y_MODE == "loss":
        # Same physical constraint, inverted convention: loss >= 0 dB, so 0 is the
        # floor here and sits at the top of the (intentionally inverted) axis.
        bottom, top = ax.get_ylim()
        ax.set_ylim(bottom=max(bottom, top), top=0)

    if len(drawn) > 1:
        ax.legend(fontsize=9, loc="best")

    fig.tight_layout()
    emit(fig)
    report(name, sweep, drawn, meta, is_loss)


def report(name, sweep, drawn, meta, is_loss):
    step = (sweep[-1] - sweep[0]) / (len(sweep) - 1) * 1000 if len(sweep) > 1 else 0
    print(f"Test: {name}")
    print(f"  source     : {meta_get(meta, 'source_file', default='?')}")
    print(f"  date       : {meta_get(meta, 'date', default='?')}")
    print(f"  TLS        : {meta_get(meta, 'TLSPower (dBm)', default='?')} dBm, "
          f"{meta_get(meta, 'TLSOutput', default='?')}, "
          f"{meta_get(meta, 'NumberOfScans', default='?')} scan(s)")
    print(f"  sweep      : {sweep[0]:.3f} - {sweep[-1]:.3f} nm  "
          f"({len(sweep)} pts, {step:.1f} pm step)")
    print(f"  convention : source was "
          f"{'insertion loss' if is_loss else 'transmission'} "
          f"-> plotted as {Y_MODE}")
    for tname, ys in drawn:
        ends = extremes(sweep, ys)
        if not ends:
            continue
        (hi_v, hi_x), (lo_v, lo_x) = ends
        gaps = sum(1 for y in ys if y != y)
        note = f"   [{gaps} gap(s)]" if gaps else ""
        print(f"  {str(tname):<12} max {hi_v:8.3f} dB @ {hi_x:.3f} nm   "
              f"min {lo_v:8.3f} dB @ {lo_x:.3f} nm   "
              f"p-p {hi_v - lo_v:.3f} dB{note}")


plot(test)