# S-parameter (S11/S21/S12/S22) plot for the dashboard.
#
# The dashboard's own generic parser only understands simple sweep+trace
# CSV-style files and fails on Touchstone (.s2p) format — that's the
# "Could not extract a plottable sweep + trace from this file" message.
# This script bypasses that and parses the raw Touchstone text itself.
#
# Globals expected from the dashboard:
#   tests : list of every test on this BB (dict with metadata + raw content)
#   data  : dict, may or may not be populated — not relied on here
#
# ── Edit these to change what's plotted ─────────────────────────────────────
PARAMS            = ['s11', 's21', 's12', 's22']   # which S-parameters to show
LAYOUT            = 'grid'      # 'grid' (one subplot per param) or 'overlay' (all on one axes)
FREQ_UNIT_DIVISOR = 1e9          # divides the raw frequency array for the x-axis (Hz -> GHz)
TARGET_NAME       = None         # e.g. 'CH1_P2_circut_P1_intrinsic_btm_0dBm' to pick a
                                 # specific test; None = first parseable .s2p found

import io, base64
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


def emit(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, bbox_inches="tight")
    print("DASH_PLOT_PNG:" + base64.b64encode(buf.getvalue()).decode())
    plt.close(fig)


# ── Raw Touchstone (.s2p) parser ─────────────────────────────────────────────
# Format (2-port, dB-angle, as exported by the Keysight PNA):
#   ! ...comment lines, ignored...
#   # Hz S dB R 50          <- header: freq unit, format, impedance
#   freq S11dB S11ang S21dB S21ang S12dB S12ang S22dB S22ang   <- one row per point
#
# Handles: comment lines, the header line, blank lines, and any trailing
# non-numeric junk (e.g. text accidentally appended after the data) by
# simply skipping any line that doesn't parse as exactly 9 numbers.

_FREQ_UNIT_SCALE = {'hz': 1.0, 'khz': 1e3, 'mhz': 1e6, 'ghz': 1e9}


def parse_touchstone(raw_text: str):
    """
    Parse raw Touchstone (.s2p) text into frequency (Hz) + S-param dB arrays.
    Returns a dict: {'freq': np.array, 's11': np.array(dB), 's21': ..., 's12': ..., 's22': ...}
    or None if no data rows were found at all.
    """
    freq_scale = 1.0  # default Hz, overridden if header says otherwise
    rows = []
    n_skipped = 0

    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith('!'):
            continue   # comment line
        if line.startswith('#'):
            # header line, e.g. "# Hz S dB R 50" — extract the frequency unit
            parts = line[1:].strip().lower().split()
            if parts:
                freq_scale = _FREQ_UNIT_SCALE.get(parts[0], 1.0)
            continue
        # attempt to parse as 9 whitespace-separated floats: freq + 4×(dB, angle)
        tokens = line.split()
        if len(tokens) != 9:
            n_skipped += 1
            continue
        try:
            vals = [float(t) for t in tokens]
        except ValueError:
            n_skipped += 1
            continue   # non-numeric junk (e.g. trailing text) — skip, don't crash
        rows.append(vals)

    if not rows:
        return None

    arr = np.array(rows)   # shape (n_points, 9)
    freq_hz = arr[:, 0] * freq_scale
    return {
        'freq': freq_hz,
        's11': arr[:, 1],   # dB magnitude columns (angle columns 2,4,6,8 are ignored for mag plots)
        's21': arr[:, 3],
        's12': arr[:, 5],
        's22': arr[:, 7],
        '_n_skipped_lines': n_skipped,
    }


# ── Locate raw Touchstone text on the test object ───────────────────────────
# The exact field name holding raw file content isn't confirmed, so this
# tries several common ones. If none match, it reports what keys ARE
# present so the real field name can be added here.
_RAW_TEXT_KEYS = ('raw', 'raw_text', 'content', 'text', 'body', 'file_content', 'contents', 'data')


def find_raw_text(test_obj):
    if isinstance(test_obj, str):
        return test_obj
    if isinstance(test_obj, dict):
        for key in _RAW_TEXT_KEYS:
            val = test_obj.get(key)
            if isinstance(val, str) and len(val) > 20:
                return val
    return None


# ── Main ──────────────────────────────────────────────────────────────────
_tests = globals().get('tests', [])

if not _tests:
    print("No 'tests' available on this BB yet.")
    print("Drop an .s2p file into the Test results tab and re-run.")
else:
    _candidate = None
    _candidate_name = None
    _candidate_raw = None

    for t in _tests:
        name = None
        if isinstance(t, dict):
            name = t.get('name') or t.get('filename')
        if TARGET_NAME and name != TARGET_NAME:
            continue
        raw = find_raw_text(t)
        if raw is None:
            continue
        parsed = parse_touchstone(raw)
        if parsed is not None:
            _candidate = t
            _candidate_name = name or '(unnamed)'
            _candidate_raw = parsed
            break

    if _candidate_raw is None:
        print("No parseable Touchstone (.s2p) content found among tests.")
        if _tests:
            _sample = _tests[0]
            if isinstance(_sample, dict):
                print(f"First test's available keys: {list(_sample.keys())}")
            else:
                print(f"First test's type: {type(_sample)}")
    else:
        freqs_scaled = _candidate_raw['freq'] / FREQ_UNIT_DIVISOR
        x_label = 'Frequency (GHz)' if FREQ_UNIT_DIVISOR == 1e9 else 'Frequency'
        n_pts = len(_candidate_raw['freq'])
        n_skip = _candidate_raw['_n_skipped_lines']

        print(f"Parsed '{_candidate_name}': {n_pts} data point(s)"
              + (f", {n_skip} non-data line(s) skipped" if n_skip else ""))

        available = {p: _candidate_raw[p] for p in PARAMS if p in _candidate_raw}

        if LAYOUT == 'grid':
            ncols = 2
            nrows = int(np.ceil(len(available) / ncols))
            fig, axes = plt.subplots(nrows, ncols, figsize=(5 * ncols, 3.5 * nrows), squeeze=False)
            axes_flat = axes.flatten()
            for ax, (p, mag_db) in zip(axes_flat, available.items()):
                ax.plot(freqs_scaled, mag_db, lw=1.2)
                ax.set_title(p.upper())
                ax.set_xlabel(x_label)
                ax.set_ylabel(f'{p.upper()} (dB)')
                ax.grid(True, alpha=0.3)
            for ax in axes_flat[len(available):]:
                ax.set_visible(False)
            fig.suptitle(_candidate_name, y=1.02)
        else:  # overlay
            fig, ax = plt.subplots(figsize=(9, 5))
            for p, mag_db in available.items():
                ax.plot(freqs_scaled, mag_db, lw=1.2, label=p.upper())
            ax.set_xlabel(x_label)
            ax.set_ylabel('Magnitude (dB)')
            ax.set_title(_candidate_name)
            ax.grid(True, alpha=0.3)
            ax.legend(fontsize=9, loc='best')

        fig.tight_layout()
        emit(fig)
        print(f"Plotted {list(available)} from '{_candidate_name}'.")