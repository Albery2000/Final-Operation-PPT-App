import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
import base64
from openpyxl import load_workbook

# =============================================================================
# REPORT SHEET PARSER (supports the new "Detailed Production Report" layout)
# =============================================================================

def _read_bytes(file_content):
    """Return raw bytes from an uploaded file / path / bytes object."""
    if isinstance(file_content, (bytes, bytearray)):
        return bytes(file_content)
    if isinstance(file_content, str):
        with open(file_content, 'rb') as f:
            return f.read()
    if hasattr(file_content, 'getvalue'):
        return file_content.getvalue()
    file_content.seek(0)
    return file_content.read()


def _norm(value):
    """Normalise a header cell: lower-case, no line breaks / dots / extra spaces."""
    if value is None:
        return ''
    text = str(value).replace('\n', ' ').replace('\r', ' ').lower()
    for ch in ['.', ':', '*', '(', ')']:
        text = text.replace(ch, ' ')
    return ' '.join(text.split())


def _clean_text(value):
    if value is None:
        return ''
    return ' '.join(str(value).replace('\n', ' ').replace('\r', ' ').split())


def load_report_grid(file_content, max_rows=400, max_cols=30):
    """
    Read only the top part of the 'Report' sheet as a plain grid (list of rows).
    Using read_only mode keeps this fast even though the workbook contains
    huge helper sheets (e.g. 'Test data' with ~100k rows).
    """
    wb = load_workbook(io.BytesIO(_read_bytes(file_content)), data_only=True, read_only=True)
    try:
        sheet_name = None
        for name in wb.sheetnames:
            if _norm(name) == 'report':
                sheet_name = name
                break
        if sheet_name is None:
            raise ValueError("Could not find a sheet named 'Report' in the workbook. "
                             f"Sheets found: {', '.join(wb.sheetnames[:15])}")
        ws = wb[sheet_name]
        return [list(r) for r in ws.iter_rows(min_row=1, max_row=max_rows,
                                              max_col=max_cols, values_only=True)]
    finally:
        wb.close()


def _find_header_row(grid):
    """Row index (0-based) of the row that holds the 'Well Name' sub-header."""
    for i, row in enumerate(grid[:60]):
        if any(_norm(c) in ('well name', 'well') for c in row):
            return i
    raise ValueError("Could not locate the header row (a cell called 'Well Name').")


# Each entry: internal key -> (output column name, matcher(group_header, sub_header))
_COLUMN_RULES = [
    ('field',      'Field',                    lambda g, s: g == 'field'),
    ('well',       'Well Name',                lambda g, s: s in ('well name', 'well')),
    ('formation',  'Formation',                lambda g, s: s == 'formation'),
    ('lifting',    'Lifting Method / Status',  lambda g, s: 'lifting' in s or s == 'status'),
    ('gross_diff', 'Gross Diff STB',           lambda g, s: 'gross' in s and 'diff' in s),
    ('gross',      'Gross STB',                lambda g, s: 'gross' in s and 'diff' not in s),
    ('net_diff',   'Net Diff BO',              lambda g, s: 'net' in s and 'diff' in s),
    ('net_bo',     'Net BO',                   lambda g, s: 'net' in s and 'bo' in s and 'diff' not in s),
    ('hours',      'Flowing Hours',            lambda g, s: 'flowing' in s or 'hours' in s),
    ('wc',         'W/C %',                    lambda g, s: g == 'w/c' or (g.startswith('w/c') and s in ('%', ''))),
    ('water',      'Water Prod. BW',           lambda g, s: 'water prod' in g or s == 'bw'),
    ('gas',        'Gas MMSCFD',               lambda g, s: g == 'gas' or 'mmscfd' in s),
    ('gor',        'GOR SCF/BBL',              lambda g, s: g == 'gor' or 'scf' in s),
    ('cum_oil',    'Cum. Oil (M BBLS)',        lambda g, s: 'cumulative' in g and 'oil' in s),
    ('cum_water',  'Cum. Water (M BBLS)',      lambda g, s: 'cumulative' in g and 'water' in s),
]

# Only these are required for the analysis; the rest are shown when present.
_REQUIRED_KEYS = ['field', 'well', 'net_bo', 'net_diff']


def detect_report_columns(grid):
    """
    Detect the header row and map every column of the report to a clean name.
    Group headers (row above, merged cells) are forward-filled so that
    'TOTAL PRODUCTION' applies to all its sub-columns.
    Returns (header_row_index, {key: (col_index, clean_name)}).
    """
    hdr = _find_header_row(grid)
    sub_row = grid[hdr]
    group_row = grid[hdr - 1] if hdr > 0 else [None] * len(sub_row)

    groups, last = [], ''
    for c in group_row:
        if c is not None and str(c).strip() != '':
            last = _norm(c)
        groups.append(last)

    mapping = {}
    for idx, sub in enumerate(sub_row):
        g, s = groups[idx] if idx < len(groups) else '', _norm(sub)
        if not g and not s:
            continue
        for key, name, rule in _COLUMN_RULES:
            if key in mapping:
                continue
            try:
                if rule(g, s):
                    mapping[key] = (idx, name)
                    break
            except Exception:
                continue

    # Fallbacks for older layouts where the well-name / field headers were blank
    if 'field' not in mapping:
        mapping['field'] = (0, 'Field')
    if 'well' not in mapping:
        mapping['well'] = (mapping['field'][0] + 1, 'Well Name')

    return hdr, mapping


def extract_report_table(grid):
    """
    Build the well-level DataFrame from the grid.
    Returns (df, sheet_totals, report_date, column_map).
    """
    hdr, mapping = detect_report_columns(grid)

    missing = [k for k in _REQUIRED_KEYS if k not in mapping]
    if missing:
        labels = {'net_bo': "'Net BO'", 'net_diff': "'Net diff. BO'",
                  'field': "'Field'", 'well': "'Well Name'"}
        raise ValueError("Missing required column(s): " + ", ".join(labels[m] for m in missing))

    field_idx = mapping['field'][0]
    well_idx = mapping['well'][0]

    records, sheet_totals, cum_row = [], None, None
    current_field = ''
    for row in grid[hdr + 1:]:
        a = _clean_text(row[field_idx]) if field_idx < len(row) else ''
        upper_a = a.upper()

        if upper_a.startswith('CUM'):
            cum_row = row
            continue
        if upper_a.startswith('TOTAL'):
            sheet_totals = row
            break

        if a:
            current_field = a
        well = _clean_text(row[well_idx]) if well_idx < len(row) else ''
        if not well or well.lower() == current_field.lower():
            continue

        rec = {'Field': current_field}
        for key, (idx, name) in mapping.items():
            if key == 'field':
                continue
            val = row[idx] if idx < len(row) else None
            rec[name] = _clean_text(val) if key in ('well', 'formation', 'lifting') else val
        records.append(rec)

    df = pd.DataFrame(records)

    numeric_keys = ['gross', 'gross_diff', 'net_bo', 'net_diff', 'hours', 'wc',
                    'water', 'gas', 'gor', 'cum_oil', 'cum_water']
    for key in numeric_keys:
        if key in mapping and mapping[key][1] in df.columns:
            df[mapping[key][1]] = pd.to_numeric(df[mapping[key][1]], errors='coerce')

    # Report date (cell right of the 'Date' label in the header block)
    report_date = None
    for row in grid[:hdr]:
        for i, c in enumerate(row[:-1]):
            if _norm(c) == 'date' and row[i + 1] is not None:
                report_date = row[i + 1]
                break
        if report_date is not None:
            break

    totals = {}
    if sheet_totals is not None:
        for key, (idx, name) in mapping.items():
            if idx < len(sheet_totals) and key not in ('field', 'well', 'formation', 'lifting'):
                totals[name] = pd.to_numeric(sheet_totals[idx], errors='coerce')
        # Total number of wells is written in the 'Well Name' column of the TOTAL row
        totals['Well Count'] = pd.to_numeric(sheet_totals[well_idx], errors='coerce') \
            if well_idx < len(sheet_totals) else np.nan

    return df, totals, report_date, mapping


def _find_row(grid, col, startswith, start=0):
    target = _norm(startswith)
    for i in range(start, len(grid)):
        row = grid[i]
        if col < len(row) and _norm(row[col]).startswith(target):
            return i
    return None


def _num(v):
    n = pd.to_numeric(v, errors='coerce')
    return n


def extract_forecast_table(grid):
    """'Actual vs Forecast' block (Daily BOPD and Cumulative)."""
    r = _find_row(grid, 1, 'Actual')
    if r is None:
        return None
    head = grid[r]
    plan_label = _clean_text(head[3]) or 'Plan'
    rows = []
    for row in grid[r + 1:r + 4]:
        label = _clean_text(row[0])
        if not label:
            break
        rows.append({
            'Item': label,
            'Actual': _num(row[1]),
            plan_label: _num(row[3]),
            'Achievement (%)': _num(row[5]),
        })
    return pd.DataFrame(rows) if rows else None


def extract_field_storage_table(grid):
    """Storage / stock balance per field (Gross Opening Stock ... Gross Available Space)."""
    r = _find_row(grid, 0, 'Gross Opening Stock')
    if r is None or r == 0:
        return None
    header = grid[r - 1]
    cols = [(i, _clean_text(header[i])) for i in range(3, 11) if _clean_text(header[i])]
    unit_col = (cols[-1][0] + 1) if cols else 10
    rows = []
    for row in grid[r:r + 12]:
        label = _clean_text(row[0])
        if not label or label.startswith('*'):
            break
        rec = {'Item': label}
        for i, name in cols:
            rec[name] = _num(row[i])
        rec['Unit'] = _clean_text(row[unit_col]) if unit_col < len(row) else ''
        rows.append(rec)
    return pd.DataFrame(rows) if rows else None


def extract_nra_tables(grid):
    """
    NRA (receiving station) tables:
      - summary  : open/close stock, drain, transferred, received, safe capacity, cum. shipped
      - shipping : per-tank shipping to Qarun CPF
      - tank_stock : per-tank gross/net/quality + cumulative shipped
      - tank_closing : closing stock per NRA tank
    """
    out = {'summary': None, 'shipping': None, 'tank_stock': None, 'tank_closing': None}

    r = _find_row(grid, 0, 'Open Stock')
    if r is not None and r + 2 < len(grid):
        head, sub, first = grid[r], grid[r + 1], grid[r + 2]
        rec = {}
        for i in range(0, 5):
            rec[_clean_text(head[i])] = _num(first[i])
        for i in (14, 15):
            if i < len(head) and _clean_text(head[i]):
                rec[_clean_text(head[i])] = _num(first[i])
        out['summary'] = pd.DataFrame([rec])

        # per-tank shipping table (columns F..N)
        names = ['Tank #']
        for i in range(6, 14):
            top, low = _clean_text(head[i]), _clean_text(sub[i])
            names.append(low or top)
        rows = []
        for row in grid[r + 2:r + 12]:
            tank = _clean_text(row[5])
            if not tank:
                break
            rows.append([tank] + [row[i] for i in range(6, 14)])
            if tank.lower() == 'total':
                break
        if rows:
            df = pd.DataFrame(rows, columns=names)
            for c in df.columns[1:]:
                df[c] = pd.to_numeric(df[c], errors='coerce')
            out['shipping'] = df

    # per-tank stock table (second 'Tank #' header in column F)
    first_tank = _find_row(grid, 5, 'Tank #')
    second = _find_row(grid, 5, 'Tank #', (first_tank + 1) if first_tank is not None else 0)
    if second is not None:
        head = grid[second]
        idxs = [i for i in range(5, 16) if _clean_text(head[i])]
        names = [_clean_text(head[i]) for i in idxs]
        rows = []
        for row in grid[second + 1:second + 10]:
            tank = _clean_text(row[5])
            if not tank:
                break
            rows.append([row[i] for i in idxs])
            if tank.lower() == 'total':
                break
        if rows:
            df = pd.DataFrame(rows, columns=names)
            df[names[0]] = df[names[0]].astype(str)
            for c in names[1:]:
                df[c] = pd.to_numeric(df[c], errors='coerce')
            out['tank_stock'] = df

    r = _find_row(grid, 0, 'NRA Tank')
    if r is not None and r + 1 < len(grid):
        head, vals = grid[r], grid[r + 1]
        rec = {}
        for i in range(2, 11):
            name = _clean_text(head[i])
            if name:
                rec[name] = _num(vals[i])
        out['tank_closing'] = pd.DataFrame([rec]) if rec else None

    return out


def extract_all_report_tables(file_content):
    """Convenience wrapper used by the UI: everything except the well table."""
    grid = load_report_grid(file_content)
    tables = {
        'forecast': extract_forecast_table(grid),
        'field_storage': extract_field_storage_table(grid),
    }
    tables.update(extract_nra_tables(grid))
    return tables


def extract_wells_with_net_diff_bo(file_content):
    """
    Read the 'Report' sheet of the Detailed Production Report and return the
    well-level analysis (wells with non-zero Net Diff BO, zero-production wells, statistics).

    Works with the current layout:
      Field | Well Name | Formation | Lifting Method / Status | Gross STB | Gross diff.yest. STB |
      Net BO | Net diff.yest. BO | Flowing hours | W/C % | Water prod. BW | GAS | GOR | CUMULATIVE PROD.
    Headers are detected by name (not by fixed position), so small layout shifts don't break it.
    """
    try:
        grid = load_report_grid(file_content)
        df_all, sheet_totals, report_date, mapping = extract_report_table(grid)

        if df_all.empty:
            st.error("❌ No well rows were found under the report header.")
            return None, None, None, None, None, None

        # Clean column names used everywhere downstream
        field_col, well_name_col = 'Field', 'Well Name'
        net_bo_col, net_diff_bo_col = 'Net BO', 'Net Diff BO'
        wc_col = 'W/C %' if 'W/C %' in df_all.columns else None

        with st.expander("🔍 Detected report structure", expanded=False):
            st.caption(f"Report date: {pd.Timestamp(report_date).strftime('%d-%m-%Y') if report_date is not None else 'not found'}")
            detected = pd.DataFrame(
                [{'Excel column': openpyxl_col_letter(idx), 'Mapped to': name}
                 for key, (idx, name) in sorted(mapping.items(), key=lambda kv: kv[1][0])]
            )
            st.dataframe(detected, use_container_width=True, hide_index=True)

        if wc_col is None:
            st.warning("⚠️ Could not find the 'W/C %' column, but continuing with analysis")

        df_before_total = df_all.copy()
        for col in [net_bo_col, net_diff_bo_col] + ([wc_col] if wc_col else []):
            df_before_total[col] = pd.to_numeric(df_before_total[col], errors='coerce')

        # Reconcile with the TOTAL row written in the sheet
        sheet_net = sheet_totals.get('Net BO') if sheet_totals else None
        if sheet_net is not None and pd.notna(sheet_net):
            calc_net = df_before_total[net_bo_col].sum()
            if abs(calc_net - sheet_net) < 1:
                st.success(f"✅ Well rows reconcile with the sheet TOTAL row: {calc_net:,.0f} Net BO across {len(df_before_total)} wells")
            else:
                st.warning(f"⚠️ Sum of wells ({calc_net:,.0f}) differs from the sheet TOTAL row ({sheet_net:,.0f}) - please check the file")

        all_wells_count = len(df_before_total)
        total_net_bo_all = df_before_total[net_bo_col].sum()
        total_net_diff_bo_all = df_before_total[net_diff_bo_col].sum()

        # Field-level W/C: use the sheet's TOTAL value when present, else water / gross
        if wc_col:
            sheet_wc = sheet_totals.get('W/C %') if sheet_totals else None
            if sheet_wc is not None and pd.notna(sheet_wc):
                total_wc_all = float(sheet_wc)
            elif 'Water Prod. BW' in df_before_total.columns and 'Gross STB' in df_before_total.columns \
                    and df_before_total['Gross STB'].sum() > 0:
                total_wc_all = df_before_total['Water Prod. BW'].sum() / df_before_total['Gross STB'].sum() * 100
            else:
                total_wc_all = df_before_total[wc_col].mean()
        else:
            total_wc_all = 0

        # Wells with non-zero Net Diff BO (positive and negative)
        filtered_df = df_before_total[
            df_before_total[net_diff_bo_col].notna() & (df_before_total[net_diff_bo_col] != 0)
        ].copy()

        # Wells with ZERO Net BO
        zero_net_bo_df = df_before_total[
            df_before_total[net_bo_col].notna() & (df_before_total[net_bo_col] == 0)
        ].copy().reset_index(drop=True)

        if filtered_df.empty:
            st.warning("⚠️ No wells found with non-zero Net Diff BO values")
            return None, None, None, None, None, zero_net_bo_df

        zero_wells_count = int((df_before_total[net_diff_bo_col] == 0).sum())
        st.info(f"📊 Filtered out {zero_wells_count} wells with zero Net Diff BO values")

        positive_count = int((filtered_df[net_diff_bo_col] > 0).sum())
        negative_count = int((filtered_df[net_diff_bo_col] < 0).sum())
        st.info(f"📈 Value distribution: {positive_count} positive, {negative_count} negative Net Diff BO values")

        # Column order: the 5 core columns first (charts / PPT rely on this order), then extras
        original_columns = [field_col, well_name_col, net_bo_col, net_diff_bo_col]
        if wc_col:
            original_columns.append(wc_col)
        extra_columns = [c for c in ['Lifting Method / Status', 'Formation', 'Gross STB', 'Gross Diff STB',
                                     'Flowing Hours'] if c in df_before_total.columns]

        result_df = filtered_df[original_columns + extra_columns].copy().reset_index(drop=True)

        well_count_non_zero = len(result_df)
        total_net_bo_non_zero = result_df[net_bo_col].sum()
        total_net_diff_bo_non_zero = result_df[net_diff_bo_col].sum()
        total_wc_non_zero = result_df[wc_col].sum() if wc_col else 0

        def wc_stat(fn):
            return getattr(result_df[wc_col], fn)() if wc_col else 0

        stats = {
            # All wells
            'Total All Wells': all_wells_count,
            'Total Net BO (All Wells)': total_net_bo_all,
            'Total Net Diff BO (All Wells)': total_net_diff_bo_all,
            'Total W/C (All Wells)': total_wc_all,
            'Average Net BO (All Wells)': df_before_total[net_bo_col].mean(),
            'Average Net Diff BO (All Wells)': df_before_total[net_diff_bo_col].mean(),
            'Average W/C (All Wells)': df_before_total[wc_col].mean() if wc_col else 0,

            # Non-zero Net Diff BO wells
            'Total Wells with Non-Zero Net Diff BO': well_count_non_zero,
            'Positive Net Diff BO Wells': positive_count,
            'Negative Net Diff BO Wells': negative_count,
            'Total Net BO (Non-Zero Wells)': total_net_bo_non_zero,
            'Total Net Diff BO (Non-Zero Wells)': total_net_diff_bo_non_zero,
            'Total W/C (Non-Zero Wells)': total_wc_non_zero,
            'Average Net BO (Non-Zero Wells)': result_df[net_bo_col].mean(),
            'Average Net Diff BO (Non-Zero Wells)': result_df[net_diff_bo_col].mean(),
            'Average W/C (Non-Zero Wells)': wc_stat('mean'),
            'Maximum Net BO': result_df[net_bo_col].max(),
            'Maximum Net Diff BO': result_df[net_diff_bo_col].max(),
            'Maximum W/C': wc_stat('max'),
            'Minimum Net BO': result_df[net_bo_col].min(),
            'Minimum Net Diff BO': result_df[net_diff_bo_col].min(),
            'Minimum W/C': wc_stat('min'),
            'Median Net BO': result_df[net_bo_col].median(),
            'Median Net Diff BO': result_df[net_diff_bo_col].median(),
            'Median W/C': wc_stat('median'),
            'Standard Deviation Net BO': result_df[net_bo_col].std(),
            'Standard Deviation Net Diff BO': result_df[net_diff_bo_col].std(),
            'Standard Deviation W/C': wc_stat('std'),

            # Zero Net BO
            'Zero Net BO Wells Count': len(zero_net_bo_df),
        }

        # Extra field-level figures from the new layout
        if 'Gross STB' in df_before_total.columns:
            stats['Total Gross STB (All Wells)'] = df_before_total['Gross STB'].sum()
        if 'Lifting Method / Status' in df_before_total.columns:
            status = df_before_total['Lifting Method / Status'].str.upper()
            stats['Shut-in Wells (S/I)'] = int(status.str.replace(' ', '').isin(['S/I', 'SI']).sum())
        if report_date is not None:
            stats['Report Date'] = pd.Timestamp(report_date).strftime('%d-%m-%Y')

        # Display copy: rounded, with TOTAL row appended
        final_df = result_df.copy()
        for col in final_df.columns:
            if pd.api.types.is_float_dtype(final_df[col]):
                final_df[col] = final_df[col].round(2)

        total_row = {field_col: 'TOTAL (All Wells)',
                     well_name_col: f'{all_wells_count} Total Wells',
                     net_bo_col: total_net_bo_all,
                     net_diff_bo_col: total_net_diff_bo_all}
        if wc_col:
            total_row[wc_col] = round(total_wc_all, 2)
        if 'Gross STB' in final_df.columns:
            total_row['Gross STB'] = round(df_before_total['Gross STB'].sum(), 2)
        final_df = pd.concat([final_df, pd.DataFrame([total_row])], ignore_index=True)

        st.success(f"✅ Successfully extracted {well_count_non_zero} wells with non-zero Net Diff BO values")

        return final_df, well_count_non_zero, stats, original_columns, df_before_total, zero_net_bo_df

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None, None, None, None, None, None


def openpyxl_col_letter(idx):
    """0-based column index -> Excel letter (0 -> A)."""
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx + 1)


def _fmt_table(df):
    """Round floats for display."""
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_float_dtype(out[c]):
            out[c] = out[c].round(2)
    return out


def display_report_tables(tables, stats=None):
    """Show the additional tables of the Report sheet (forecast, storage, NRA)."""
    st.markdown("---")
    st.header("🗂️ Field Summary Tables")

    forecast = tables.get('forecast')
    if forecast is not None and not forecast.empty:
        st.subheader("🎯 Actual vs. Forecast")
        cols = st.columns(len(forecast))
        for col, (_, row) in zip(cols, forecast.iterrows()):
            ach = row.get('Achievement (%)')
            with col:
                st.metric(row['Item'], f"{row['Actual']:,.0f}",
                          delta=f"{ach:,.1f}% of plan" if pd.notna(ach) else None)
        st.dataframe(_fmt_table(forecast), use_container_width=True, hide_index=True)

    storage = tables.get('field_storage')
    if storage is not None and not storage.empty:
        st.subheader("🛢️ Field Stations - Stock Balance")
        st.dataframe(_fmt_table(storage), use_container_width=True, hide_index=True)

    summary = tables.get('summary')
    if summary is not None and not summary.empty:
        st.subheader("🚚 NRA - Daily Stock Movement")
        st.dataframe(_fmt_table(summary), use_container_width=True, hide_index=True)

    closing = tables.get('tank_closing')
    if closing is not None and not closing.empty:
        st.subheader("🛢️ NRA - Closing Stock per Tank")
        st.dataframe(_fmt_table(closing), use_container_width=True, hide_index=True)

    with st.expander("NRA tanks - shipping and quality details", expanded=False):
        shipping = tables.get('shipping')
        if shipping is not None and not shipping.empty:
            st.markdown("**Shipping to Qarun CPF (per tank)**")
            st.dataframe(_fmt_table(shipping), use_container_width=True, hide_index=True)
        tank_stock = tables.get('tank_stock')
        if tank_stock is not None and not tank_stock.empty:
            st.markdown("**Tank stock and quality**")
            st.dataframe(_fmt_table(tank_stock), use_container_width=True, hide_index=True)


def create_zero_net_bo_table(zero_net_bo_df, original_columns):
    """
    Create a modern styled table for wells with zero Net BO
    """
    if zero_net_bo_df is None or zero_net_bo_df.empty:
        return None
    
    try:
        # Extract column names
        field_col = original_columns[0]
        well_name_col = original_columns[1]
        net_bo_col = original_columns[2]
        net_diff_bo_col = original_columns[3]
        wc_col = original_columns[4] if len(original_columns) > 4 else None
        
        # Select relevant columns for display
        display_columns = [field_col, well_name_col]
        display_columns += [c for c in ['Lifting Method / Status', 'Formation', 'Gross STB']
                            if c in zero_net_bo_df.columns]
        display_columns += [net_bo_col, net_diff_bo_col]
        if wc_col and wc_col in zero_net_bo_df.columns:
            display_columns.append(wc_col)
        
        # Create display dataframe
        display_df = zero_net_bo_df[display_columns].copy()
        
        # Format numeric columns
        for col in [net_bo_col, net_diff_bo_col]:
            if col in display_df.columns and display_df[col].dtype in [np.float64, np.int64]:
                display_df[col] = display_df[col].round(2)
        
        if wc_col and wc_col in display_df.columns and display_df[wc_col].dtype in [np.float64, np.int64]:
            display_df[wc_col] = display_df[wc_col].round(2)
        
        return display_df
        
    except Exception as e:
        st.error(f"Error creating zero Net BO table: {str(e)}")
        return None

def create_visualizations(data_without_total, original_columns, all_wells_data):
    """
    Create high-resolution statistical visualizations suitable for printing
    """
    try:
        # Check if we have valid data for visualizations
        if data_without_total.empty or all_wells_data.empty:
            st.warning("No data available for visualizations")
            return None
            
        # Extract the column names
        field_col = original_columns[0]      # ('Field', 'Unnamed: 0_level_1')
        well_name_col = original_columns[1]  # ('RUNNING WELLS', 'Unnamed: 1_level_1')
        net_bo_col = original_columns[2]     # ('TOTAL PRODUCTION', 'Net\nBO')
        net_diff_bo_col = original_columns[3] # ('TOTAL PRODUCTION', 'Net diff. BO')
        wc_col = original_columns[4] if len(original_columns) > 4 else None  # ('W/C', '%')
        
        # Create clean copies for visualization
        viz_data_non_zero = data_without_total.copy()
        viz_data_all = all_wells_data.copy()
        
        # Remove rows with NaN values in the key columns for visualization
        viz_data_non_zero = viz_data_non_zero[
            viz_data_non_zero[well_name_col].notna() & 
            viz_data_non_zero[net_bo_col].notna() & 
            viz_data_non_zero[net_diff_bo_col].notna()
        ]
        
        viz_data_all = viz_data_all[
            viz_data_all[well_name_col].notna() & 
            viz_data_all[net_bo_col].notna()
        ]
        
        # Check if we have W/C data in all wells
        has_wc_data_all = wc_col and wc_col in viz_data_all.columns and viz_data_all[wc_col].notna().any()
        
        # Check if we have any data left after cleaning
        if viz_data_non_zero.empty or viz_data_all.empty:
            st.warning("No valid data available for visualizations after removing NaN values")
            return None
        
        # Extract clean data for visualization
        # Non-zero wells data
        well_names_non_zero = viz_data_non_zero[well_name_col]
        net_bo_data_non_zero = viz_data_non_zero[net_bo_col]
        net_diff_bo_data_non_zero = viz_data_non_zero[net_diff_bo_col]
        
        # All wells data
        well_names_all = viz_data_all[well_name_col]
        net_bo_data_all = viz_data_all[net_bo_col]
        wc_data_all = viz_data_all[wc_col] if has_wc_data_all else None
        
        # Check for finite values
        if (net_bo_data_non_zero.isna().all() or net_diff_bo_data_non_zero.isna().all() or 
            not np.isfinite(net_bo_data_non_zero).any() or not np.isfinite(net_diff_bo_data_non_zero).any() or
            net_bo_data_all.isna().all() or not np.isfinite(net_bo_data_all).any()):
            st.warning("No finite values available for visualization")
            return None
        
        # HIGH-RESOLUTION SETTINGS FOR PRINTING
        # Create very large figure for high resolution
        fig, axes = plt.subplots(1, 3, figsize=(36, 16))  # Increased width for better text fitting
        fig.suptitle('PRODUCTION ANALYSIS DASHBOARD', fontsize=24, fontweight='bold', y=0.98)
        
        # Set high DPI for the entire figure
        fig.set_dpi(300)
        
        # 1. Net Diff BO by Well (INCLUDING ZERO NET BO WELLS) - HIGH RESOLUTION
        if len(net_diff_bo_data_non_zero) > 0 and len(well_names_non_zero) > 0:
            # Create display data including all wells with non-zero Net Diff BO (regardless of Net BO)
            display_data = pd.DataFrame({
                'well_name': well_names_non_zero,
                'net_diff_bo': net_diff_bo_data_non_zero,
                'net_bo': net_bo_data_non_zero
            })
            
            # Sort by absolute Net Diff BO to show most significant changes first
            display_data['abs_net_diff'] = display_data['net_diff_bo'].abs()
            display_data = display_data.sort_values('abs_net_diff', ascending=False).head(15)
            
            display_wells = display_data['well_name']
            display_net_diff = display_data['net_diff_bo']
            display_net_bo = display_data['net_bo']
            
            # Create bars with optimal spacing for printing
            x_positions = np.arange(len(display_wells))
            bar_width = 0.7
            
            # Color coding: green for positive, red for negative, and special marker for zero Net BO wells
            colors = []
            edge_colors = []
            for diff, net_bo in zip(display_net_diff, display_net_bo):
                if net_bo == 0:
                    # Special color for wells with zero Net BO
                    colors.append('#ffa500')  # Orange
                    edge_colors.append('#cc8400')  # Darker orange
                elif diff >= 0:
                    colors.append('#2ecc71')  # Green
                    edge_colors.append('#27ae60')  # Darker green
                else:
                    colors.append('#e74c3c')  # Red
                    edge_colors.append('#c0392b')  # Darker red
            
            bars = axes[0].bar(x_positions, display_net_diff, 
                              width=bar_width,
                              color=colors,
                              alpha=0.85,
                              edgecolor=edge_colors,
                              linewidth=2.0)  # Thicker borders for better visibility
            
            axes[0].set_xlabel('WELLS', fontsize=18, fontweight='bold', labelpad=15)
            axes[0].set_ylabel('NET DIFF BO', fontsize=18, fontweight='bold', labelpad=15)
            axes[0].set_title('NET DIFF BO PERFORMANCE\n(Top 15 Wells by Change Magnitude)', 
                             fontsize=20, fontweight='bold', pad=25)
            axes[0].set_xticks(x_positions)
            
            # High-resolution text for well names with better spacing and bold font
            axes[0].set_xticklabels(display_wells, rotation=45, ha='right', fontsize=20, 
                                   rotation_mode='anchor', fontweight='bold')
            
            # Increase tick label size and padding
            axes[0].tick_params(axis='x', which='major', pad=15, labelsize=20)
            axes[0].tick_params(axis='y', which='major', labelsize=20)
            
            # Enhanced grid
            axes[0].grid(True, alpha=0.4, linestyle='-', linewidth=1.0, axis='y')
            
            # Adjust y limits with generous margins for printing and text placement
            y_min = display_net_diff.min() * 1.2 if display_net_diff.min() < 0 else -1
            y_max = display_net_diff.max() * 1.4 if display_net_diff.max() > 0 else 1
            axes[0].set_ylim([y_min, y_max])
            
            # Add prominent zero reference line
            axes[0].axhline(y=0, color='black', linestyle='-', alpha=0.8, linewidth=3)
            
            # High-resolution value labels placed directly on the bars
            for bar, value, net_bo in zip(bars, display_net_diff, display_net_bo):
                height = bar.get_height()
                
                # Determine text color and position based on bar value and Net BO
                if net_bo == 0:
                    text_color = 'black'
                    # For zero Net BO wells, add special marker in label
                    value_str = f'{value:.1f}*'
                else:
                    text_color = 'black'
                    value_str = f'{value:.1f}'
                
                # Position text in the middle of the bar
                if height >= 0:
                    y_pos = height * 0.7  # 70% up the bar height
                    va = 'center'
                else:
                    y_pos = height * 0.3  # 30% up from the bottom of negative bar
                    va = 'center'
                
                # Ensure text is always visible - adjust position for very small bars
                if abs(height) < (y_max - y_min) * 0.05:  # Very small bars
                    if height >= 0:
                        y_pos = height + (y_max - y_min) * 0.02  # Place slightly above
                        va = 'bottom'
                    else:
                        y_pos = height - (y_max - y_min) * 0.02  # Place slightly below
                        va = 'top'
                
                # Adjust font size and weight for better readability
                font_size = 16 if abs(height) < (y_max - y_min) * 0.1 else 18
                
                # Add value label directly on the bar
                axes[0].text(bar.get_x() + bar.get_width()/2., y_pos,
                            value_str, 
                            ha='center', 
                            va=va, 
                            fontsize=font_size, 
                            fontweight='bold',
                            color=text_color,
                            bbox=dict(boxstyle="round,pad=0.2", facecolor="white", 
                                    alpha=0.9, edgecolor='gray', linewidth=1))
            
            # Enhanced summary text for printing
            positive_count = (display_net_diff > 0).sum()
            negative_count = (display_net_diff < 0).sum()
            zero_net_bo_count = (display_net_bo == 0).sum()
            
            summary_text = f'POSITIVE: {positive_count} | NEGATIVE: {negative_count}'
            if zero_net_bo_count > 0:
                summary_text += f' | ZERO NET BO: {zero_net_bo_count}*'
            
            axes[0].text(0.02, 0.02, summary_text, 
                        transform=axes[0].transAxes, 
                        fontsize=14, 
                        color='navy',
                        fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow", 
                                alpha=0.9, edgecolor='navy', linewidth=2.0),
                        verticalalignment='bottom')
            
            # Add legend for zero Net BO wells
            if zero_net_bo_count > 0:
                axes[0].text(0.98, 0.98, '* = Zero Net BO Well', 
                            transform=axes[0].transAxes, 
                            fontsize=16, 
                            color='darkorange',
                            fontweight='bold',
                            bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", 
                                    alpha=0.9, edgecolor='darkorange', linewidth=1.5),
                            verticalalignment='top',
                            horizontalalignment='right')
            
            # Generous margins for printing
            axes[0].margins(x=0.15, y=0.25)
            
        else:
            axes[0].text(0.5, 0.5, 'NO NET DIFF BO DATA AVAILABLE', 
                        ha='center', va='center', 
                        transform=axes[0].transAxes,
                        fontsize=18,
                        fontweight='bold',
                        bbox=dict(boxstyle="round,pad=1.0", facecolor="lightgray", 
                                alpha=0.8, edgecolor='black', linewidth=2))
            axes[0].set_title('NET DIFF BO PERFORMANCE\n(Top 15 Wells by Change Magnitude)', 
                             fontsize=20, fontweight='bold')
        
        # 2. Top 10 Wells with Highest W/C values (EXCLUDING ZERO NET BO WELLS) - HIGH RESOLUTION
        if has_wc_data_all and len(wc_data_all) > 0 and len(well_names_all) > 0:
            # Create dataframe with W/C and Net BO data
            wc_analysis_data = pd.DataFrame({
                'well_name': well_names_all,
                'wc_value': wc_data_all,
                'net_bo': net_bo_data_all
            })
            
            # FILTER OUT WELLS WITH ZERO NET BO
            wc_analysis_data = wc_analysis_data[wc_analysis_data['net_bo'] > 0]
            
            if not wc_analysis_data.empty:
                # Get top 10 wells with highest W/C values (excluding zero Net BO wells)
                top_wc_wells = wc_analysis_data.nlargest(15, 'wc_value')
                
                # High-resolution horizontal bar chart
                bars = axes[1].barh(range(len(top_wc_wells)), top_wc_wells['wc_value'], 
                                   color='#3498db', alpha=0.85, edgecolor='#2980b9', linewidth=2.0)
                axes[1].set_xlabel('W/C VALUE (%)', fontsize=20, fontweight='bold', labelpad=15)
                axes[1].set_ylabel('WELLS', fontsize=28, fontweight='bold', labelpad=15)
                axes[1].set_title('TOP 15 WELLS WITH HIGHEST W/C VALUES\n(Excluding Zero Net BO Wells)', 
                                 fontsize=20, fontweight='bold', pad=25)
                axes[1].set_yticks(range(len(top_wc_wells)))
                
                # High-resolution y-axis labels with bold font
                axes[1].set_yticklabels(top_wc_wells['well_name'], fontsize=26, fontweight='bold')
                axes[1].tick_params(axis='both', which='major', labelsize=20)
                axes[1].grid(True, alpha=0.4, linestyle='-', linewidth=1.0, axis='x')
                
                # Adjust x-axis limits with generous margins
                max_wc_value = top_wc_wells['wc_value'].max()
                axes[1].set_xlim([0, max_wc_value * 1.25])  # 25% margin
                
                # High-resolution value labels with bold font
                for bar, value in zip(bars, top_wc_wells['wc_value']):
                    width = bar.get_width()
                    axes[1].text(width + max_wc_value * 0.015, bar.get_y() + bar.get_height()/2.,
                                f'{value:.1f}%', 
                                ha='left', va='center', 
                                fontsize=20, fontweight='bold',
                                bbox=dict(boxstyle="round,pad=0.4", facecolor="white", 
                                        alpha=0.95, edgecolor='gray', linewidth=1.2))     
                
                # Add note about filtering
                axes[1].text(0.02, 0.02, '✅ Excluding wells with zero Net BO', 
                            transform=axes[1].transAxes, 
                            fontsize=12, color='darkgreen', fontweight='bold',
                            bbox=dict(boxstyle="round,pad=0.4", facecolor="lightgreen", 
                                    alpha=0.9, edgecolor='darkgreen', linewidth=1.5),
                            verticalalignment='bottom')
            else:
                axes[1].text(0.5, 0.5, 'NO W/C DATA AVAILABLE\nAFTER FILTERING ZERO NET BO WELLS', 
                            ha='center', va='center', 
                            transform=axes[1].transAxes,
                            fontsize=16,
                            fontweight='bold',
                            bbox=dict(boxstyle="round,pad=1.0", facecolor="lightgray", 
                                alpha=0.8, edgecolor='black', linewidth=2))
                axes[1].set_title('TOP 15 WELLS WITH HIGHEST W/C VALUES\n(Excluding Zero Net BO Wells)', 
                                 fontsize=20, fontweight='bold')
                        
        else:
            axes[1].text(0.5, 0.5, 'NO W/C DATA AVAILABLE', 
                        ha='center', va='center', 
                        transform=axes[1].transAxes,
                        fontsize=18,
                        fontweight='bold',
                        bbox=dict(boxstyle="round,pad=1.0", facecolor="lightgray", 
                                alpha=0.8, edgecolor='black', linewidth=2))
            axes[1].set_title('TOP 15 WELLS WITH HIGHEST W/C VALUES\n(Excluding Zero Net BO Wells)', 
                             fontsize=20, fontweight='bold')
        
        # 3. Top 15 Wells with Highest Net BO (ALL WELLS) - HIGH RESOLUTION
        if len(net_bo_data_all) > 0 and len(well_names_all) > 0:
            # Get top 10 wells with highest Net BO from ALL wells
            top_wells_all = pd.DataFrame({
                'well_name': well_names_all,
                'net_bo': net_bo_data_all
            }).nlargest(15, 'net_bo')
            
            # High-resolution horizontal bar chart
            bars = axes[2].barh(range(len(top_wells_all)), top_wells_all['net_bo'], 
                               color='#f39c12', alpha=0.85, edgecolor='#e67e22', linewidth=2.0)
            axes[2].set_xlabel('NET BO', fontsize=20, fontweight='bold', labelpad=15)
            axes[2].set_ylabel('WELLS', fontsize=28, fontweight='bold', labelpad=15)
            axes[2].set_title('TOP 15 HIGHEST PRODUCING WELLS\n(All Wells)', 
                             fontsize=18, fontweight='bold', pad=25)
            axes[2].set_yticks(range(len(top_wells_all)))
            
            # High-resolution y-axis labels with bold font
            axes[2].set_yticklabels(top_wells_all['well_name'], fontsize=26, fontweight='bold')
            axes[2].tick_params(axis='both', which='major', labelsize=22)
            axes[2].grid(True, alpha=0.4, linestyle='-', linewidth=1.0, axis='x')
            
            # Adjust x-axis limits with generous margins
            max_net_bo = top_wells_all['net_bo'].max()
            axes[2].set_xlim([0, max_net_bo * 1.25])  # 25% margin
            
            # High-resolution value labels with bold font
            for bar, value in zip(bars, top_wells_all['net_bo']):
                width = bar.get_width()
                axes[2].text(width + max_net_bo * 0.015, bar.get_y() + bar.get_height()/2.,
                            f'{value:.0f}', 
                            ha='left', va='center', 
                            fontsize=22, fontweight='bold',
                            bbox=dict(boxstyle="round,pad=0.4", facecolor="white", 
                                    alpha=0.95, edgecolor='gray', linewidth=1.2))
        else:
            axes[2].text(0.5, 0.5, 'NO PRODUCTION DATA AVAILABLE', 
                        ha='center', va='center', 
                        transform=axes[2].transAxes,
                        fontsize=20,
                        fontweight='bold',
                        bbox=dict(boxstyle="round,pad=1.0", facecolor="lightgray", 
                                alpha=0.8, edgecolor='black', linewidth=2))
            axes[2].set_title('TOP 15 HIGHEST PRODUCING WELLS\n(All Wells)', 
                             fontsize=16, fontweight='bold')
        
        # HIGH-RESOLUTION LAYOUT SETTINGS
        # Adjust layout with generous padding for printing
        plt.tight_layout(pad=8.0)
        
        return fig
        
    except Exception as e:
        st.error(f"❌ Error creating visualizations: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

def _add_dataframe_slide(prs, title_text, df, font_pt=11, max_rows=10):
    """Add a slide with a DataFrame rendered as a native PowerPoint table."""
    from pptx.util import Pt
    if df is None or df.empty:
        return
    df = df.head(max_rows)
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # title only
    slide.shapes.title.text = title_text
    rows, cols = len(df) + 1, len(df.columns)
    shape = slide.shapes.add_table(rows, cols, Inches(0.4), Inches(1.5),
                                   Inches(9.2), Inches(0.4 * rows))
    table = shape.table
    for j, name in enumerate(df.columns):
        table.cell(0, j).text = str(name)
    for i, (_, row) in enumerate(df.iterrows(), 1):
        for j, name in enumerate(df.columns):
            v = row[name]
            if pd.isna(v):
                txt = ''
            elif isinstance(v, (int, float, np.integer, np.floating)):
                txt = f"{v:,.1f}" if abs(v) < 1000 else f"{v:,.0f}"
            else:
                txt = str(v)
            table.cell(i, j).text = txt
    for r in range(rows):
        for c in range(cols):
            for p in table.cell(r, c).text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(font_pt)


# =============================================================================
# TEMPLATE-BASED POWERPOINT (Operation Summary / Production Summary layout)
# =============================================================================

# Embedded copy of the Operation Summary template (used when no template file is found)
_EMBEDDED_TEMPLATE_B64 = (
    "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"
    "CP9jnrIHAAApNwAAFQAAAAAAAAAAAAAAgAE+OwAAcHB0L3NsaWRlcy9zbGlkZTIueG1sUEsBAhQDFAAAAAgA8lw4XRYaagfmAAAA"
    "VQIAACAAAAAAAAAAAAAAAIABI0MAAHBwdC9zbGlkZXMvX3JlbHMvc2xpZGUyLnhtbC5yZWxzUEsBAhQDFAAAAAgA8lw4XUP1W9Q+"
    "PwAA50EAABUAAAAAAAAAAAAAAIABR0QAAHBwdC9tZWRpYS9pbWFnZTEuanBlZ1BLAQIUAxQAAAAIAPJcOF1bHrisnwIAAGYGAAAf"
    "AAAAAAAAAAAAAACAAbiDAABwcHQvbm90ZXNTbGlkZXMvbm90ZXNTbGlkZTIueG1sUEsBAhQDFAAAAAgA8lw4XWG3yRzOAAAAvgEA"
    "ACoAAAAAAAAAAAAAAIABlIYAAHBwdC9ub3Rlc1NsaWRlcy9fcmVscy9ub3Rlc1NsaWRlMi54bWwucmVsc1BLAQIUAxQAAAAIAPJc"
    "OF0FDPmblwUAAG4dAAAhAAAAAAAAAAAAAACAAaqHAABwcHQvbm90ZXNNYXN0ZXJzL25vdGVzTWFzdGVyMS54bWxQSwECFAMUAAAA"
    "CADyXDhdO9yinbQAAAAjAQAALAAAAAAAAAAAAAAAgAGAjQAAcHB0L25vdGVzTWFzdGVycy9fcmVscy9ub3Rlc01hc3RlcjEueG1s"
    "LnJlbHNQSwECFAMUAAAACADyXDhde0O8XZwGAADPIAAAFAAAAAAAAAAAAAAAgAF+jgAAcHB0L3RoZW1lL3RoZW1lMi54bWxQSwEC"
    "FAMUAAAACADyXDhdSBSBBhcEAAAHDgAAIQAAAAAAAAAAAAAAgAFMlQAAcHB0L3NsaWRlTGF5b3V0cy9zbGlkZUxheW91dDEueG1s"
    "UEsBAhQDFAAAAAgA8lw4XYBl4Yi3AAAANgEAACwAAAAAAAAAAAAAAIABopkAAHBwdC9zbGlkZUxheW91dHMvX3JlbHMvc2xpZGVM"
    "YXlvdXQxLnhtbC5yZWxzUEsBAhQDFAAAAAgA8lw4XYGZkbMPBwAA7DEAACEAAAAAAAAAAAAAAIABo5oAAHBwdC9zbGlkZU1hc3Rl"
    "cnMvc2xpZGVNYXN0ZXIxLnhtbFBLAQIUAxQAAAAIAPJcOF0Zy/H5DQEAAMYHAAAsAAAAAAAAAAAAAACAAfGhAABwcHQvc2xpZGVN"
    "YXN0ZXJzL19yZWxzL3NsaWRlTWFzdGVyMS54bWwucmVsc1BLAQIUAxQAAAAIAPJcOF3nm11FowQAABQSAAAhAAAAAAAAAAAAAACA"
    "AUijAABwcHQvc2xpZGVMYXlvdXRzL3NsaWRlTGF5b3V0OC54bWxQSwECFAMUAAAACADyXDhdgGXhiLcAAAA2AQAALAAAAAAAAAAA"
    "AAAAgAEqqAAAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDgueG1sLnJlbHNQSwECFAMUAAAACADyXDhde+KNEFIE"
    "AAD6EAAAIQAAAAAAAAAAAAAAgAErqQAAcHB0L3NsaWRlTGF5b3V0cy9zbGlkZUxheW91dDMueG1sUEsBAhQDFAAAAAgA8lw4XYBl"
    "4Yi3AAAANgEAACwAAAAAAAAAAAAAAIABvK0AAHBwdC9zbGlkZUxheW91dHMvX3JlbHMvc2xpZGVMYXlvdXQzLnhtbC5yZWxzUEsB"
    "AhQDFAAAAAgA8lw4XcTw6nToAgAAaQcAACEAAAAAAAAAAAAAAIABva4AAHBwdC9zbGlkZUxheW91dHMvc2xpZGVMYXlvdXQ3Lnht"
    "bFBLAQIUAxQAAAAIAPJcOF2AZeGItwAAADYBAAAsAAAAAAAAAAAAAACAAeSxAABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3NsaWRl"
    "TGF5b3V0Ny54bWwucmVsc1BLAQIUAxQAAAAIAPJcOF1IOhNrcwMAAAgLAAAhAAAAAAAAAAAAAACAAeWyAABwcHQvc2xpZGVMYXlv"
    "dXRzL3NsaWRlTGF5b3V0Mi54bWxQSwECFAMUAAAACADyXDhdgGXhiLcAAAA2AQAALAAAAAAAAAAAAAAAgAGXtgAAcHB0L3NsaWRl"
    "TGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDIueG1sLnJlbHNQSwECFAMUAAAACADyXDhdI21zjhkDAACSCAAAIQAAAAAAAAAAAAAA"
    "gAGYtwAAcHB0L3NsaWRlTGF5b3V0cy9zbGlkZUxheW91dDYueG1sUEsBAhQDFAAAAAgA8lw4XYBl4Yi3AAAANgEAACwAAAAAAAAA"
    "AAAAAIAB8LoAAHBwdC9zbGlkZUxheW91dHMvX3JlbHMvc2xpZGVMYXlvdXQ2LnhtbC5yZWxzUEsBAhQDFAAAAAgA8lw4XW4SW0nD"
    "AwAAIAwAACIAAAAAAAAAAAAAAIAB8bsAAHBwdC9zbGlkZUxheW91dHMvc2xpZGVMYXlvdXQxMS54bWxQSwECFAMUAAAACADyXDhd"
    "gGXhiLcAAAA2AQAALQAAAAAAAAAAAAAAgAH0vwAAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDExLnhtbC5yZWxz"
    "UEsBAhQDFAAAAAgA8lw4XVxVrtjDBAAAFBgAACEAAAAAAAAAAAAAAIAB9sAAAHBwdC9zbGlkZUxheW91dHMvc2xpZGVMYXlvdXQ1"
    "LnhtbFBLAQIUAxQAAAAIAPJcOF2AZeGItwAAADYBAAAsAAAAAAAAAAAAAACAAfjFAABwcHQvc2xpZGVMYXlvdXRzL19yZWxzL3Ns"
    "aWRlTGF5b3V0NS54bWwucmVsc1BLAQIUAxQAAAAIAPJcOF0aCuxFjgMAAEALAAAiAAAAAAAAAAAAAACAAfnGAABwcHQvc2xpZGVM"
    "YXlvdXRzL3NsaWRlTGF5b3V0MTAueG1sUEsBAhQDFAAAAAgA8lw4XYBl4Yi3AAAANgEAAC0AAAAAAAAAAAAAAIABx8oAAHBwdC9z"
    "bGlkZUxheW91dHMvX3JlbHMvc2xpZGVMYXlvdXQxMC54bWwucmVsc1BLAQIUAxQAAAAIAPJcOF2D3RFptQMAAEUOAAAhAAAAAAAA"
    "AAAAAACAAcnLAABwcHQvc2xpZGVMYXlvdXRzL3NsaWRlTGF5b3V0NC54bWxQSwECFAMUAAAACADyXDhdgGXhiLcAAAA2AQAALAAA"
    "AAAAAAAAAAAAgAG9zwAAcHB0L3NsaWRlTGF5b3V0cy9fcmVscy9zbGlkZUxheW91dDQueG1sLnJlbHNQSwECFAMUAAAACADyXDhd"
    "UfzdqHQEAADqEQAAIQAAAAAAAAAAAAAAgAG+0AAAcHB0L3NsaWRlTGF5b3V0cy9zbGlkZUxheW91dDkueG1sUEsBAhQDFAAAAAgA"
    "8lw4XYBl4Yi3AAAANgEAACwAAAAAAAAAAAAAAIABcdUAAHBwdC9zbGlkZUxheW91dHMvX3JlbHMvc2xpZGVMYXlvdXQ5LnhtbC5y"
    "ZWxzUEsBAhQDFAAAAAgA8lw4XVWgHCpvAQAAFAMAABEAAAAAAAAAAAAAAIABctYAAHBwdC92aWV3UHJvcHMueG1sUEsBAhQDFAAA"
    "AAgA8lw4XUD7lb25FgAACiYBABUAAAAAAAAAAAAAAIABENgAAHBwdC9zbGlkZXMvc2xpZGUxLnhtbFBLAQIUAxQAAAAIAPJcOF1y"
    "pDj1IAEAAF8EAAAgAAAAAAAAAAAAAACAAfzuAABwcHQvc2xpZGVzL19yZWxzL3NsaWRlMS54bWwucmVsc1BLAQIUAxQAAAAIAPJc"
    "OF2LgK8LYAoAAGQLAAAWAAAAAAAAAAAAAACAAVrwAABwcHQvbWVkaWEvaGRwaG90bzIud2RwUEsBAhQDFAAAAAgA8lw4Xat0P1mf"
    "AgAAZgYAAB8AAAAAAAAAAAAAAIAB7voAAHBwdC9ub3Rlc1NsaWRlcy9ub3Rlc1NsaWRlMS54bWxQSwECFAMUAAAACADyXDhdSR7X"
    "RMwAAAC+AQAAKgAAAAAAAAAAAAAAgAHK/QAAcHB0L25vdGVzU2xpZGVzL19yZWxzL25vdGVzU2xpZGUxLnhtbC5yZWxzUEsBAhQD"
    "FAAAAAgA8lw4XVvuHmFeBwAA6QcAABQAAAAAAAAAAAAAAIAB3v4AAHBwdC9tZWRpYS9pbWFnZTMucG5nUEsBAhQDFAAAAAgA8lw4"
    "XVqb2b29DwAA+BAAABYAAAAAAAAAAAAAAIABbgYBAHBwdC9tZWRpYS9oZHBob3RvMS53ZHBQSwECFAMUAAAACADyXDhdhtyAtFsW"
    "AADLFgAAFAAAAAAAAAAAAAAAgAFfFgEAcHB0L21lZGlhL2ltYWdlMi5wbmdQSwECFAMUAAAACADyXDhdxluxjikCAABKBQAAEQAA"
    "AAAAAAAAAAAAgAHsLAEAcHB0L3ByZXNQcm9wcy54bWxQSwECFAMUAAAACADyXDhdYMwzkSsEAABDCwAAFQAAAAAAAAAAAAAAgAFE"
    "LwEAcHB0L3NsaWRlcy9zbGlkZTMueG1sUEsBAhQDFAAAAAgA8lw4XYuh3bTmAAAAVQIAACAAAAAAAAAAAAAAAIABojMBAHBwdC9z"
    "bGlkZXMvX3JlbHMvc2xpZGUzLnhtbC5yZWxzUEsBAhQDFAAAAAgA8lw4XfIg5zpbAwAAygkAAB8AAAAAAAAAAAAAAIABxjQBAHBw"
    "dC9ub3Rlc1NsaWRlcy9ub3Rlc1NsaWRlMy54bWxQSwECFAMUAAAACADyXDhdRtLsnc4AAAC+AQAAKgAAAAAAAAAAAAAAgAFeOAEA"
    "cHB0L25vdGVzU2xpZGVzL19yZWxzL25vdGVzU2xpZGUzLnhtbC5yZWxzUEsBAhQDFAAAAAgA8lw4XT+FuHETAgAAKAsAABMAAAAA"
    "AAAAAAAAAIABdDkBAHBwdC90YWJsZVN0eWxlcy54bWxQSwECFAMUAAAACADyXDhdGNJz7ScCAADZBQAAEAAAAAAAAAAAAAAAgAG4"
    "OwEAZG9jUHJvcHMvYXBwLnhtbFBLBQYAAAAANwA3AJgQAAANPgEAAAA="
)

TEMPLATE_GLOB = 'Operation_Summary*.pptx'


def _find_template_source():
    """
    Template priority: 1) file uploaded in the sidebar, 2) newest 'Operation_Summary*.pptx'
    that sits in the same folder as this script.
    """
    from pathlib import Path
    uploaded = st.session_state.get('ppt_template_bytes')
    if uploaded:
        return io.BytesIO(uploaded)
    try:
        here = Path(__file__).resolve().parent
    except NameError:
        here = Path.cwd()
    candidates = sorted(here.glob(TEMPLATE_GLOB), key=lambda p: p.stat().st_mtime, reverse=True)
    if candidates:
        return str(candidates[0])
    # Built-in copy of the template: the app always produces the template layout
    import base64
    return io.BytesIO(base64.b64decode(''.join(_EMBEDDED_TEMPLATE_B64)))


def _drop_shape(slide, shape):
    """Remove a shape and release the image relationships only it used."""
    el = shape._element
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rids = {v for node in el.iter() for k, v in node.attrib.items() if k.startswith('{%s}' % r_ns)}
    el.getparent().remove(el)
    for rid in rids:
        try:
            slide.part.drop_rel(rid)
        except Exception:
            pass


def _duplicate_slide(prs, src):
    """Copy a slide (shapes + image relationships) to the end of the deck."""
    import copy
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)

    rid_map = {}
    for rel in src.part.rels.values():
        if rel.is_external or rel.reltype.endswith('/slideLayout') or rel.reltype.endswith('/notesSlide'):
            continue
        rid_map[rel.rId] = new.part.relate_to(rel.target_part, rel.reltype)

    for el in src.shapes._spTree:
        if el.tag.endswith('}nvGrpSpPr') or el.tag.endswith('}grpSpPr'):
            continue
        cp = copy.deepcopy(el)
        for node in cp.iter():
            for k, v in list(node.attrib.items()):
                if k.startswith('{%s}' % r_ns) and v in rid_map:
                    node.set(k, rid_map[v])
        new.shapes._spTree.append(cp)
    return new


def _move_slide(prs, slide, new_index):
    lst = prs.slides._sldIdLst
    items = list(lst)
    for sldId in items:
        if prs.part.related_part(sldId.rId) is slide.part:
            lst.remove(sldId)
            lst.insert(new_index, sldId)
            return


def _delete_slide(prs, slide):
    lst = prs.slides._sldIdLst
    for sldId in list(lst):
        if prs.part.related_part(sldId.rId) is slide.part:
            prs.part.drop_rel(sldId.rId)
            lst.remove(sldId)
            return


def _set_text_keep_format(shape, text):
    """Replace the text of a shape but keep the formatting of its first run."""
    p = shape.text_frame.paragraphs[0]
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.text = text


def _fit_font_pt(text, base=18, thresholds=((170, 18), (260, 16), (360, 14), (500, 12), (700, 11))):
    n = len(text)
    for limit, size in thresholds:
        if n <= limit:
            return min(base, size)
    return 10


def _add_operation_cards(slide, entries, icons, mode):
    """Draw the well cards (name + rig on the left, operation text on the right)."""
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    top0, card_h, pitch = 1560000, 1050000, 1200000
    panel_l, panel_w = 169579, 2531560
    text_l, text_w = 2861783, 9162183

    for i, e in enumerate(entries):
        top = top0 + i * pitch

        # accent + light panel (same two-layer look as the template)
        accent = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(panel_l), Emu(top),
                                        Emu(2500000), Emu(card_h))
        accent.adjustments[0] = 0.0972
        accent.fill.solid(); accent.fill.fore_color.rgb = RGBColor(0x00, 0x6D, 0xFF)
        accent.line.fill.background(); accent.shadow.inherit = False
        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(panel_l + 32700), Emu(top),
                                       Emu(2500000), Emu(card_h))
        panel.adjustments[0] = 0.0764
        panel.fill.solid(); panel.fill.fore_color.rgb = RGBColor(0xF8, 0xF9, 0xFA)
        panel.line.fill.background(); panel.shadow.inherit = False

        # well name
        name = (e.get('well_name') or 'Unknown Well').upper()
        nb = slide.shapes.add_textbox(Emu(panel_l + 380000), Emu(top + 150000), Emu(2100000), Emu(420000))
        nb.text_frame.word_wrap = True
        nb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = nb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = name
        r.font.bold = True; r.font.size = Pt(21 if len(name) <= 12 else 17 if len(name) <= 16 else 14 if len(name) <= 20 else 12)
        r.font.color.rgb = RGBColor(0x16, 0x36, 0x5A)

        # rig name
        rig = e.get('rig_name') or 'Unknown Rig'
        rb = slide.shapes.add_textbox(Emu(panel_l + 380000), Emu(top + 580000), Emu(2100000), Emu(360000))
        rb.text_frame.word_wrap = True
        p = rb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = f"Rig: {rig}"
        r.font.bold = True; r.font.size = Pt(20 if len(rig) <= 10 else 16)

        # icon
        summary_text = f"{e.get('last_24_summary', '')} {e.get('next_24_forecast', '')}".upper()
        kind = 'drill' if any(k in summary_text for k in ('DRLG', 'DRILL', 'SPUD', 'REAM', 'CSG')) else 'pump'
        icon = icons.get(kind) or icons.get('pump') or icons.get('drill')
        if icon:
            blob, crop, w, h = icon
            pic = slide.shapes.add_picture(io.BytesIO(blob), Emu(panel_l + 90000),
                                           Emu(top + 130000), Emu(w), Emu(h))
            pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = crop

        # operation text box
        last = e.get('last_24_summary', 'Not Found')
        nxt = e.get('next_24_forecast', 'Not Found')
        paras = []
        if mode in ('Last 24 hours', 'Both') and last != 'Not Found':
            paras.append(last if mode == 'Last 24 hours' else f"Last 24 hrs: {last}")
        if mode in ('Next 24 hours', 'Both') and nxt != 'Not Found':
            paras.append(nxt if mode == 'Next 24 hours' else f"Next 24 hrs: {nxt}")
        if not paras:
            paras = ['No operation summary found in the report']
        paras = [t if len(t) <= 700 else t[:697] + '...' for t in paras]
        total_len = sum(len(t) for t in paras)
        size = _fit_font_pt('x' * total_len)

        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(text_l), Emu(top),
                                     Emu(text_w), Emu(card_h))
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xF0, 0xF8, 0xFF)
        box.line.color.rgb = RGBColor(0xCC, 0xE1, 0xF7)
        box.shadow.inherit = False
        tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for k, text in enumerate(paras):
            para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.LEFT
            pPr = para._p.get_or_add_pPr()
            pPr.set('marL', '285750'); pPr.set('indent', '-285750')
            from lxml import etree
            ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            bu_font = etree.SubElement(pPr, '{%s}buFont' % ns); bu_font.set('typeface', 'Arial')
            bu = etree.SubElement(pPr, '{%s}buChar' % ns); bu.set('char', '\u2022')
            run = para.add_run(); run.text = text
            run.font.size = Pt(size); run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def _fill_production_table(slide, data_df, original_columns, stats):
    """Fill the template table (Wells | NET BO | NET diff BO | W/C) and resize rows to fit."""
    import copy
    from pptx.util import Pt
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

    gf = next(sh for sh in slide.shapes if sh.has_table)
    tbl = gf.table._tbl
    rows = tbl.findall(A + 'tr')
    data_proto, total_proto = copy.deepcopy(rows[1]), copy.deepcopy(rows[-1])
    for r in rows[1:]:
        tbl.remove(r)

    well_c, net_c, diff_c = original_columns[1], original_columns[2], original_columns[3]
    wc_c = original_columns[4] if len(original_columns) > 4 else None

    body = data_df[data_df[original_columns[0]] != 'TOTAL (All Wells)']

    def fmt(v, dec=1):
        if pd.isna(v):
            return ''
        v = round(float(v), dec)
        return f"{int(v)}" if v == int(v) else f"{v:g}"

    def set_row(tr, values):
        for tc, val in zip(tr.findall(A + 'tc'), values):
            t_nodes = tc.findall('.//' + A + 't')
            if t_nodes:
                t_nodes[0].text = val
                for extra in t_nodes[1:]:
                    extra.text = ''
            elif val:
                p = tc.find('.//' + A + 'p')
                r = copy.deepcopy(next(iter(data_proto.iter(A + 'r'))))
                r.find(A + 't').text = val
                for old in p.findall(A + 'endParaRPr'):
                    p.remove(old)
                p.append(r)

    for _, row in body.iterrows():
        tr = copy.deepcopy(data_proto)
        set_row(tr, [str(row[well_c]), fmt(row[net_c], 0), fmt(row[diff_c], 0),
                     fmt(row[wc_c]) if wc_c else ''])
        tbl.append(tr)

    total_tr = copy.deepcopy(total_proto)
    set_row(total_tr, [f"{stats['Total All Wells']} Total Wells",
                       fmt(stats['Total Net BO (All Wells)'], 0),
                       fmt(stats['Total Net Diff BO (All Wells)'], 0), ''])
    tbl.append(total_tr)

    # Fit rows between the title (1.3M EMU) and the banner (5.3M EMU)
    n_rows = len(tbl.findall(A + 'tr'))
    avail_pt = 4000000 / 12700 - 6
    if n_rows <= 11:
        size, tight = 18, False
    else:
        tight = True
        size = max(7, min(18, int((avail_pt / n_rows - 3) / 1.2)))
    for idx, tr in enumerate(tbl.findall(A + 'tr')[1:], 1):
        is_total = idx == n_rows - 1
        for rpr in list(tr.iter(A + 'rPr')) + list(tr.iter(A + 'endParaRPr')):
            rpr.set('sz', str(int((min(size + 2, 20) if is_total else size) * 100)))
        if tight:
            for tcPr in tr.iter(A + 'tcPr'):
                tcPr.set('marT', '18000'); tcPr.set('marB', '18000')
    if tight:
        for tcPr in tbl.findall(A + 'tr')[0].iter(A + 'tcPr'):
            tcPr.set('marT', '18000'); tcPr.set('marB', '18000')


def _add_zero_wells_banner(slide, zero_count):
    """Native version of the red 'Zero Production Wells' banner."""
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(90000), Emu(5342493),
                                 Emu(12010000), Emu(1250000))
    shp.adjustments[0] = 0.08
    shp.shadow.inherit = False
    shp.line.fill.background()
    shp.fill.gradient()
    shp.fill.gradient_angle = 315
    stops = shp.fill.gradient_stops
    stops[0].color.rgb = RGBColor(0xFF, 0x6B, 0x6B); stops[0].position = 0
    stops[1].color.rgb = RGBColor(0xEE, 0x5A, 0x52); stops[1].position = 1

    tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Emu(250000)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = f"\U0001F6AB Zero Production Wells ({zero_count} wells)"
    r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "These wells are currently showing zero Net BO production and may require attention."
    r2.font.size = Pt(14); r2.font.color.rgb = RGBColor(255, 255, 255)


def create_template_powerpoint(template_src, data_df, stats, original_columns, visualization_fig,
                               drilling_summaries=None, ops_mode='Last 24 hours'):
    """Build the deck from the Operation Summary template (3 slide design)."""
    from pptx.util import Emu

    prs = Presentation(template_src)
    if len(prs.slides) < 3:
        raise ValueError("Template must contain 3 slides (Operation Summary, Production table, Dashboard)")
    s_ops, s_table, s_dash = prs.slides[0], prs.slides[1], prs.slides[2]

    # ---- date (only slide 1 carries it) -------------------------------------------------
    report_date = stats.get('Report Date')
    date_txt = report_date.replace('-', '/') if report_date else pd.Timestamp.now().strftime('%d/%m/%Y')

    # ---- Slide 2: production table + zero wells banner ---------------------------------
    _fill_production_table(s_table, data_df, original_columns, stats)
    for shp in [sh for sh in s_table.shapes if sh.shape_type == 13]:   # old banner screenshots
        _drop_shape(s_table, shp)
    _add_zero_wells_banner(s_table, stats.get('Zero Net BO Wells Count', 0))

    # ---- Slide 3: dashboard picture -----------------------------------------------------
    if visualization_fig is not None:
        pics = [sh for sh in s_dash.shapes if sh.shape_type == 13]
        if pics:
            big = max(pics, key=lambda s: s.width * s.height)
            box_l, box_t, box_w, box_h = big.left, big.top, big.width, big.height
            _drop_shape(s_dash, big)
        else:
            box_l, box_t, box_w, box_h = Emu(0), Emu(1110249), Emu(12086121), Emu(5248357)
        buf = io.BytesIO()
        visualization_fig.savefig(buf, format='png', dpi=200, bbox_inches='tight',
                                  facecolor='white', edgecolor='none')
        buf.seek(0)
        from PIL import Image
        iw, ih = Image.open(io.BytesIO(buf.getvalue())).size
        scale = min(box_w / iw, box_h / ih)
        w, h = int(iw * scale), int(ih * scale)
        s_dash.shapes.add_picture(buf, Emu(int(box_l + (box_w - w) / 2)), Emu(int(box_t + (box_h - h) / 2)),
                                  Emu(w), Emu(h))

    # ---- Slide 1: operation summary cards ----------------------------------------------
    entries = drilling_summaries or []
    if entries:
        # icons from the template (drilling rig / pump jack)
        icons = {}
        for sh in s_ops.shapes:
            if sh.shape_type == 13 and sh.name in ('Picture 2', 'Picture 6') and \
                    ('drill' if sh.name == 'Picture 2' else 'pump') not in icons:
                icons['drill' if sh.name == 'Picture 2' else 'pump'] = (
                    sh.image.blob, (sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom),
                    sh.width, sh.height)
        keep = {'Freeform 8', 'TextBox 12', 'TextBox 13', 'TextBox 6', 'Rectangle: Rounded Corners 2058'}
        for sh in list(s_ops.shapes):
            if sh.name not in keep:
                _drop_shape(s_ops, sh)

        pages = [entries[i:i + 4] for i in range(0, len(entries), 4)]
        slides_ops = [s_ops] + [_duplicate_slide(prs, s_ops) for _ in pages[1:]]
        for k, (sl, chunk) in enumerate(zip(slides_ops, pages)):
            if k > 0:
                _move_slide(prs, sl, k)
            for sh in sl.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip().lower().startswith('date'):
                    _set_text_keep_format(sh, f"Date : {date_txt}")
            _add_operation_cards(sl, chunk, icons, ops_mode)
    else:
        _delete_slide(prs, s_ops)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


def create_comprehensive_powerpoint(data_df, well_count, stats, original_columns, visualization_fig,
                                    extra_tables=None):
    """
    Build the PowerPoint from the Operation Summary template (sidebar upload, a file in the app
    folder, or the built-in copy). The generic layout is only a last-resort fallback on error.
    """
    template_src = _find_template_source()
    if template_src is not None:
        try:
            return create_template_powerpoint(
                template_src, data_df, stats, original_columns, visualization_fig,
                drilling_summaries=st.session_state.get('drilling_summaries') or [],
                ops_mode=st.session_state.get('ppt_ops_text', 'Last 24 hours'))
        except Exception as e:
            import traceback
            st.warning(f"⚠️ Could not use the PowerPoint template ({e}). Using the standard layout instead.")
            st.code(traceback.format_exc())
    return create_generic_powerpoint(data_df, well_count, stats, original_columns, visualization_fig, extra_tables)


def create_generic_powerpoint(data_df, well_count, stats, original_columns, visualization_fig, extra_tables=None):
    """
    Create a comprehensive PowerPoint presentation with data, statistics, and visualizations
    """
    try:
        # Keep the slide table to the core columns (extra detail columns stay in the app/Excel)
        data_df = data_df[[c for c in original_columns if c in data_df.columns]]

        # Create a new presentation
        prs = Presentation()
        
        # Title slide
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Production Analysis Report"
        subtitle.text = f"Comprehensive Well Performance Analysis\nTotal Wells: {stats['Total All Wells']}\nGenerated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}\nCreated by: Geol. Hassan Gamal Albery - Geologist @ Norpetco"
        
        # Executive Summary Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Executive Summary"
        
        # Add summary content
        content_left = Inches(0.5)
        content_top = Inches(1.5)
        content_width = Inches(9.0)
        content_height = Inches(5.0)
        
        text_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        # Add summary points
        summary_points = [
            f"• Total Wells Analyzed: {stats['Total All Wells']}",
            f"• Wells with Non-Zero Net Diff BO: {stats['Total Wells with Non-Zero Net Diff BO']}",
            f"• Positive Performance Wells: {stats['Positive Net Diff BO Wells']}",
            f"• Wells Requiring Attention: {stats['Negative Net Diff BO Wells']}",
            f"• Total Net BO Production: {stats['Total Net BO (All Wells)']:,.0f}",
            f"• Average Net BO per Well: {stats['Average Net BO (All Wells)']:,.0f}",
            f"• Highest Producing Well: {stats['Maximum Net BO']:,.0f}",
            f"• Performance Range: {stats['Minimum Net BO']:,.0f} to {stats['Maximum Net BO']:,.0f}"
        ]
        
        # Add W/C statistics if available
        if stats['Total W/C (All Wells)'] != 0:
            summary_points.extend([
                f"• Total W/C: {stats['Total W/C (All Wells)']:,.2f}%",
                f"• Average W/C: {stats['Average W/C (All Wells)']:,.2f}%"
            ])
        
        for point in summary_points:
            p = text_frame.add_paragraph()
            p.text = point
            p.space_after = Inches(0.05)
        
        # Main Data Table Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Production Data - Key Wells"
        
        # Create main data table (show only first 15 rows for readability, including TOTAL row if present)
        display_data = data_df.head(15) if len(data_df) > 15 else data_df
        
        rows = len(display_data) + 1
        cols = len(display_data.columns)
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9.0)
        height = Inches(0.8 * min(rows, 12))  # Limit height
        
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        
        # Set column headers
        for i, column in enumerate(display_data.columns):
            table.cell(0, i).text = str(column)
        
        # Fill table with data
        for row_idx, (_, row_data) in enumerate(display_data.iterrows(), 1):
            for col_idx, column in enumerate(display_data.columns):
                value = row_data[column]
                if isinstance(value, (int, float)) and column not in [original_columns[0], original_columns[1]]:
                    table.cell(row_idx, col_idx).text = f"{value:,.2f}"
                else:
                    table.cell(row_idx, col_idx).text = str(value)
        
        # Key Metrics Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Key Performance Metrics"
        
        # Create key metrics table
        key_metrics = {
            'Total Wells': stats['Total All Wells'],
            'Wells with Significant Changes': stats['Total Wells with Non-Zero Net Diff BO'],
            'Positive Performance Wells': stats['Positive Net Diff BO Wells'],
            'Wells Requiring Attention': stats['Negative Net Diff BO Wells'],
            'Total Net BO Production': stats['Total Net BO (All Wells)'],
            'Total Net Diff BO': stats['Total Net Diff BO (All Wells)'],
            'Average Net BO per Well': stats['Average Net BO (All Wells)'],
            'Highest Producing Well': stats['Maximum Net BO'],
            'Performance Standard Deviation': stats['Standard Deviation Net BO']
        }
        
        # Add W/C metrics if available
        if stats['Total W/C (All Wells)'] != 0:
            key_metrics.update({
                'Total W/C': stats['Total W/C (All Wells)'],
                'Average W/C': stats['Average W/C (All Wells)']
            })
        
        stats_rows = len(key_metrics) + 1
        stats_cols = 2
        left = Inches(1.0)
        top = Inches(1.5)
        width = Inches(8.0)
        height = Inches(0.8 * min(stats_rows, 15))
        
        stats_table = slide.shapes.add_table(stats_rows, stats_cols, left, top, width, height).table
        stats_table.cell(0, 0).text = "Metric"
        stats_table.cell(0, 1).text = "Value"
        
        for idx, (metric, value) in enumerate(key_metrics.items(), 1):
            stats_table.cell(idx, 0).text = metric
            if isinstance(value, (int, float)):
                if value > 1000:
                    stats_table.cell(idx, 1).text = f"{value:,.0f}"
                else:
                    stats_table.cell(idx, 1).text = f"{value:,.2f}"
            else:
                stats_table.cell(idx, 1).text = str(value)
        
        # Visualization Slides
        if visualization_fig:
            # Save figure to bytes with HIGH RESOLUTION settings
            img_buffer = io.BytesIO()
            visualization_fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight', 
                                     facecolor='white', edgecolor='none')
            img_buffer.seek(0)
            
            # Create individual visualization slides
            visualization_titles = [
                "Net Diff BO Performance",
                "Top 10 Wells with Highest W/C Values", 
                "Top 10 Highest Producing Wells"
            ]
            
            for viz_title in visualization_titles:
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                title.text = f"Analysis - {viz_title}"
                
                # Add the HIGH-RESOLUTION visualization image
                left = Inches(0.5)
                top = Inches(1.0)
                width = Inches(9.0)
                slide.shapes.add_picture(img_buffer, left, top, width=width)
        
        # Report tables (forecast / stock / NRA)
        if extra_tables:
            _add_dataframe_slide(prs, "Actual vs. Forecast", extra_tables.get('forecast'), font_pt=12)
            storage = extra_tables.get('field_storage')
            if storage is not None and not storage.empty:
                _add_dataframe_slide(prs, "Field Stations - Stock Balance",
                                     storage.drop(columns=['Unit'], errors='ignore').round(1), font_pt=9)
            _add_dataframe_slide(prs, "NRA - Daily Stock Movement", extra_tables.get('summary'), font_pt=11)
            _add_dataframe_slide(prs, "NRA - Closing Stock per Tank", extra_tables.get('tank_closing'), font_pt=10)

        # Recommendations Slide
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Recommendations & Next Steps"
        
        text_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        recommendations = [
            "🎯 Focus Areas:",
            "• Analyze top performing wells for best practices replication",
            "• Review wells with negative Net Diff BO for improvement opportunities",
            "• Monitor wells with significant performance deviations",
            "",
            "📊 Operational Actions:",
            "• Optimize production parameters for underperforming wells",
            "• Implement preventive maintenance for critical wells",
            "• Share best practices from top performers",
            "",
            "📈 Continuous Improvement:",
            "• Regular monitoring of Net Diff BO trends",
            "• Periodic review of well performance categories",
            "• Update operational strategies based on performance data"
        ]
        
        for recommendation in recommendations:
            p = text_frame.add_paragraph()
            p.text = recommendation
            p.space_after = Inches(0.03)
        
        # Save to bytes buffer
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        return ppt_buffer
        
    except Exception as e:
        st.error(f"❌ Error creating PowerPoint: {str(e)}")
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

def create_excel_with_visualizations(data_df, stats, visualization_fig, extra_tables=None):
    """
    Create an Excel file with data, statistics, and embedded visualizations
    """
    try:
        # Create Excel writer
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            # Write main data (include TOTAL row)
            data_df.to_excel(writer, sheet_name='Production Data', index=False)
            
            # Write statistics
            stats_df = pd.DataFrame(list(stats.items()), columns=['Metric', 'Value'])
            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
            
            # Additional report tables (forecast, storage, NRA)
            sheet_names = {'forecast': 'Actual vs Forecast', 'field_storage': 'Field Stock Balance',
                           'summary': 'NRA Summary', 'tank_closing': 'NRA Closing Stock',
                           'shipping': 'NRA Shipping', 'tank_stock': 'NRA Tank Quality'}
            for key, sheet in sheet_names.items():
                tbl = (extra_tables or {}).get(key)
                if tbl is not None and not tbl.empty:
                    tbl.to_excel(writer, sheet_name=sheet, index=False)
                    writer.sheets[sheet].set_column('A:Z', 22)

            # Get workbook and worksheets
            workbook = writer.book
            
            # Format worksheets
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Format data sheet
            data_sheet = writer.sheets['Production Data']
            for col_num, value in enumerate(data_df.columns.values):
                data_sheet.write(0, col_num, str(value), header_format)
            data_sheet.set_column('A:Z', 15)
            
            # Format statistics sheet
            stats_sheet = writer.sheets['Statistics']
            stats_sheet.write(0, 0, 'Metric', header_format)
            stats_sheet.write(0, 1, 'Value', header_format)
            stats_sheet.set_column('A:A', 35)
            stats_sheet.set_column('B:B', 20)
            
            # Add visualization if available
            if visualization_fig:
                # Save figure to bytes with HIGH RESOLUTION settings
                img_buffer = io.BytesIO()
                visualization_fig.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight',
                                         facecolor='white', edgecolor='none')
                img_buffer.seek(0)
                
                # Create visualization sheet
                viz_sheet = workbook.add_worksheet('Visualizations')
                
                # Insert the HIGH-RESOLUTION image
                viz_sheet.insert_image('A1', 'visualization.png', {'image_data': img_buffer, 'x_scale': 0.8, 'y_scale': 0.8})
                viz_sheet.set_column('A:A', 60)
                viz_sheet.set_row(0, 400)
        
        excel_buffer.seek(0)
        return excel_buffer
        
    except Exception as e:
        st.error(f"❌ Error creating Excel file: {str(e)}")
        return None

# =============================================================================
# DRILLING REPORTS UPLOAD FUNCTIONS
# =============================================================================

def extract_operation_summary_from_excel(uploaded_file):
    """
    Extract operation summary, well name, and rig name from uploaded Excel file
    """
    try:
        # Read the Excel file
        uploaded_file.seek(0)
        wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=True)
        sheet = wb.active
        
        # Initialize variables
        well_name = ""
        rig_name = ""
        last_24_summary = ""
        next_24_forecast = ""
        
        # Search for well name
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "WELL NAME" in str(cell).upper():
                    # Get the well name from adjacent cells
                    if i + 1 < len(row) and row[i + 1]:
                        well_name = str(row[i + 1])
                        break
                    # Also check other cells in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "WELL NAME" not in str(cell2).upper() and cell2:
                            well_name = str(cell2)
                            break
                    break
        
        # Search for rig name
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "RIG NAME" in str(cell).upper():
                    # Get the rig name from adjacent cells
                    if i + 1 < len(row) and row[i + 1]:
                        rig_name = str(row[i + 1])
                        break
                    # Also check other cells in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "RIG NAME" not in str(cell2).upper() and cell2:
                            rig_name = str(cell2)
                            break
                    break
        
        # Search for LAST 24 SUMMARY
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "LAST 24 SUMMARY" in str(cell).upper():
                    # Get the summary from the next cell
                    if i + 1 < len(row) and row[i + 1]:
                        last_24_summary = str(row[i + 1])
                        break
                    # If not in next cell, try to find in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "LAST 24 SUMMARY" not in str(cell2).upper() and cell2:
                            last_24_summary = str(cell2)
                            break
                    break
        
        # Search for NEXT 24 FORECAST
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "NEXT 24 FORECAST" in str(cell).upper():
                    # Get the forecast from the next cell
                    if i + 1 < len(row) and row[i + 1]:
                        next_24_forecast = str(row[i + 1])
                        break
                    # If not in next cell, try to find in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "NEXT 24 FORECAST" not in str(cell2).upper() and cell2:
                            next_24_forecast = str(cell2)
                            break
                    break
        
        # Clean up the extracted data
        well_name = well_name.replace(':-', '').replace(':', '').strip() if well_name else "Not Found"
        rig_name = rig_name.replace(':-', '').replace(':', '').strip() if rig_name else "Not Found"
        last_24_summary = last_24_summary.replace(':-', '').replace(':', '').strip() if last_24_summary else "Not Found"
        next_24_forecast = next_24_forecast.replace(':-', '').replace(':', '').strip() if next_24_forecast else "Not Found"
        
        return {
            'file_name': uploaded_file.name,
            'well_name': well_name,
            'rig_name': rig_name,
            'last_24_summary': last_24_summary,
            'next_24_forecast': next_24_forecast
        }
        
    except Exception as e:
        st.error(f"Error processing file {uploaded_file.name}: {str(e)}")
        return None

def create_operation_summary_display(last_24_summary, next_24_forecast):
    """
    Create a formatted operation summary for display
    """
    if last_24_summary == "Not Found" and next_24_forecast == "Not Found":
        return "❌ No operation summary found in file"
    
    summary_html = f"""
    <div style="padding: 10px; border-radius: 5px; background-color: #f0f8ff;">
        <div style="margin-bottom: 15px;">
            <h4 style="margin: 0; color: #1f77b4; font-size: 14px;">📅 LAST 24 HOURS:</h4>
            <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">{last_24_summary if last_24_summary != 'Not Found' else 'No data available'}</p>
        </div>
        <div>
            <h4 style="margin: 0; color: #2ca02c; font-size: 14px;">🔮 NEXT 24 HOURS:</h4>
            <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">{next_24_forecast if next_24_forecast != 'Not Found' else 'No data available'}</p>
        </div>
    </div>
    """
    return summary_html

def collect_drilling_summaries(uploaded_files):
    """Read all uploaded drilling reports and share the result with the PowerPoint builder."""
    if not uploaded_files:
        st.session_state['drilling_summaries'] = []
        return []
    summaries = []
    with st.spinner("🔍 Analyzing drilling reports..."):
        for f in uploaded_files:
            s = extract_operation_summary_from_excel(f)
            if s:
                summaries.append(s)
    st.session_state['drilling_summaries'] = summaries
    return summaries


def drilling_reports_section(uploaded_files):
    """Drilling operations results (shown on the same page as the production analysis)."""
    if uploaded_files:
        st.markdown("---")
        st.header("🏗️ Drilling Operations Summary")
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")

        all_summaries = st.session_state.get('drilling_summaries', [])

        if all_summaries:
            # Create the main summary table with two columns
            st.subheader("📊 Operations Summary")
            st.markdown("### Current Drilling Operations Overview")
            
            # Display statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📁 Total Files", len(all_summaries))
            with col2:
                unique_wells = len(set([s['well_name'] for s in all_summaries if s['well_name'] != "Not Found"]))
                st.metric("🛢️ Active Wells", unique_wells)
            with col3:
                unique_rigs = len(set([s['rig_name'] for s in all_summaries if s['rig_name'] != "Not Found"]))
                st.metric("🔧 Active Rigs", unique_rigs)
            
            # Create the main two-column display
            for i, summary in enumerate(all_summaries):
                # Create a container for each row
                with st.container():
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        # Rig and Well information
                        st.markdown(f"""
                        <div style="padding: 15px; background-color: #f8f9fa; border-radius: 10px; border-left: 4px solid #007bff;">
                            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">{summary['well_name'] if summary['well_name'] != 'Not Found' else 'Unknown Well'}</h3>
                            <p style="margin: 0; color: #7f8c8d; font-size: 14px;">
                                <strong>Rig:</strong> {summary['rig_name'] if summary['rig_name'] != 'Not Found' else 'Unknown Rig'}
                            </p>
                            <p style="margin: 5px 0 0 0; color: #95a5a6; font-size: 12px;">
                                File: {summary['file_name']}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        # Operation summary
                        operation_display = create_operation_summary_display(
                            summary['last_24_summary'], 
                            summary['next_24_forecast']
                        )
                        st.markdown(operation_display, unsafe_allow_html=True)
                    
                    # Add some spacing between entries
                    st.markdown("<br>", unsafe_allow_html=True)
            
            # Detailed expandable sections
            st.subheader("🔍 Detailed Operation Views")
            st.markdown("Click on any operation below to see full details:")
            
            for i, summary in enumerate(all_summaries):
                with st.expander(f"🔧 {summary['well_name']} - {summary['rig_name']} | 📄 {summary['file_name']}", expanded=False):
                    
                    # Create two columns for detailed view
                    detail_col1, detail_col2 = st.columns(2)
                    
                    with detail_col1:
                        st.markdown("### 📋 Well & Rig Information")
                        st.info(f"""
                        **Well Name:** {summary['well_name'] if summary['well_name'] != 'Not Found' else '❌ Not found'}
                        \n**Rig Name:** {summary['rig_name'] if summary['rig_name'] != 'Not Found' else '❌ Not found'}
                        \n**Source File:** {summary['file_name']}
                        """)
                    
                    with detail_col2:
                        st.markdown("### 📊 Operation Status")
                        if summary['last_24_summary'] != "Not Found":
                            st.success("✅ Operations data successfully extracted")
                        else:
                            st.warning("⚠️ Limited operation data available")
                    
                    # Operation details in full width
                    st.markdown("### 🕐 Operation Details")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 📅 Last 24 Hours")
                        if summary['last_24_summary'] != "Not Found":
                            st.info(summary['last_24_summary'])
                        else:
                            st.warning("No last 24 hours summary found")
                    
                    with col2:
                        st.markdown("#### 🔮 Next 24 Hours")
                        if summary['next_24_forecast'] != "Not Found":
                            st.success(summary['next_24_forecast'])
                        else:
                            st.warning("No next 24 hours forecast found")
                    
                    st.markdown("---")
            
            # Download section
            st.subheader("💾 Export Data")
            
            # Prepare data for download
            download_data = []
            for summary in all_summaries:
                download_data.append({
                    'Well Name': summary['well_name'],
                    'Rig Name': summary['rig_name'],
                    'Last 24 Hours Summary': summary['last_24_summary'],
                    'Next 24 Hours Forecast': summary['next_24_forecast'],
                    'Source File': summary['file_name']
                })
            
            download_df = pd.DataFrame(download_data)
            csv = download_df.to_csv(index=False)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Summary as CSV",
                    data=csv,
                    file_name="drilling_operations_summary.csv",
                    mime="text/csv",
                    help="Download all operation summaries as a CSV file"
                )
            with col2:
                # Fix for Excel download - actually create Excel file
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    download_df.to_excel(writer, index=False, sheet_name='Drilling Operations')
                excel_buffer.seek(0)
                
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=excel_buffer,
                    file_name="drilling_operations_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download all operation summaries as an Excel file"
                )
            
        else:
            st.error("❌ No valid operation summaries could be extracted from the uploaded files.")
            st.info("💡 Please make sure your Excel files contain the required fields: WELL NAME, RIG NAME, LAST 24 SUMMARY, and NEXT 24 FORECAST")

def render_sidebar():
    """Sidebar: developer info, quick start, template and tools."""
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; margin-bottom: 1rem;">
            <span style="font-size: 3rem;">📊</span>
            <h2>Production Analytics</h2>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="sidebar-developer">
        <h4>👨‍💻 Developed by</h4>
        <h3>Geol. Hassan Gamal Albery</h3>
        <p>Geologist @ Norpetco</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("🚀 Quick Start")
        st.markdown("""
        1. **Upload** the Detailed Production Report
        2. **Upload** the drilling reports (optional)
        3. **Review** the analysis
        4. **Download** reports / PowerPoint
        """)

        st.markdown("---")
        st.subheader("📋 Supported Files")
        st.markdown("""
        • Production: .xlsx, .xls, .xlsm
        • Drilling reports: .xlsx
        """)

        st.markdown("---")
        st.subheader("🎨 PowerPoint Template")
        tpl_file = st.file_uploader("Use a different template (optional)", type=['pptx'], key="ppt_template_uploader")
        if tpl_file is not None:
            st.session_state['ppt_template_bytes'] = tpl_file.getvalue()
        else:
            st.session_state.pop('ppt_template_bytes', None)
        st.caption("Default: Operation_Summary*.pptx in the app folder if present, otherwise the built-in template.")

        st.markdown("---")
        st.subheader("🛠️ Tools")
        if st.button("🔄 Clear Cache & Refresh", use_container_width=True):
            st.cache_data.clear()
            st.success("✅ Application refreshed!")


def render_intro():
    """Short description shown before anything is uploaded."""
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
        <div class="info-box">
        <h3>🎯 What This Dashboard Does</h3>
        <p>Upload the Detailed Production Report (and optionally the drilling reports) to get:</p>
        <ul>
        <li><b>Well Performance Insights</b> - Identify top performers and areas for improvement</li>
        <li><b>Production Trends</b> - Track Net BO and Net Diff BO metrics</li>
        <li><b>W/C Analysis</b> - Monitor water cut percentages for each well</li>
        <li><b>Operation Summary</b> - Last / next 24 hours for every drilling well</li>
        <li><b>One-click PowerPoint</b> - Built on the Operation Summary template</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="info-box">
        <h3>📈 Key Features</h3>
        <p>• Automated Data Extraction<br>
           • Smart Column Detection<br>
           • Field / Stock / NRA Tables<br>
           • Interactive Visualizations<br>
           • Multi-format Export<br>
           • Professional Reporting</p>
        </div>
        """, unsafe_allow_html=True)


def render_upload_area():
    """Both uploaders side by side. Returns (production_file, drilling_files)."""
    st.markdown("---")
    col_prod, col_drill = st.columns(2)

    with col_prod:
        st.markdown('<div class="upload-section"><h3>📁 Detailed Production Report</h3>'
                    '<p>Excel file with the "Report" sheet</p></div>', unsafe_allow_html=True)
        production_file = st.file_uploader(
            "Choose your production Excel file",
            type=['xlsx', 'xls', 'xlsm'],
            help="Upload the Detailed Production Report. The app detects the columns automatically.",
            label_visibility="collapsed",
            key="production_uploader"
        )

    with col_drill:
        st.markdown('<div class="upload-section"><h3>🏗️ Drilling Reports</h3>'
                    '<p>One or more daily drilling report files</p></div>', unsafe_allow_html=True)
        drilling_files = st.file_uploader(
            "Choose drilling report files",
            type=['xlsx'],
            accept_multiple_files=True,
            help="Upload one or more drilling report Excel files (optional)",
            label_visibility="collapsed",
            key="drilling_uploader"
        )
        st.radio(
            "Text shown on the PowerPoint 'Operation Summary' slide",
            ['Last 24 hours', 'Next 24 hours', 'Both'],
            horizontal=True, key='ppt_ops_text')

    return production_file, drilling_files


def production_analysis_section(uploaded_file, show_guide=True):
    """Production analysis results for the uploaded Detailed Production Report."""
    if uploaded_file is not None:
        try:
            # Process file without toggle status
            with st.spinner("🔄 Processing your file... This may take a few moments."):
                result_df, well_count, stats, original_columns, all_wells_data, zero_net_bo_df = extract_wells_with_net_diff_bo(uploaded_file)
                
                if result_df is not None and not result_df.empty:
                    # Additional tables from the Report sheet (forecast, storage, NRA)
                    try:
                        report_tables = extract_all_report_tables(uploaded_file)
                    except Exception as tbl_err:
                        report_tables = {}
                        st.warning(f"⚠️ Could not read the summary tables: {tbl_err}")

                    # Generate visualizations (exclude TOTAL row for visualization)
                    data_without_total = result_df[result_df[original_columns[0]] != 'TOTAL (All Wells)']
                    fig = create_visualizations(data_without_total, original_columns, all_wells_data)
                    
                    # Generate PowerPoint automatically
                    ppt_buffer = create_comprehensive_powerpoint(result_df, well_count, stats, original_columns, fig, report_tables)
                    
                    # Success message
                    st.markdown(f"""
                    <div class="success-box">
                    <h3>✅ Analysis Complete!</h3>
                    <p>Successfully processed <b>{stats['Total All Wells']}</b> total wells and identified <b>{well_count}</b> wells with significant Net Diff BO values.</p>
                    <p><b>PowerPoint report has been automatically generated and is ready for download below.</b></p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Enhanced metrics display
                    st.markdown("---")
                    st.header("📊 Key Performance Indicators")
                    
                    # Check if W/C data is available
                    has_wc_data = stats['Total W/C (All Wells)'] != 0
                    
                    if has_wc_data:
                        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                        
                        with kpi1:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Total Wells</h3>
                            <h2>{stats['Total All Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi2:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Wells with Changes</h3>
                            <h2>{stats['Total Wells with Non-Zero Net Diff BO']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi3:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Positive Performance</h3>
                            <h2>{stats['Positive Net Diff BO Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi4:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Needs Attention</h3>
                            <h2>{stats['Negative Net Diff BO Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi5:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Zero Net BO Wells</h3>
                            <h2>{stats['Zero Net BO Wells Count']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                    else:
                        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                        
                        with kpi1:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Total Wells</h3>
                            <h2>{stats['Total All Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi2:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Wells with Changes</h3>
                            <h2>{stats['Total Wells with Non-Zero Net Diff BO']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi3:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Positive Performance</h3>
                            <h2>{stats['Positive Net Diff BO Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi4:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Needs Attention</h3>
                            <h2>{stats['Negative Net Diff BO Wells']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with kpi5:
                            st.markdown(f"""
                            <div class="metric-card">
                            <h3>Zero Net BO Wells</h3>
                            <h2>{stats['Zero Net BO Wells Count']}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                    
                    # Field summary tables (forecast / stock / NRA)
                    if report_tables:
                        display_report_tables(report_tables, stats)

                    # Data preview (with TOTAL row included)
                    st.markdown("---")
                    st.header("📋 Production Data Overview")
                    st.dataframe(result_df, use_container_width=True, height=400)
                    
                    # NEW: Zero Net BO Wells Table
                    if zero_net_bo_df is not None and not zero_net_bo_df.empty:
                        st.markdown("---")
                        st.header("⚠️ Wells with Zero Net BO")
                        
                        # Create the modern styled table for zero Net BO wells
                        zero_net_bo_display = create_zero_net_bo_table(zero_net_bo_df, original_columns)
                        
                        if zero_net_bo_display is not None:
                            # Add custom CSS for the zero Net BO table
                            st.markdown("""
                            <style>
                            .zero-net-bo-table {
                                background: linear-gradient(135deg, #ff6b6b 0%, #ee5a52 100%);
                                border-radius: 10px;
                                padding: 20px;
                                color: white;
                                margin-bottom: 20px;
                            }
                            .zero-net-bo-table h3 {
                                color: white;
                                margin-bottom: 15px;
                            }
                            .dataframe {
                                border: none !important;
                            }
                            .dataframe thead th {
                                background-color: #c44569 !important;
                                color: white !important;
                                font-weight: bold !important;
                            }
                            .dataframe tbody tr:nth-child(even) {
                                background-color: #ff7979 !important;
                            }
                            .dataframe tbody tr:nth-child(odd) {
                                background-color: #ff8c8c !important;
                            }
                            .dataframe tbody tr:hover {
                                background-color: #e66767 !important;
                            }
                            </style>
                            """, unsafe_allow_html=True)
                            
                            # Display the table with a custom header
                            st.markdown(f"""
                            <div class="zero-net-bo-table">
                                <h3>🚫 Zero Production Wells ({len(zero_net_bo_display)} wells)</h3>
                                <p>These wells are currently showing zero Net BO production and may require attention.</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            # Display the dataframe with custom styling
                            st.dataframe(zero_net_bo_display, use_container_width=True, height=400)
                            
                            # Add download button for zero Net BO wells
                            st.download_button(
                                label="📥 Download Zero Net BO Wells as CSV",
                                data=zero_net_bo_display.to_csv(index=False),
                                file_name="zero_net_bo_wells.csv",
                                mime="text/csv",
                                use_container_width=True
                            )
                        else:
                            st.info("📊 No wells with zero Net BO found in the dataset.")
                    
                    # Visualizations
                    st.markdown("---")
                    st.header("📈 Performance Analytics")
                    if fig:
                        st.pyplot(fig)
                        st.caption("Figure 1: High-resolution production performance analysis - Suitable for printing")
                    else:
                        st.info("📊 Visualizations not available due to insufficient data")
                    
                    # Enhanced Download section
                    st.markdown("---")
                    st.header("💾 Download Reports")
                    
                    st.markdown("""
                    <div class="info-box">
                    <h3>🎁 Export Your Analysis</h3>
                    <p>Choose from multiple formats to share your insights with your team:</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    download_col1, download_col2, download_col3 = st.columns(3)
                    
                    with download_col1:
                        st.subheader("📄 CSV Export")
                        st.markdown("Simple data format for spreadsheets")
                        # Export with TOTAL row included
                        csv = result_df.to_csv(index=False)
                        st.download_button(
                            label="📥 Download CSV",
                            data=csv,
                            file_name="production_analysis.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                    
                    with download_col2:
                        st.subheader("📊 Excel Report")
                        st.markdown("Complete analysis with charts")
                        if st.button("🔄 Generate Excel Report", use_container_width=True, key="excel_gen"):
                            with st.spinner("Creating comprehensive Excel report..."):
                                excel_buffer = create_excel_with_visualizations(result_df, stats, fig, report_tables)
                            
                            if excel_buffer:
                                st.download_button(
                                    label="📥 Download Excel",
                                    data=excel_buffer,
                                    file_name="production_analysis.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="excel_download"
                                )
                            else:
                                st.error("❌ Failed to create Excel report")
                    
                    with download_col3:
                        st.subheader("🎤 PowerPoint")
                        st.markdown("Professional presentation")
                        if ppt_buffer:
                            st.download_button(
                                label="📥 Download PowerPoint",
                                data=ppt_buffer,
                                file_name="production_presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                                key="ppt_download"
                            )
                            st.success("✅ PowerPoint ready for download!")
                        else:
                            st.error("❌ Failed to create PowerPoint presentation")
                
                else:
                    st.error("❌ No valid data found in the uploaded file. Please check your file format and try again.")
                    
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.markdown("""
            <div class="warning-box">
            <h3>💡 Troubleshooting Tips</h3>
            <ul>
            <li>Ensure your Excel file has data in the 'Report' worksheet</li>
            <li>Check that the header block has 'Well Name' in the sub-header row</li>
            <li>Verify that required columns are present (Field, Well Name, Net BO, Net diff.yest. BO)</li>
            <li>Try saving your file as .xlsx format if issues persist</li>
            </ul>
            </div>
            """, unsafe_allow_html=True)
    
    elif show_guide:
        # Enhanced instructions when no file is uploaded
        st.markdown("---")
        st.header("📖 Getting Started Guide")
        
        guide_col1, guide_col2 = st.columns(2)
        
        with guide_col1:
            st.subheader("🎯 Step-by-Step Process")
            steps = [
                {"step": "1", "title": "Prepare Your Data", "desc": "Ensure your Excel file has production data in the 'Report' sheet with proper headers"},
                {"step": "2", "title": "Upload File", "desc": "Use the upload section above to select your Excel file (.xlsx, .xls, or .xlsm)"},
                {"step": "3", "title": "Automatic Analysis", "desc": "The app will automatically detect columns and process your data"},
                {"step": "4", "title": "Review Results", "desc": "Examine the insights, visualizations, and key metrics"},
                {"step": "5", "title": "Export Reports", "desc": "Download your analysis in CSV, Excel, or PowerPoint format"}
            ]
            
            for step in steps:
                with st.container():
                    st.markdown(f"**{step['step']}. {step['title']}**")
                    st.caption(step['desc'])
                    st.markdown("---")
        
        with guide_col2:
            st.subheader("📋 File Requirements")
            requirements = [
                "✅ **File Types**: .xlsx, .xls, or .xlsm (Macro-enabled Excel)",
                "✅ **Worksheet**: Data must be in 'Report' sheet",
                "✅ **Headers**: Detected automatically by name (Field, Well Name, Net BO, Net diff.yest. BO, W/C %, ...)",
                "✅ **Required Columns**:",
                "   - Field column (fields may be merged cells)",
                "   - Well Name column", 
                "   - Net BO production values",
                "   - Net Diff BO performance values",
                "   - W/C percentage values (if available)",
                "✅ **Data Format**: Wells are read until the 'CUM. PROD.' / 'TOTAL' row",
                "✅ **Extra tables read**: Actual vs Forecast, field stock balance, NRA tanks"
            ]
            
            for req in requirements:
                st.markdown(req)
            
            st.markdown("---")
            st.subheader("🔍 Expected Output")
            st.markdown("""
            • **Data Table**: Non-zero Net Diff BO wells with status, formation, W/C and TOTAL\n            • **Field Tables**: Actual vs Forecast, stock balance, NRA tanks
            • **Zero Net BO Table**: Special table highlighting wells with zero production
            • **Key Metrics**: Performance statistics including W/C analysis
            • **Visual Charts**: Three comprehensive visualizations
            • **Export Options**: CSV, Excel and a PowerPoint in the Operation Summary template style
            """)
        
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; padding: 2rem; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 15px; color: white;">
        <h2>🚀 Ready to Analyze Your Production Data?</h2>
        <p>Upload your Excel file above to unlock powerful insights and generate professional reports!</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title="Oil & Gas Analytics Dashboard", 
        page_icon="🛢️", 
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Enhanced Custom CSS for better user experience
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .info-box {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin: 1rem 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        color: white;
        margin: 0.5rem;
    }
    .upload-section {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        padding: 0.5rem 1rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .sidebar-developer {
        text-align: center;
        margin-bottom: 1rem;
        padding: 1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        color: white;
    }
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 5px 5px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1f77b4;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Main title
    st.markdown('<h1 class="main-header">🛢️ Oil & Gas Analytics Dashboard</h1>', unsafe_allow_html=True)
    
    render_sidebar()

    production_file, drilling_files = render_upload_area()

    # Drilling summaries must be ready before the PowerPoint is generated
    collect_drilling_summaries(drilling_files)

    if not production_file and not drilling_files:
        render_intro()

    production_analysis_section(production_file, show_guide=not drilling_files)
    drilling_reports_section(drilling_files)

if __name__ == "__main__":
    main()

