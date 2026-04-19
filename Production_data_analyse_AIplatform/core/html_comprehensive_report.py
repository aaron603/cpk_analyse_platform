"""
html_comprehensive_report.py
-----------------------------
Generates a comprehensive multi-tab HTML analysis report using Chart.js.

Tabs:
  1. Overview       – KPI cards, alert banner, yield trend, fail type doughnut, sheet summary
  2. Fail Analysis  – Top 25 fail bar chart (click→detail), fail detail table with SN search
  3. CPK Analysis   – Cpk bar chart (color-coded) + full CPK table with search
  4. Distribution   – Button selector → histogram (pass/fail stacked) + stats panel
  5. Fail Patterns  – Summary cards, hourly heatmap, multi-fail SN list
  6. Fault Replay   – SN list panel + expandable per-sheet detail
"""

from __future__ import annotations

import json
import math
import os
import random
import sys
from collections import defaultdict
from datetime import datetime
from typing import Optional

# ---------------------------------------------------------------------------
# Bundled Chart.js — loaded inline so reports work without internet access.
# Falls back to CDN if the local asset file is missing.
# ---------------------------------------------------------------------------
def _chartjs_asset_path() -> str:
    """Resolve chart.umd.min.js path for both dev and PyInstaller frozen modes."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS  # type: ignore[attr-defined]
        return os.path.join(base, 'core', 'assets', 'chart.umd.min.js')
    return os.path.join(os.path.dirname(__file__), 'assets', 'chart.umd.min.js')

_CHARTJS_ASSET = _chartjs_asset_path()


def _chartjs_script_tag() -> str:
    """Return a <script> tag: inline JS if asset exists, CDN fallback otherwise."""
    try:
        with open(_CHARTJS_ASSET, 'r', encoding='utf-8') as _f:
            _js = _f.read()
        return f'<script>/* Chart.js 4.4.0 bundled */\n{_js}\n</script>'
    except OSError:
        return (
            '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"'
            ' async onload="if(window._onChartJsLoaded)window._onChartJsLoaded()"></script>'
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_comprehensive_report(
    analysis_data: dict,
    output_path: str,
    title: str = '',
    generated_at: str = '',
    fail_data: dict = None,
    log_cb=None,
) -> str:
    """
    Build and write an HTML report.

    Parameters
    ----------
    analysis_data : dict
        {stype: {sheet_name: {point_name: {n, mean, std, lsl, usl, cp, cpk,
                                            values: [(barcode, val, is_pass)]}}}}
    output_path : str
        Destination file path for the HTML.
    title : str
        Report title shown in the topbar.
    generated_at : str
        Generation timestamp string, e.g. "2026-04-16 14:30:00".
    fail_data : dict, optional
        {stype: {barcode_stats, fail_barcodes, never_pass_barcodes, all_fail_items}}
    log_cb : callable, optional
        Logging callback for diagnostic messages.

    Returns
    -------
    str
        output_path
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    if not generated_at:
        generated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if not title:
        title = 'CPK Comprehensive Analysis Report'

    # ── Diagnostic: inspect analysis_data before building ──────────────────
    _log('[综合报告] ===== 数据诊断 =====')
    if not analysis_data:
        _log('[综合报告] ⚠ analysis_data 为空 → 报告将无测试数据（Pass 0 / Fail 0）')
    else:
        total_pts = 0
        total_vals = 0
        for stype, sheets in analysis_data.items():
            n_pts = sum(len(pts) for pts in sheets.values())
            n_vals = sum(
                sum(len(st.get('values') or []) for st in pts.values())
                for pts in sheets.values()
            )
            total_pts += n_pts
            total_vals += n_vals
            _log(f'[综合报告]   [{stype}]: {len(sheets)} Sheet, '
                 f'{n_pts} 测试项, {n_vals} 测量值')
        if total_vals == 0:
            _log('[综合报告] ⚠ 测试项有CPK但 values 列表全为空 → SN数将为0')
            _log('[综合报告]   原因：analyze_xlsx/json_folder 未保存 values 字段')
        else:
            _log(f'[综合报告] ✓ 共 {total_pts} 测试项, {total_vals} 测量值 → 将写入报告')

    # ── Build per-station data ──────────────────────────────────────────────
    stypes = list(analysis_data.keys()) if analysis_data else []
    per_station: dict = {}  # {stype: (data_dict, sn_detail)}

    if len(stypes) > 1:
        # "All" merged view first, then individual station views
        dd_all, sn_all = _build_data(analysis_data, fail_data, title, generated_at)
        per_station['All'] = (dd_all, sn_all)
        for stype in stypes:
            stype_analysis = {stype: analysis_data[stype]}
            stype_fail = (
                {stype: fail_data[stype]} if fail_data and stype in fail_data else None
            )
            dd, sn = _build_data(stype_analysis, stype_fail, title, generated_at)
            per_station[stype] = (dd, sn)
        data_dict = dd_all
    elif stypes:
        dd, sn = _build_data(analysis_data, fail_data, title, generated_at)
        per_station[stypes[0]] = (dd, sn)
        data_dict = dd
    else:
        dd, sn = _build_data({}, None, title, generated_at)
        per_station['All'] = (dd, sn)
        data_dict = dd

    # ── Diagnostic: inspect built data_dict ────────────────────────────────
    _log(f'[综合报告] data_dict: total={data_dict["total"]}, '
         f'pass={data_dict["pass_n"]}, fail={data_dict["fail_n"]}, '
         f'yield={data_dict["yield_pct"]}%, '
         f'cpk_list={len(data_dict["cpk_list"])}, '
         f'sheet_summary={len(data_dict["sheet_summary"])}')
    _log('[综合报告] =================')

    html_str = _build_html(per_station, title, generated_at)

    with open(output_path, 'w', encoding='utf-8') as fh:
        fh.write(html_str)

    _log(f'[综合报告] 文件写入完成: {len(html_str):,} 字符, '
         f'{os.path.getsize(output_path) // 1024} KB')

    return output_path


# ---------------------------------------------------------------------------
# Internal: build data structures
# ---------------------------------------------------------------------------

def _build_data(
    analysis_data: dict,
    fail_data: Optional[dict],
    title: str,
    generated_at: str,
) -> tuple:
    """
    Returns (data_dict, sn_detail).

    data_dict  – serialised as JS constant DATA
    sn_detail  – serialised as JS constant SN_DETAIL (pruned to FAIL SNs + up to 200 PASS)
    """

    # ---- Step 1: flatten all measurements --------------------------------
    # bc_measurements: {barcode: {(sheet, point): (val, is_pass, lsl, usl, unit)}}
    # point_meta:      {(sheet, point): {n, mean, std, lsl, usl, cp, cpk, unit}}
    # point_values:    {(sheet, point): [(barcode, val, is_pass)]}

    bc_measurements: dict[str, dict] = defaultdict(dict)
    point_meta: dict[tuple, dict] = {}
    point_values: dict[tuple, list] = defaultdict(list)

    for stype, sheets in analysis_data.items():
        for sheet_name, points in sheets.items():
            for point_name, stats in points.items():
                key = (sheet_name, point_name)
                unit = stats.get('unit', '')
                point_meta[key] = {
                    'n': stats.get('n', 0),
                    'mean': stats.get('mean'),
                    'std': stats.get('std'),
                    'lsl': stats.get('lsl'),
                    'usl': stats.get('usl'),
                    'cp': stats.get('cp'),
                    'cpk': stats.get('cpk'),
                    'unit': unit,
                    'sheet': sheet_name,
                }
                for bc, val, is_pass in (stats.get('values') or []):
                    bc_measurements[bc][(sheet_name, point_name)] = (
                        val, is_pass, stats.get('lsl'), stats.get('usl'), unit
                    )
                    point_values[key].append((bc, val, is_pass))

    # ---- Step 2: barcode-level pass/fail ---------------------------------
    all_barcodes = set(bc_measurements.keys())

    # Also collect barcodes mentioned only in fail_data
    if fail_data:
        for stype, fd in fail_data.items():
            for bc in (fd.get('fail_barcodes') or []):
                all_barcodes.add(bc)
            for bc in (fd.get('never_pass_barcodes') or []):
                all_barcodes.add(bc)
            for bc in (fd.get('barcode_stats') or {}):
                all_barcodes.add(bc)

    fail_bcs: set[str] = set()
    for bc, meas in bc_measurements.items():
        if any(not is_p for (_, is_p, _, _, _) in meas.values()):
            fail_bcs.add(bc)

    pass_bcs = all_barcodes - fail_bcs
    total = len(all_barcodes)
    pass_n = len(pass_bcs)
    fail_n = len(fail_bcs)
    yield_pct = round(pass_n / total * 100, 2) if total else 0.0

    # ---- Step 3: never_pass & multi_fail from fail_data ------------------
    never_pass_list: list[str] = []
    multi_fail_list: list[dict] = []
    retry_pass_count = 0

    fail_data_bcs: dict[str, dict] = {}  # barcode → barcode_stats entry

    if fail_data:
        for stype, fd in fail_data.items():
            np_bcs = set(fd.get('never_pass_barcodes') or [])
            never_pass_list = sorted(np_bcs)
            bs = fd.get('barcode_stats') or {}
            fail_data_bcs.update(bs)

            # retry-pass: in fail_barcodes but has pass_count > 0
            for bc in (fd.get('fail_barcodes') or []):
                entry = bs.get(bc, {})
                if entry.get('pass_count', 0) > 0:
                    retry_pass_count += 1
    else:
        # fail_data not provided — derive never_pass from measurement values:
        # a barcode is "never passed" if ALL its measurements failed
        bc_has_any_pass = {
            bc for bc, meas in bc_measurements.items()
            if any(is_p for (_, is_p, _, _, _) in meas.values())
        }
        never_pass_list = sorted(fail_bcs - bc_has_any_pass)

    # multi_fail: barcodes with ≥2 distinct failing test items
    bc_fail_items: dict[str, list] = defaultdict(list)
    for (sheet, point), pvlist in point_values.items():
        for bc, val, is_pass in pvlist:
            if not is_pass:
                bc_fail_items[bc].append(point)

    for bc, items in bc_fail_items.items():
        if len(items) >= 2:
            multi_fail_list.append({'sn': bc, 'fc': len(items)})
    multi_fail_list.sort(key=lambda x: -x['fc'])

    # ---- Step 4: time range & hourly data --------------------------------
    hourly_data: list[dict] = []
    all_times: list[str] = []

    if fail_data_bcs:
        hour_pass: dict[str, int] = defaultdict(int)
        hour_fail: dict[str, int] = defaultdict(int)

        fail_set_all: set[str] = set()
        if fail_data:
            for stype, fd in fail_data.items():
                fail_set_all.update(fd.get('fail_barcodes') or [])

        for bc, entry in fail_data_bcs.items():
            times = entry.get('times') or []
            for t in times:
                if t:
                    all_times.append(t)
            # Use latest time for the hour bucket
            if times:
                t_latest = max(times)
                hk = t_latest[:13]  # "YYYY-MM-DD HH"
                if bc in fail_set_all:
                    hour_fail[hk] += 1
                else:
                    hour_pass[hk] += 1

        all_hours = sorted(set(hour_pass) | set(hour_fail))
        for hk in all_hours:
            p = hour_pass[hk]
            f = hour_fail[hk]
            t = p + f
            hourly_data.append({
                'hour': hk,
                'pass': p,
                'fail': f,
                'total': t,
                'yield': round(p / t * 100, 1) if t else 0.0,
            })

    time_range = '—'
    if all_times:
        t_min = min(all_times)[:16]
        t_max = max(all_times)[:16]
        time_range = f'{t_min} ~ {t_max}' if t_min != t_max else t_min

    # ---- Step 5: sheet_summary ------------------------------------------
    sheet_bc_pass: dict[str, set] = defaultdict(set)
    sheet_bc_fail: dict[str, set] = defaultdict(set)
    sheet_bc_all: dict[str, set] = defaultdict(set)

    for (sheet, point), pvlist in point_values.items():
        for bc, val, is_pass in pvlist:
            sheet_bc_all[sheet].add(bc)
            if is_pass:
                sheet_bc_pass[sheet].add(bc)
            else:
                sheet_bc_fail[sheet].add(bc)

    sheet_summary = []
    for sheet in sheet_bc_all:
        all_s = sheet_bc_all[sheet]
        fail_s = sheet_bc_fail[sheet]
        pass_s = all_s - fail_s
        total_s = len(all_s)
        fail_s_n = len(fail_s)
        sheet_summary.append({
            'sheet': sheet,
            'total': total_s,
            'pass': len(pass_s),
            'fail': fail_s_n,
            'fail_rate': round(fail_s_n / total_s * 100, 1) if total_s else 0.0,
        })
    sheet_summary.sort(key=lambda x: -x['fail'])

    # ---- Step 6: fail_point_stats (top 25) ------------------------------
    point_fail_data: dict[tuple, list] = defaultdict(list)  # key → fail_recs
    point_fail_total: dict[tuple, int] = {}

    # We need time per barcode from fail_data
    bc_time: dict[str, str] = {}
    for bc, entry in fail_data_bcs.items():
        times = entry.get('times') or []
        if times:
            bc_time[bc] = max(times)

    for (sheet, point), pvlist in point_values.items():
        total_p = len(pvlist)
        point_fail_total[(sheet, point)] = total_p
        meta = point_meta.get((sheet, point), {})
        for bc, val, is_pass in pvlist:
            if not is_pass:
                point_fail_data[(sheet, point)].append({
                    'sn': bc,
                    'data': val,
                    'lsl': meta.get('lsl'),
                    'usl': meta.get('usl'),
                    'time': bc_time.get(bc, ''),
                    'ch': '',
                })

    fail_point_stats = []
    for (sheet, point), fail_recs in point_fail_data.items():
        if not fail_recs:
            continue
        total_p = point_fail_total.get((sheet, point), len(fail_recs))
        meta = point_meta.get((sheet, point), {})
        fail_point_stats.append({
            'name': point,
            'fail': len(fail_recs),
            'total': total_p,
            'fail_rate': round(len(fail_recs) / total_p * 100, 1) if total_p else 0.0,
            'lsl': meta.get('lsl'),
            'usl': meta.get('usl'),
            'unit': meta.get('unit', ''),
            'fail_recs': fail_recs,
        })

    fail_point_stats.sort(key=lambda x: -x['fail'])
    fail_point_stats = fail_point_stats[:25]

    # ---- Step 7: cpk_list -----------------------------------------------
    cpk_list = []
    for (sheet, point), meta in point_meta.items():
        cpk_val = meta.get('cpk')
        if cpk_val is None:
            continue
        cpk_list.append({
            'name': point,
            'unit': meta.get('unit', ''),
            'n': meta.get('n', 0),
            'mean': meta.get('mean') or 0.0,
            'std': meta.get('std') or 0.0,
            'lsl': meta.get('lsl'),
            'usl': meta.get('usl'),
            'cp': meta.get('cp') or 0.0,
            'cpk': cpk_val,
        })
    cpk_list.sort(key=lambda x: x['cpk'])

    # ---- Step 8: dist_data (max 500 values per point) -------------------
    dist_data: dict[str, dict] = {}
    MAX_DIST = 500

    for (sheet, point), pvlist in point_values.items():
        meta = point_meta.get((sheet, point), {})
        all_vals = [(v, ip) for (_, v, ip) in pvlist if v is not None]
        fail_vals = [v for (v, ip) in all_vals if not ip]
        pass_vals = [v for (v, ip) in all_vals if ip]

        # Sample if needed: keep all fail_vals, sample pass_vals
        if len(all_vals) > MAX_DIST:
            budget = MAX_DIST - len(fail_vals)
            if budget > 0 and len(pass_vals) > budget:
                pass_vals = random.sample(pass_vals, budget)
            elif budget <= 0:
                pass_vals = []
        combined = pass_vals + fail_vals

        if not combined:
            continue

        dist_data[point] = {
            'vals': combined,
            'fail_vals': fail_vals,
            'lsl': meta.get('lsl'),
            'usl': meta.get('usl'),
            'unit': meta.get('unit', ''),
        }

    # ---- Step 9: fault_type_list ----------------------------------------
    fault_type_list: list[dict] = []
    if fail_data:
        retry_pass_n = 0
        persist_fail_n = 0
        fail_set_all2: set[str] = set()
        for stype, fd in fail_data.items():
            fail_set_all2.update(fd.get('fail_barcodes') or [])
            bs = fd.get('barcode_stats') or {}
            for bc in fail_set_all2:
                entry = bs.get(bc, {})
                if entry.get('pass_count', 0) > 0:
                    retry_pass_n += 1
                else:
                    persist_fail_n += 1
        if retry_pass_n:
            fault_type_list.append({'type': 'Retry-Pass', 'count': retry_pass_n})
        if persist_fail_n:
            fault_type_list.append({'type': 'Persistent Fail', 'count': persist_fail_n})
    elif fail_bcs:
        # fail_data not provided — heuristic classification from measurement values:
        # barcodes with mixed pass/fail measurements → likely had a retry-pass run;
        # barcodes with only failing measurements → persistent fail.
        bc_has_any_pass_set = {
            bc for bc, meas in bc_measurements.items()
            if any(is_p for (_, is_p, _, _, _) in meas.values())
        }
        retry_n2 = len(fail_bcs & bc_has_any_pass_set)
        persist_n2 = len(fail_bcs - bc_has_any_pass_set)
        if retry_n2:
            fault_type_list.append({'type': 'Retry-Pass', 'count': retry_n2})
        if persist_n2:
            fault_type_list.append({'type': 'Persistent Fail', 'count': persist_n2})

    # ---- Step 10: sn_detail (FAIL SNs + up to 200 PASS) ----------------
    full_sn_detail: dict[str, dict] = {}

    for bc in all_barcodes:
        meas = bc_measurements.get(bc, {})
        is_fail = bc in fail_bcs
        overall = 'FAIL' if is_fail else 'PASS'
        time_str = bc_time.get(bc, '')

        fail_items = []
        sheets_detail: dict[str, list] = defaultdict(list)

        for (sheet, point), (val, is_pass, lsl, usl, unit) in meas.items():
            row = {
                'point': point,
                'data': val,
                'lsl': lsl,
                'usl': usl,
                'result': 'PASS' if is_pass else 'FAIL',
                'unit': unit,
            }
            sheets_detail[sheet].append(row)
            if not is_pass:
                fail_items.append({
                    'sheet': sheet,
                    'point': point,
                    'data': val,
                    'lsl': lsl,
                    'usl': usl,
                    'unit': unit,
                })

        full_sn_detail[bc] = {
            'overall': overall,
            'time': time_str,
            'fail_items': fail_items,
            'sheets': {s: rows for s, rows in sheets_detail.items()},
        }

    # Prune: keep all FAIL barcodes, up to 200 random PASS barcodes
    fail_detail = {bc: v for bc, v in full_sn_detail.items() if v['overall'] == 'FAIL'}
    pass_detail = {bc: v for bc, v in full_sn_detail.items() if v['overall'] == 'PASS'}
    if len(pass_detail) > 200:
        sampled_pass_keys = random.sample(list(pass_detail.keys()), 200)
        pass_detail = {k: pass_detail[k] for k in sampled_pass_keys}
    sn_detail = {**fail_detail, **pass_detail}

    # ---- Step 11: assemble data_dict ------------------------------------
    data_dict = {
        'title': title,
        'generated_at': generated_at,
        'total': total,
        'pass_n': pass_n,
        'fail_n': fail_n,
        'yield_pct': yield_pct,
        'time_range': time_range,
        'fail_sn_list': sorted(fail_bcs),
        'hourly': hourly_data,
        'sheet_summary': sheet_summary,
        'fail_point_stats': fail_point_stats,
        'cpk_list': cpk_list,
        'dist_data': dist_data,
        'fault_type_list': fault_type_list,
        'never_pass': never_pass_list,
        'multi_fail': multi_fail_list,
    }

    return data_dict, sn_detail


# ---------------------------------------------------------------------------
# Internal: build HTML
# ---------------------------------------------------------------------------

_CSS = r"""
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI','Microsoft YaHei',sans-serif;background:#f0f2f5;color:#1a202c;font-size:14px;line-height:1.5}
.topbar{position:fixed;top:0;left:0;right:0;z-index:999;background:linear-gradient(135deg,#1e3a8a,#1a56db);color:#fff;padding:0 24px;height:56px;display:flex;align-items:center;gap:16px;box-shadow:0 2px 8px rgba(0,0,0,.25)}
.topbar-title{font-size:18px;font-weight:700}
.topbar-sub{font-size:12px;opacity:.8;margin-left:4px}
.topbar-badges{display:flex;gap:8px;margin-left:auto}
.tbadge{padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600}
.tbadge-pass{background:rgba(5,122,85,.85);color:#d1fae5}
.tbadge-fail{background:rgba(200,30,30,.85);color:#fee2e2}
.tabnav{position:sticky;top:56px;z-index:990;background:#fff;border-bottom:2px solid #e5e7eb;display:flex;overflow-x:auto;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.tabnav::-webkit-scrollbar{height:0}
.tabBtn{padding:13px 20px;border:none;background:none;cursor:pointer;font-size:13px;font-weight:600;color:#6b7280;white-space:nowrap;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s}
.tabBtn:hover{color:#1a56db;background:#f0f4ff}
.tabBtn.active{color:#1a56db;border-bottom-color:#1a56db}
.main{max-width:1600px;margin:0 auto;padding:76px 20px 40px}
.page{display:none}.page.active{display:block}
.card{background:#fff;border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:20px;margin-bottom:20px;border:1px solid #e5e7eb}
.card-title{font-size:15px;font-weight:700;color:#111827;margin-bottom:16px;display:flex;align-items:center;gap:8px}
.card-title::before{content:'';display:inline-block;width:4px;height:18px;background:#1a56db;border-radius:2px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:16px;margin-bottom:20px}
.kpi{background:#fff;border-radius:12px;padding:20px 16px;box-shadow:0 1px 3px rgba(0,0,0,.1);border:1px solid #e5e7eb;text-align:center;border-top:4px solid}
.kpi.blue{border-top-color:#1a56db}.kpi.green{border-top-color:#057a55}.kpi.red{border-top-color:#c81e1e}.kpi.orange{border-top-color:#b45309}.kpi.purple{border-top-color:#6c2bd9}
.kpi-val{font-size:36px;font-weight:800;line-height:1.1}
.kpi.blue .kpi-val{color:#1a56db}.kpi.green .kpi-val{color:#057a55}.kpi.red .kpi-val{color:#c81e1e}.kpi.orange .kpi-val{color:#b45309}.kpi.purple .kpi-val{color:#6c2bd9}
.kpi-lbl{font-size:12px;color:#6b7280;margin-top:4px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
@media(max-width:900px){.g2,.g3,.g4{grid-template-columns:1fr}}
.ch300{position:relative;height:300px}.ch360{position:relative;height:360px}.ch260{position:relative;height:260px}.ch220{position:relative;height:220px}
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#f9fafb;color:#374151;padding:9px 12px;text-align:left;font-weight:600;border-bottom:2px solid #e5e7eb;white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid #f3f4f6}
tr:hover td{background:#f0f4ff}
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}
.b-pass{background:#d1fae5;color:#065f46}.b-fail{background:#fee2e2;color:#991b1b}
.b-warn{background:#fef3c7;color:#92400e}.b-blue{background:#dbeafe;color:#1e40af}
.b-red{background:#fee2e2;color:#991b1b}.b-orange{background:#ffedd5;color:#9a3412}
.cpk-best{color:#065f46;font-weight:700}.cpk-good{color:#1e40af;font-weight:700}
.cpk-ok{color:#92400e;font-weight:700}.cpk-bad{color:#991b1b;font-weight:700}.cpk-worst{color:#7f1d1d;font-weight:800}
.alert{border-radius:10px;padding:14px 18px;margin-bottom:20px;display:flex;align-items:flex-start;gap:12px;font-size:13px}
.alert-red{background:#fef2f2;border:1px solid #fecaca}
.alert-icon{font-size:20px}.alert-body{flex:1}
.alert-title{font-weight:700;font-size:14px;margin-bottom:4px}
.heat-table{border-collapse:collapse;font-size:11px}
.heat-table th,.heat-table td{padding:4px 7px;border:1px solid #e5e7eb;white-space:nowrap;text-align:center}
.heat-table th{background:#f9fafb;font-weight:600;text-align:left}
.heat-table td{font-weight:600;min-width:40px}
.replay-wrap{display:grid;grid-template-columns:280px 1fr;gap:16px}
.sn-list{background:#f9fafb;border-radius:8px;border:1px solid #e5e7eb;overflow-y:auto;max-height:620px}
.sn-list::-webkit-scrollbar{width:4px}.sn-list::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:2px}
.sn-item{padding:10px 14px;cursor:pointer;border-bottom:1px solid #f3f4f6;transition:background .15s;display:flex;justify-content:space-between;align-items:center}
.sn-item:hover{background:#eef2ff}.sn-item.selected{background:#dbeafe;border-left:3px solid #1a56db}
.sn-code{font-family:monospace;font-size:12px;font-weight:600}
.sn-meta{font-size:11px;color:#6b7280;margin-top:2px}
.detail-pane{background:#fff;border-radius:8px;border:1px solid #e5e7eb;padding:16px;overflow-y:auto;max-height:620px}
.sheet-block{margin-bottom:10px}
.sheet-hdr{background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:8px 12px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-weight:600;font-size:13px;user-select:none}
.sheet-hdr.has-fail{background:#fef2f2;border-color:#fecaca;color:#991b1b}
.sheet-rows{display:none;margin-top:4px}.sheet-rows.open{display:block}
.test-row{display:grid;grid-template-columns:2fr 1fr 1fr 1fr 100px;gap:4px;padding:5px 10px;border-radius:4px;font-size:12px;align-items:center}
.row-fail{background:#fef2f2}.row-pass{background:#fafafa}
.dev-tag{font-size:11px;font-weight:700;color:#c81e1e;margin-left:4px}
.dist-btns{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px}
.dist-btn{padding:4px 10px;border:1px solid #d1d5db;background:#f9fafb;color:#374151;border-radius:20px;cursor:pointer;font-size:11px;transition:all .15s;white-space:nowrap}
.dist-btn:hover,.dist-btn.active{background:#1a56db;border-color:#1a56db;color:#fff}
.search-row{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.search-row input,.search-row select{background:#f9fafb;border:1px solid #d1d5db;color:#1a202c;padding:8px 12px;border-radius:8px;font-size:13px;outline:none;flex:1;min-width:200px}
.search-row input:focus{border-color:#1a56db}
.pbar-wrap{height:6px;background:#f3f4f6;border-radius:3px;overflow:hidden;margin-top:3px}
.pbar{height:100%;border-radius:3px}
.stype-bar{background:#f0f2f5;border-bottom:2px solid #dde3f0;padding:4px 16px;display:flex;gap:6px;overflow-x:auto;position:sticky;top:56px;z-index:991;flex-shrink:0}
.stype-btn{padding:5px 14px;border:1px solid #c5cae9;background:white;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;color:#374151;transition:all .15s}
.stype-btn.active{background:#3949ab;color:white;border-color:#3949ab}
"""


def _build_html(per_station: dict, title: str, generated_at: str) -> str:
    # per_station: {stype: (data_dict, sn_detail)}
    station_list = list(per_station.keys())
    multi = len(station_list) > 1
    default_stype = station_list[0] if station_list else 'All'

    # Serialize per-station data for JS — dist_data lazy-loaded separately
    station_js_entries = []
    json_script_tags = []
    for stype in station_list:
        data_dict_s, sn_detail_s = per_station[stype]
        data_for_js = dict(data_dict_s)
        dist_data_s = data_for_js.pop('dist_data', {})
        safe_id = stype.replace(' ', '_').replace('-', '_')
        sn_id   = f'_snJson_{safe_id}'
        dist_id = f'_distJson_{safe_id}'
        data_json_s    = json.dumps(_sanitize_for_json(data_for_js),   ensure_ascii=False, default=_json_default)
        sn_json_safe   = json.dumps(_sanitize_for_json(sn_detail_s),   ensure_ascii=False, default=_json_default).replace('</', '<\\/')
        dist_json_safe = json.dumps(_sanitize_for_json(dist_data_s),   ensure_ascii=False, default=_json_default).replace('</', '<\\/')
        station_js_entries.append(
            f'"{_esc_js(stype)}":{{data:{data_json_s},sn_id:"{sn_id}",dist_id:"{dist_id}"}}'
        )
        json_script_tags.append(
            f'<script type="application/json" id="{sn_id}">{sn_json_safe}</script>\n'
            f'<script type="application/json" id="{dist_id}">{dist_json_safe}</script>'
        )

    station_data_js = '{' + ',\n'.join(station_js_entries) + '}'
    json_scripts_html = '\n'.join(json_script_tags)

    # Python-side static banner values from the default station
    data_dict_0 = per_station[default_stype][0] if per_station else {}
    py_total  = data_dict_0.get('total', 0)
    py_pass   = data_dict_0.get('pass_n', 0)
    py_fail   = data_dict_0.get('fail_n', 0)
    py_yield  = data_dict_0.get('yield_pct', 0.0)
    py_cpk    = len(data_dict_0.get('cpk_list', []))
    py_sheets = len(data_dict_0.get('sheet_summary', []))

    # Conditional station tab bar HTML
    station_btns_html = '\n  '.join(
        f'<button class="stype-btn{" active" if i == 0 else ""}" data-stype="{_esc(stype)}" '
        f'onclick="switchStationType(\'{_esc_js(stype)}\')">{_esc(stype)}</button>'
        for i, stype in enumerate(station_list)
    )
    stype_bar_html = (
        f'<!-- STATION TYPE TAB BAR -->\n'
        f'<div class="stype-bar" id="stypeBar">\n  {station_btns_html}\n</div>\n'
    ) if multi else ''
    tabnav_override = (
        '<style>.tabnav{top:96px}.main{padding-top:116px}</style>\n'
    ) if multi else ''

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{_esc(title)}</title>
<style>{_CSS}</style>
{tabnav_override}</head>
<body>

<!-- Python static data summary — fixed at bottom, written directly by Python, no JS/CDN dependency -->
<div id="pyDataBanner" style="position:fixed;bottom:0;left:0;right:0;z-index:1000;background:#e8f5e9;border-top:2px solid #66bb6a;padding:4px 14px;font-size:11.5px;color:#1b5e20;display:flex;gap:18px;align-items:center;flex-wrap:wrap;font-family:monospace;box-shadow:0 -2px 6px rgba(0,0,0,.1)">
  <span style="font-weight:700">📊 Python Data Summary (No JS)</span>
  <span>Total SN: <b>{py_total}</b></span>
  <span>Pass: <b>{py_pass}</b></span>
  <span>Fail: <b>{py_fail}</b></span>
  <span>Yield: <b>{py_yield}%</b></span>
  <span>CPK Items: <b>{py_cpk}</b></span>
  <span>Sheets: <b>{py_sheets}</b></span>
  <span style="margin-left:auto;cursor:pointer;padding:2px 8px;background:#c8e6c9;border-radius:4px" onclick="this.parentNode.style.display='none'">Close ✕</span>
</div>

<!-- TOP BAR -->
<div class="topbar">
  <span class="topbar-title" id="topTitle">{_esc(title)}</span>
  <span class="topbar-sub" id="hdrTime"></span>
  <div class="topbar-badges">
    <span class="tbadge tbadge-pass" id="hdrPass">Pass —</span>
    <span class="tbadge tbadge-fail" id="hdrFail">Fail —</span>
  </div>
</div>

{stype_bar_html}<!-- TAB NAV -->
<nav class="tabnav">
  <button class="tabBtn active" onclick="showTab('overview',this)">Overview</button>
  <button class="tabBtn" onclick="showTab('fail',this)">Fail Analysis</button>
  <button class="tabBtn" onclick="showTab('cpk',this)">CPK Analysis</button>
  <button class="tabBtn" onclick="showTab('dist',this)">Distribution</button>
  <button class="tabBtn" onclick="showTab('pattern',this)">Fail Patterns</button>
  <button class="tabBtn" onclick="showTab('replay',this)">Fault Replay</button>
</nav>

<div class="main">

<!-- ===== PAGE: OVERVIEW ===== -->
<div id="page-overview" class="page active">
  <div id="kpiRow" class="kpi-row"></div>
  <div id="alertBanner"></div>
  <div class="g2">
    <div class="card">
      <div class="card-title">Yield Trend</div>
      <div id="yieldChartWrap" class="ch300"><canvas id="yieldChart"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title">Failure Type Distribution</div>
      <div id="failTypeWrap" class="ch300"><canvas id="failTypeChart"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-title">Test Sheet Summary</div>
    <div class="tbl-wrap">
      <table id="sheetTbl">
        <thead><tr><th>Sheet</th><th>Total</th><th>Pass</th><th>Fail</th><th>Fail Rate</th><th>Status</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ===== PAGE: FAIL ANALYSIS ===== -->
<div id="page-fail" class="page">
  <div id="failCards" class="kpi-row"></div>
  <div class="card">
    <div class="card-title">Top 25 Failed Test Items</div>
    <div id="failBarWrap" style="position:relative;height:480px"><canvas id="failBar"></canvas></div>
  </div>
  <div class="card" id="failDetailCard" style="display:none">
    <div class="card-title" id="failDetailTitle">Failure Detail</div>
    <div class="search-row">
      <input id="failSearch" placeholder="Search barcode..." oninput="filterFail()">
      <span id="failCnt" style="color:#6b7280;font-size:13px;white-space:nowrap"></span>
    </div>
    <div class="tbl-wrap">
      <table id="failDetailTbl">
        <thead><tr><th>Barcode</th><th>Test Time</th><th>Channel</th><th>Value</th><th>LSL</th><th>USL</th><th>Deviation</th><th>Direction</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ===== PAGE: CPK ===== -->
<div id="page-cpk" class="page">
  <div class="card">
    <div class="card-title">CPK Chart (Items with Cpk &lt; 1.67)</div>
    <div id="cpkWrap" style="position:relative;height:400px"></div>
  </div>
  <div class="card">
    <div class="card-title">CPK Detail Data</div>
    <div class="search-row">
      <input id="cpkSearch" placeholder="Search test item..." oninput="filterCpk()">
    </div>
    <div class="tbl-wrap">
      <table id="cpkTbl">
        <thead><tr><th>Test Item</th><th>Unit</th><th>n</th><th>Mean</th><th>σ</th><th>LSL</th><th>USL</th><th>Cp</th><th>Cpk</th><th>Grade</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ===== PAGE: DISTRIBUTION ===== -->
<div id="page-dist" class="page">
  <div class="card">
    <div class="card-title">Select Test Item</div>
    <div id="distBtns" class="dist-btns"></div>
  </div>
  <div class="card" id="distCard" style="display:none">
    <div class="card-title" id="distTitle">Data Distribution</div>
    <div class="g2">
      <div class="ch300"><canvas id="distHist"></canvas></div>
      <div id="distStats" style="display:grid;grid-template-columns:1fr 1fr;gap:10px;align-content:start"></div>
    </div>
  </div>
</div>

<!-- ===== PAGE: PATTERN (Fail Patterns) ===== -->
<div id="page-pattern" class="page">
  <div id="patternCards" class="kpi-row" style="grid-template-columns:repeat(4,1fr)"></div>
  <div class="card" id="heatCard" style="display:none">
    <div class="card-title">Failure Hour Heatmap</div>
    <div class="tbl-wrap" id="heatDiv"></div>
  </div>
  <div class="card">
    <div class="card-title">Multi-Failure SN List (≥2 failed items)</div>
    <div class="tbl-wrap">
      <table id="multiFail Tbl">
        <thead><tr><th>Barcode</th><th>Failed Items</th><th>Main Failed Items</th></tr></thead>
        <tbody id="multiFailBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ===== PAGE: REPLAY ===== -->
<div id="page-replay" class="page">
  <div class="replay-wrap">
    <div>
      <div class="search-row" style="margin-bottom:8px">
        <input id="snSearch" placeholder="Search barcode..." oninput="filterSN()" style="flex:1;min-width:0">
        <select id="snFilt" onchange="filterSN()" style="flex:0 0 90px;min-width:0">
          <option value="">All</option>
          <option value="FAIL">FAIL</option>
          <option value="PASS">PASS</option>
        </select>
      </div>
      <div style="font-size:12px;color:#6b7280;margin-bottom:6px">Total: <span id="snCnt">0</span></div>
      <div class="sn-list" id="snList"></div>
    </div>
    <div class="detail-pane" id="detPane">
      <div style="padding:40px;text-align:center;color:#9ca3af">← Click an SN on the left to view details</div>
    </div>
  </div>
</div>

</div><!-- .main -->

{_chartjs_script_tag()}
<script>
var STATION_DATA = {station_data_js};
var _curStype = "{_esc_js(default_stype)}";
var DATA = STATION_DATA[_curStype].data;
var SN_DETAIL = null;   // lazy-loaded per station on first Fault Replay tab open
var DIST_DATA = null;   // lazy-loaded per station on first Distribution tab open

// ---------------------------------------------------------------------------
// Station switching
// ---------------------------------------------------------------------------
function switchStationType(stype) {{
  if (stype === _curStype) return;
  _curStype = stype;
  DATA = STATION_DATA[stype].data;
  SN_DETAIL = null;
  DIST_DATA = null;
  // Destroy known charts so they can be re-created cleanly
  ['yieldChart','failTypeChart','failBar','distHist'].forEach(function(cid) {{
    var el = document.getElementById(cid);
    if (el && typeof Chart !== 'undefined') {{ var c = Chart.getChart(el); if(c) c.destroy(); }}
  }});
  var cpkWrap = document.getElementById('cpkWrap');
  if (cpkWrap && typeof Chart !== 'undefined') {{
    cpkWrap.querySelectorAll('canvas').forEach(function(cv) {{
      var c = Chart.getChart(cv); if(c) c.destroy();
    }});
    cpkWrap.innerHTML = '';
  }}
  // Reset all tab init flags
  _initDone = {{}};
  _overviewChartsDone = false;
  _overviewD = null;
  // Update station tab active state
  document.querySelectorAll('.stype-btn').forEach(function(b) {{
    b.classList.toggle('active', b.dataset.stype === stype);
  }});
  // Re-run overview and re-init the currently active content tab
  runOverview();
  var activeBtn = document.querySelector('.tabBtn.active');
  if (activeBtn) {{
    var m = activeBtn.getAttribute('onclick').match(/showTab\\('(\\w+)'/);
    if (m && m[1] !== 'overview') {{ _initDone[m[1]] = false; initTab(m[1]); }}
  }}
}}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function fmt(v,d){{return v===null||v===undefined?'N/A':(+v).toFixed(d||2);}}
function frColor(r){{return r>20?'#c81e1e':r>5?'#b45309':'#057a55';}}
function cpkCls(v){{return v>=1.67?'cpk-best':v>=1.33?'cpk-good':v>=1.0?'cpk-ok':v>=0.67?'cpk-bad':'cpk-worst';}}
function cpkBadge(v){{return v>=1.33?'<span class="badge b-pass">Excellent</span>':v>=1.0?'<span class="badge b-blue">Good</span>':v>=0.67?'<span class="badge b-warn">Caution</span>':'<span class="badge b-fail">Poor</span>';}}
// Chart.js loads async — always check dynamically so data renders immediately
// even when CDN is slow or blocked in factory environments
function tryChart(idOrEl, cfg) {{
  var el = typeof idOrEl === 'string' ? document.getElementById(idOrEl) : idOrEl;
  if (!el) return null;
  if (typeof Chart === 'undefined') {{
    // Overlay warning WITHOUT destroying the canvas element (needed for re-init after load)
    var wrap = el.parentNode;
    if (!wrap.querySelector('._chart-wait')) {{
      var msg = document.createElement('div');
      msg.className = '_chart-wait';
      msg.style.cssText = 'position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,.9);z-index:2;border-radius:8px';
      msg.innerHTML = '<span style="color:#9ca3af;font-size:13px">⚠ Chart library loading, please check network if not shown</span>';
      wrap.style.position = 'relative';
      wrap.appendChild(msg);
    }}
    return null;
  }}
  // Chart.js is ready — remove any pending overlay
  var pending = el.parentNode.querySelector('._chart-wait');
  if (pending) pending.remove();
  return new Chart(el, cfg);
}}

// ---------------------------------------------------------------------------
// Tab switching
// ---------------------------------------------------------------------------
var _initDone = {{}};
function showTab(id, btn) {{
  document.querySelectorAll('.page').forEach(function(p){{p.classList.remove('active');}});
  document.querySelectorAll('.tabBtn').forEach(function(b){{b.classList.remove('active');}});
  document.getElementById('page-'+id).classList.add('active');
  btn.classList.add('active');
  if(!_initDone[id]){{_initDone[id]=true;initTab(id);}}
}}
function initTab(id) {{
  if(id==='fail') initFail();
  else if(id==='pattern') initPattern();
  else if(id==='cpk') initCpk();
  else if(id==='dist') initDist();
  else if(id==='replay') initReplay();
}}

// ---------------------------------------------------------------------------
// OVERVIEW — data render runs immediately; charts run when Chart.js is ready
// ---------------------------------------------------------------------------
var _overviewD = null;           // shared so the CDN onload callback can reach it
var _overviewChartsDone = false; // prevent double-draw if CDN loads fast

function _drawOverviewCharts() {{
  var D = _overviewD;
  if (!D || typeof Chart === 'undefined') return;
  if (_overviewChartsDone) return;
  _overviewChartsDone = true;

  if(D.hourly && D.hourly.length > 0) {{
    var hrs  = D.hourly.map(function(h){{return h.hour.slice(5);}});
    var ylds = D.hourly.map(function(h){{return h.yield;}});
    var ptc  = ylds.map(function(y){{return y<85?'#c81e1e':y<95?'#b45309':'#057a55';}});
    tryChart('yieldChart', {{
      type: 'line',
      data: {{
        labels: hrs,
        datasets: [
          {{label:'Yield%',data:ylds,borderColor:'#1a56db',backgroundColor:'rgba(26,86,219,.08)',tension:.3,fill:true,pointBackgroundColor:ptc,pointRadius:5,yAxisID:'y'}},
          {{label:'Test Count',data:D.hourly.map(function(h){{return h.total;}}),borderColor:'#9ca3af',borderDash:[4,4],tension:.3,fill:false,pointRadius:3,yAxisID:'y2'}},
        ]
      }},
      options: {{
        responsive:true,maintainAspectRatio:false,
        interaction:{{mode:'index'}},
        plugins:{{legend:{{position:'top'}}}},
        scales:{{
          y:{{min:0,max:100,title:{{display:true,text:'Yield (%)'}}}},
          y2:{{position:'right',title:{{display:true,text:'Count'}},grid:{{drawOnChartArea:false}}}}
        }}
      }}
    }});
  }} else {{
    document.getElementById('yieldChartWrap').innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px">No time-series data</div>';
  }}

  if(D.fault_type_list && D.fault_type_list.length > 0) {{
    var ftC = ['#c81e1e','#d97706','#1a56db','#057a55','#6c2bd9','#9ca3af'];
    tryChart('failTypeChart', {{
      type: 'doughnut',
      data: {{
        labels: D.fault_type_list.map(function(t){{return t.type;}}),
        datasets: [{{
          data: D.fault_type_list.map(function(t){{return t.count;}}),
          backgroundColor: ftC,
          borderWidth: 2,
          borderColor: '#fff'
        }}]
      }},
      options: {{
        responsive:true,maintainAspectRatio:false,
        plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:14}}}}}}
      }}
    }});
  }} else {{
    document.getElementById('failTypeWrap').innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px">No data</div>';
  }}
}}

// CDN onload handler — re-draw charts in whichever tabs already initialized
window._onChartJsLoaded = function() {{
  _drawOverviewCharts();
  // Re-init tabs that were opened before Chart.js was ready
  if (_initDone['fail'])    {{ _initDone['fail']    = false; initFail(); }}
  if (_initDone['cpk'])     {{ _initDone['cpk']     = false; initCpk(); }}
  if (_initDone['dist'])    {{ _initDone['dist']    = false; initDist(); }}
  if (_initDone['pattern']) {{ _initDone['pattern'] = false; initPattern(); }}
}};

function runOverview() {{
  var D = DATA;
  _overviewD = D;  // store for CDN-onload re-draw
  _overviewChartsDone = false;  // reset so charts get re-drawn
  document.getElementById('topTitle').textContent = D.title || '{_esc_js(title)}';
  document.getElementById('hdrTime').textContent = D.time_range;
  document.getElementById('hdrPass').textContent = 'Pass ' + D.pass_n;
  document.getElementById('hdrFail').textContent = 'Fail ' + D.fail_n;

  var yc = D.yield_pct >= 95 ? 'green' : D.yield_pct >= 85 ? 'orange' : 'red';
  document.getElementById('kpiRow').innerHTML = [
    {{c:'blue',  v:D.total,           l:'Total SNs'}},
    {{c:'green', v:D.pass_n,          l:'Pass'}},
    {{c:'red',   v:D.fail_n,          l:'Fail'}},
    {{c:yc,      v:D.yield_pct+'%',   l:'Overall Yield'}},
    {{c:'purple',v:D.fail_sn_list.length, l:'Failed SNs'}},
  ].map(function(k){{
    return '<div class="kpi '+k.c+'"><div class="kpi-val">'+k.v+'</div><div class="kpi-lbl">'+k.l+'</div></div>';
  }}).join('');

  var top3 = D.fail_point_stats.slice(0,3);
  document.getElementById('alertBanner').innerHTML = '';
  if(top3.length) {{
    var items = top3.map(function(f){{return '<b>'+f.name+'</b> ('+f.fail+', '+f.fail_rate+'%)';}}).join(' &nbsp;|&nbsp; ');
    document.getElementById('alertBanner').innerHTML =
      '<div class="alert alert-red"><div class="alert-icon">⚠</div><div class="alert-body"><div class="alert-title">Top 3 Failing Test Items</div><div>'+items+'</div></div></div>';
  }}

  var tb = document.querySelector('#sheetTbl tbody');
  tb.innerHTML = '';  // clear before re-populating (needed on station switch)
  D.sheet_summary.forEach(function(s) {{
    var bg  = s.fail_rate > 20 ? '#fef2f2' : s.fail_rate > 5 ? '#fffbeb' : '';
    var bdr = s.fail_rate > 20 ? 'border-left:4px solid #c81e1e' : '';
    tb.innerHTML += '<tr style="background:'+bg+';'+bdr+'">'
      +'<td><b>'+s.sheet+'</b></td>'
      +'<td>'+s.total+'</td>'
      +'<td>'+s.pass+'</td>'
      +'<td>'+s.fail+'</td>'
      +'<td style="color:'+frColor(s.fail_rate)+';font-weight:700">'+s.fail_rate+'%'
        +'<div class="pbar-wrap"><div class="pbar" style="width:'+Math.min(s.fail_rate,100)+'%;background:'+frColor(s.fail_rate)+'"></div></div></td>'
      +'<td>'+(s.fail_rate>20?'<span class="badge b-red">Attention</span>':s.fail_rate>5?'<span class="badge b-warn">Caution</span>':'<span class="badge b-pass">Normal</span>')+'</td>'
      +'</tr>';
  }});

  // Draw overview charts immediately if Chart.js already loaded (fast CDN),
  // otherwise the window._onChartJsLoaded callback will draw them later
  _drawOverviewCharts();
}}
runOverview();  // initial load

// ---------------------------------------------------------------------------
// FAIL TAB
// ---------------------------------------------------------------------------
function initFail() {{
  var D = DATA;
  var uniqPts = D.fail_point_stats.length;
  var totalFails = D.fail_point_stats.reduce(function(a,f){{return a+f.fail;}},0);
  document.getElementById('failCards').innerHTML = [
    {{c:'red',   v:D.fail_sn_list.length, l:'Total Failed SNs'}},
    {{c:'orange',v:uniqPts,               l:'Failed Item Types'}},
    {{c:'blue',  v:totalFails,            l:'Total Failures'}},
  ].map(function(k){{
    return '<div class="kpi '+k.c+'"><div class="kpi-val">'+k.v+'</div><div class="kpi-lbl">'+k.l+'</div></div>';
  }}).join('');

  var top = D.fail_point_stats.slice(0,25);
  var h   = Math.max(480, top.length * 22);
  document.getElementById('failBarWrap').style.height = h + 'px';
  var bgs = top.map(function(f){{return f.fail_rate>30?'#c81e1e':f.fail_rate>15?'#d97706':'#f59e0b';}});
  tryChart('failBar', {{
    type: 'bar',
    data: {{
      labels: top.map(function(f){{return f.name.length>40?f.name.slice(0,40)+'...':f.name;}}),
      datasets: [{{label:'Fail Count',data:top.map(function(f){{return f.fail;}}),backgroundColor:bgs,borderRadius:3}}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{display:false}},
        tooltip: {{callbacks: {{label: function(ctx) {{
          var fp = top[ctx.dataIndex];
          return ['Fail: '+fp.fail+' | Rate: '+fp.fail_rate+'%','LSL: '+fp.lsl+'  USL: '+fp.usl+'  Unit: '+fp.unit];
        }}}}}}
      }},
      scales: {{
        x: {{title:{{display:true,text:'Fail Count'}}}},
        y: {{ticks:{{font:{{size:11}}}}}}
      }},
      onClick: function(e, el) {{
        if(el.length > 0) showFailDetail(top[el[0].index]);
      }}
    }}
  }});
}}

var _curFail = [];
function showFailDetail(fp) {{
  document.getElementById('failDetailCard').style.display = 'block';
  document.getElementById('failDetailTitle').textContent = 'Failure Detail: '+fp.name+' ('+fp.fail+' total)';
  _curFail = fp.fail_recs || [];
  renderFail(_curFail);
  document.getElementById('failDetailCard').scrollIntoView({{behavior:'smooth'}});
}}
function renderFail(recs) {{
  var tb = document.querySelector('#failDetailTbl tbody');
  tb.innerHTML = '';
  document.getElementById('failCnt').textContent = recs.length+' records';
  recs.forEach(function(r) {{
    var d=r.data, l=r.lsl, u=r.usl, dev='', dir='';
    if(d!==null && d!==undefined) {{
      if(u!==null && d>u)      {{dev='+'+(d-u).toFixed(3); dir='<span class="badge b-red">↑High</span>';}}
      else if(l!==null && d<l) {{dev=''+(d-l).toFixed(3);  dir='<span class="badge b-orange">↓Low</span>';}}
    }}
    tb.innerHTML += '<tr>'
      +'<td><code style="font-size:11px">'+r.sn+'</code></td>'
      +'<td style="font-size:11px">'+(r.time?r.time.slice(0,19):'')+'</td>'
      +'<td>'+(r.ch||'')+'</td>'
      +'<td style="font-weight:700;color:#c81e1e">'+fmt(d,3)+'</td>'
      +'<td>'+fmt(l,3)+'</td>'
      +'<td>'+fmt(u,3)+'</td>'
      +'<td style="color:#c81e1e;font-weight:700">'+dev+'</td>'
      +'<td>'+dir+'</td>'
      +'</tr>';
  }});
}}
function filterFail() {{
  var kw = document.getElementById('failSearch').value.toLowerCase();
  renderFail(kw ? _curFail.filter(function(r){{return r.sn.toLowerCase().indexOf(kw)>=0;}}) : _curFail);
}}

// ---------------------------------------------------------------------------
// PATTERN TAB (Fail Patterns)
// ---------------------------------------------------------------------------
function initPattern() {{
  var D = DATA;
  var failSNs   = D.fail_sn_list.length;
  var neverPass = (D.never_pass||[]).length;
  var multiF    = (D.multi_fail||[]).filter(function(m){{return m.fc>=2;}}).length;
  var retryPass = failSNs - neverPass; // approximation
  if(retryPass < 0) retryPass = 0;

  document.getElementById('patternCards').innerHTML = [
    {{c:'red',    v:failSNs,   l:'Total Failed SNs'}},
    {{c:'purple', v:neverPass, l:'Never-Passed SNs'}},
    {{c:'orange', v:multiF,    l:'Multi-Fail SNs (≥2)'}},
    {{c:'blue',   v:retryPass, l:'Retry-Pass SNs (est.)'}},
  ].map(function(k){{
    return '<div class="kpi '+k.c+'"><div class="kpi-val">'+k.v+'</div><div class="kpi-lbl">'+k.l+'</div></div>';
  }}).join('');

  if(D.hourly && D.hourly.length > 0) {{
    document.getElementById('heatCard').style.display = 'block';
    var byDate = {{}};
    D.hourly.forEach(function(h) {{
      var parts = h.hour.split(' ');
      var dt = parts[0], hr = parseInt(parts[1]);
      if(!byDate[dt]) byDate[dt] = {{}};
      byDate[dt][hr] = h;
    }});
    var dates = Object.keys(byDate).sort();
    var hrsArr = [];
    for(var i=0;i<24;i++) hrsArr.push(i);
    var htm = '<table class="heat-table"><thead><tr><th>Date/Hr</th>';
    hrsArr.forEach(function(h){{htm+='<th>'+h+'h</th>';}});
    htm += '</tr></thead><tbody>';
    dates.forEach(function(dt) {{
      htm += '<tr><th>'+dt.slice(5)+'</th>';
      hrsArr.forEach(function(h) {{
        var d = byDate[dt][h];
        if(!d){{htm+='<td style="background:#f9fafb;color:#d1d5db">-</td>';return;}}
        var fr = d.total>0 ? Math.round(d.fail/d.total*100) : 0;
        var bg = fr===0?'#d1fae5':fr<15?'#fef3c7':fr<30?'#fed7aa':fr<50?'#fca5a5':'#f87171';
        var fg = fr>=30?'#7f1d1d':'#1a202c';
        htm += '<td style="background:'+bg+';color:'+fg+'" title="'+dt+' '+h+'h | Total: '+d.total+' Fail: '+d.fail+' Yield: '+d.yield+'%">'+fr+'%</td>';
      }});
      htm += '</tr>';
    }});
    htm += '</tbody></table>';
    document.getElementById('heatDiv').innerHTML = htm;
  }}

  var mf = (D.multi_fail||[]).filter(function(m){{return m.fc>=2;}});
  var tbody = document.getElementById('multiFailBody');
  tbody.innerHTML = '';
  if(!mf.length) {{
    tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:#9ca3af;padding:24px">No multi-failure SNs</td></tr>';
  }} else {{
    // Build a quick lookup: sn → top fail item names from fail_point_stats fail_recs
    var snFailItems = {{}};
    (D.fail_point_stats||[]).forEach(function(fp) {{
      (fp.fail_recs||[]).forEach(function(r) {{
        if(!snFailItems[r.sn]) snFailItems[r.sn] = [];
        snFailItems[r.sn].push(fp.name);
      }});
    }});
    mf.forEach(function(m) {{
      var items = (snFailItems[m.sn]||[]).slice(0,2).join(', ') || '—';
      tbody.innerHTML += '<tr>'
        +'<td><code style="font-size:11px">'+m.sn+'</code></td>'
        +'<td><span class="badge b-fail">'+m.fc+'</span></td>'
        +'<td style="font-size:12px;color:#374151">'+items+'</td>'
        +'</tr>';
    }});
  }}
}}

// ---------------------------------------------------------------------------
// CPK TAB
// ---------------------------------------------------------------------------
function initCpk() {{
  var D = DATA;
  var items = D.cpk_list.filter(function(c){{return c.cpk<1.67;}}).slice(0,40);
  items.sort(function(a,b){{return a.cpk-b.cpk;}});
  var h = Math.max(300, items.length * 22);
  document.getElementById('cpkWrap').innerHTML = '<canvas id="cpkChart" style="height:'+h+'px"></canvas>';
  tryChart('cpkChart', {{
    type: 'bar',
    data: {{
      labels: items.map(function(c){{return c.name.length>45?c.name.slice(0,45)+'...':c.name;}}),
      datasets: [{{
        label: 'Cpk',
        data: items.map(function(c){{return c.cpk;}}),
        backgroundColor: items.map(function(c){{
          return c.cpk>=1.67?'#057a55':c.cpk>=1.33?'#1a56db':c.cpk>=1.0?'#d97706':c.cpk>=0.67?'#f59e0b':'#c81e1e';
        }}),
        borderRadius: 3
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{display:false}},
        tooltip: {{callbacks: {{label: function(ctx) {{
          var it = items[ctx.dataIndex];
          return ['Cpk: '+it.cpk+'  Cp: '+it.cp,'n: '+it.n+'  Mean: '+it.mean+'  σ: '+it.std,'LSL: '+it.lsl+'  USL: '+it.usl];
        }}}}}}
      }},
      scales: {{
        x: {{min:0,title:{{display:true,text:'Cpk'}}}},
        y: {{ticks:{{font:{{size:11}}}}}}
      }}
    }}
  }});

  var tb = document.querySelector('#cpkTbl tbody');
  tb.innerHTML = '';
  D.cpk_list.forEach(function(c) {{
    var tr = document.createElement('tr');
    tr.innerHTML = '<td>'+c.name+'</td>'
      +'<td>'+c.unit+'</td>'
      +'<td>'+c.n+'</td>'
      +'<td>'+fmt(c.mean,3)+'</td>'
      +'<td>'+fmt(c.std,4)+'</td>'
      +'<td>'+fmt(c.lsl,3)+'</td>'
      +'<td>'+fmt(c.usl,3)+'</td>'
      +'<td class="'+cpkCls(c.cp)+'">'+fmt(c.cp,3)+'</td>'
      +'<td class="'+cpkCls(c.cpk)+'">'+fmt(c.cpk,3)+'</td>'
      +'<td>'+cpkBadge(c.cpk)+'</td>';
    tb.appendChild(tr);
  }});
}}
function filterCpk() {{
  var kw = document.getElementById('cpkSearch').value.toLowerCase();
  document.querySelectorAll('#cpkTbl tbody tr').forEach(function(r) {{
    r.style.display = r.cells[0].textContent.toLowerCase().indexOf(kw)>=0 ? '' : 'none';
  }});
}}

// ---------------------------------------------------------------------------
// DISTRIBUTION TAB
// ---------------------------------------------------------------------------
var dHistInst = null;
function initDist() {{
  _ensureDistData();
  var btns = document.getElementById('distBtns');
  btns.innerHTML = '';  // clear before re-populating (needed on station switch)
  document.getElementById('distCard').style.display = 'none';
  var keys = Object.keys(DIST_DATA);
  if(!keys.length) {{
    btns.innerHTML = '<span style="color:#9ca3af">No distribution data</span>';
    return;
  }}
  keys.forEach(function(pname, i) {{
    var b = document.createElement('button');
    b.className = 'dist-btn' + (i===0?' active':'');
    b.textContent = pname.length>40 ? pname.slice(0,40)+'...' : pname;
    b.onclick = function() {{
      document.querySelectorAll('.dist-btn').forEach(function(x){{x.classList.remove('active');}});
      b.classList.add('active');
      showDist(pname);
    }};
    btns.appendChild(b);
  }});
  showDist(keys[0]);
}}
function showDist(pname) {{
  _ensureDistData();
  var d = DIST_DATA[pname];
  if(!d) return;
  document.getElementById('distCard').style.display = 'block';
  document.getElementById('distTitle').textContent = 'Distribution: '+pname;
  var vals  = d.vals.filter(function(v){{return v!==null;}});
  var failV = d.fail_vals.filter(function(v){{return v!==null;}});
  if(!vals.length) return;
  var mn=Math.min.apply(null,vals), mx=Math.max.apply(null,vals);
  var BINS=20, bw=(mx-mn)/BINS||1;
  var bins=new Array(BINS).fill(0), fBins=new Array(BINS).fill(0);
  vals.forEach(function(v){{var i=Math.min(Math.floor((v-mn)/bw),BINS-1);bins[i]++;}});
  failV.forEach(function(v){{var i=Math.min(Math.floor((v-mn)/bw),BINS-1);fBins[i]++;}});
  var lbls = bins.map(function(_,i){{return fmt(mn+i*bw+bw/2,2);}});
  if(dHistInst) dHistInst.destroy();
  dHistInst = tryChart('distHist', {{
    type: 'bar',
    data: {{
      labels: lbls,
      datasets: [
        {{label:'Pass',data:bins.map(function(v,i){{return v-fBins[i];}}),backgroundColor:'rgba(5,122,85,.45)',stack:'s'}},
        {{label:'Fail',data:fBins,backgroundColor:'rgba(200,30,30,.6)',stack:'s'}},
      ]
    }},
    options: {{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{position:'top'}}}},
      scales:{{
        x:{{title:{{display:true,text:d.unit}}}},
        y:{{title:{{display:true,text:'Count'}}}}
      }}
    }}
  }});
  var mean = vals.reduce(function(a,b){{return a+b;}},0)/vals.length;
  var std  = Math.sqrt(vals.reduce(function(a,b){{return a+(b-mean)*(b-mean);}},0)/vals.length);
  var cpkTxt = 'N/A';
  if(d.lsl!==null&&d.usl!==null&&std>1e-10) {{
    var cpk = Math.min((d.usl-mean)/(3*std),(mean-d.lsl)/(3*std));
    cpkTxt = cpk.toFixed(3);
  }}
  document.getElementById('distStats').innerHTML = [
    {{l:'N',      v:vals.length}},
    {{l:'Mean',   v:fmt(mean,3)+' '+d.unit}},
    {{l:'σ',      v:fmt(std,4)}},
    {{l:'Cpk',    v:cpkTxt}},
    {{l:'Min',    v:fmt(Math.min.apply(null,vals),3)}},
    {{l:'Max',    v:fmt(Math.max.apply(null,vals),3)}},
    {{l:'LSL',    v:fmt(d.lsl,3)}},
    {{l:'USL',    v:fmt(d.usl,3)}},
  ].map(function(s){{
    return '<div style="background:#f9fafb;border-radius:8px;padding:10px;border:1px solid #e5e7eb">'
      +'<div style="font-size:11px;color:#6b7280">'+s.l+'</div>'
      +'<div style="font-weight:700;font-size:13px">'+s.v+'</div></div>';
  }}).join('');
}}

// ---------------------------------------------------------------------------
// REPLAY TAB
// ---------------------------------------------------------------------------
var _snAll = [];
function _ensureSNDetail() {{
  if (!SN_DETAIL) {{
    SN_DETAIL = JSON.parse(document.getElementById(STATION_DATA[_curStype].sn_id).textContent);
  }}
}}
function _ensureDistData() {{
  if (!DIST_DATA) {{
    DIST_DATA = JSON.parse(document.getElementById(STATION_DATA[_curStype].dist_id).textContent);
  }}
}}
function initReplay() {{
  _ensureSNDetail();
  _snAll = Object.entries(SN_DETAIL).map(function(e) {{
    var sn=e[0], d=e[1];
    return {{sn:sn, overall:d.overall, time:d.time, fc:(d.fail_items?d.fail_items.length:0)}};
  }});
  _snAll.sort(function(a,b) {{
    var af=a.overall==='FAIL', bf=b.overall==='FAIL';
    if(af!==bf) return af?-1:1;
    return b.fc-a.fc;
  }});
  renderSNList(_snAll);
}}
function renderSNList(list) {{
  document.getElementById('snCnt').textContent = list.length+' total';
  document.getElementById('snList').innerHTML = list.map(function(s) {{
    var isFail = s.overall==='FAIL';
    return '<div class="sn-item" data-sn="'+s.sn+'" onclick="loadSN(this.dataset.sn,this)">'
      +'<div><div class="sn-code">'+s.sn+'</div>'
      +'<div class="sn-meta">'+(s.time?s.time.slice(0,16):'')+'</div></div>'
      +'<span class="badge '+(isFail?'b-fail':'b-pass')+'">'+(isFail?'✗'+s.fc:'✓')+'</span>'
      +'</div>';
  }}).join('');
}}
function filterSN() {{
  var kw  = document.getElementById('snSearch').value.toLowerCase();
  var flt = document.getElementById('snFilt').value;
  var list = _snAll;
  if(kw)  list = list.filter(function(s){{return s.sn.toLowerCase().indexOf(kw)>=0;}});
  if(flt) list = list.filter(function(s){{return s.overall.toUpperCase()===flt.toUpperCase();}});
  renderSNList(list);
}}
function loadSN(sn, el) {{
  document.querySelectorAll('.sn-item').forEach(function(x){{x.classList.remove('selected');}});
  el.classList.add('selected');
  _ensureSNDetail();
  var d = SN_DETAIL[sn];
  if(!d) {{
    document.getElementById('detPane').innerHTML = '<div style="padding:20px;color:#9ca3af">Failed to load data</div>';
    return;
  }}
  var fi = d.fail_items || [];
  var isFail = d.overall==='FAIL';
  var badge = isFail
    ? '<span style="background:#fef2f2;color:#991b1b;border:1px solid #fca5a5;border-radius:8px;padding:6px 18px;font-size:18px;font-weight:800">✗ FAIL</span>'
    : '<span style="background:#f0fdf4;color:#065f46;border:1px solid #86efac;border-radius:8px;padding:6px 18px;font-size:18px;font-weight:800">✓ PASS</span>';
  var html = '<div style="display:flex;align-items:center;gap:16px;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #e5e7eb">'
    +'<div>'+badge+'</div>'
    +'<div><div style="font-size:16px;font-weight:700;font-family:monospace">'+sn+'</div>'
    +'<div style="font-size:12px;color:#6b7280">Test Time: '+(d.time?d.time.slice(0,19):'N/A')+'</div></div></div>';

  if(fi.length) {{
    html += '<div style="margin-bottom:12px;padding:10px;background:#fef2f2;border-radius:8px;border:1px solid #fecaca">'
      +'<div style="font-weight:700;color:#991b1b;margin-bottom:6px">⚠ Failure Summary ('+fi.length+' items)</div>';
    fi.forEach(function(f) {{
      var d2=f.data, l=f.lsl, u=f.usl, note='';
      if(d2!==null&&d2!==undefined) {{
        if(u!==null&&d2>u) note='<span class="dev-tag">Above USL +'+(d2-u).toFixed(3)+f.unit+'</span>';
        else if(l!==null&&d2<l) note='<span class="dev-tag">Below LSL '+(d2-l).toFixed(3)+f.unit+'</span>';
      }}
      html += '<div style="display:flex;gap:8px;padding:4px 0;font-size:12px;align-items:center;border-bottom:1px solid #fecaca;flex-wrap:wrap">'
        +'<span style="color:#991b1b;font-weight:600;flex:2;min-width:180px">'+f.sheet+' > '+f.point+'</span>'
        +'<span style="color:#c81e1e;font-weight:700">'+fmt(d2,3)+' '+f.unit+'</span>'
        +'<span style="color:#6b7280">['+fmt(l,3)+', '+fmt(u,3)+']</span>'+note+'</div>';
    }});
    html += '</div>';
  }}

  var sheets = d.sheets || {{}};
  Object.keys(sheets).forEach(function(sheet) {{
    var rows = sheets[sheet];
    var hasFail = rows.some(function(r){{return r.result==='FAIL';}});
    html += '<div class="sheet-block">'
      +'<div class="sheet-hdr '+(hasFail?'has-fail':'')+'" onclick="togSh(this)">'+sheet
      +' '+(hasFail?'<span class="badge b-fail">Has Fail</span>':'<span class="badge b-pass">All Pass</span>')
      +'<span>'+(hasFail?'▼':'▶')+'</span></div>'
      +'<div class="sheet-rows '+(hasFail?'open':'')+'"><div style="display:grid;grid-template-columns:2fr 1fr 1fr 1fr 100px;gap:4px;padding:4px 10px;background:#f9fafb;font-size:11px;font-weight:700;color:#6b7280;border-radius:4px 4px 0 0"><span>Test Item</span><span>Value</span><span>LSL</span><span>USL</span><span>Result</span></div>';
    rows.forEach(function(r) {{
      var isFail2=r.result==='FAIL', dev='';
      if(isFail2&&r.data!==null) {{
        if(r.usl!==null&&r.data>r.usl) dev='+'+(r.data-r.usl).toFixed(3);
        else if(r.lsl!==null&&r.data<r.lsl) dev=''+(r.data-r.lsl).toFixed(3);
      }}
      html += '<div class="test-row '+(isFail2?'row-fail':'row-pass')+'">'
        +'<span style="font-size:11px;'+(isFail2?'font-weight:600;color:#991b1b':'')+';overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.point+'</span>'
        +'<span style="'+(isFail2?'font-weight:700;color:#c81e1e':'')+'">'+(r.data!==null?fmt(r.data,3)+' '+(r.unit||''):'N/A')+'</span>'
        +'<span style="color:#6b7280">'+fmt(r.lsl,3)+'</span>'
        +'<span style="color:#6b7280">'+fmt(r.usl,3)+'</span>'
        +'<span>'+(isFail2?'<span class="badge b-fail">Fail</span>'+(dev?'<span class="dev-tag">'+dev+'</span>':''):'<span class="badge b-pass">Pass</span>')+'</span></div>';
    }});
    html += '</div></div>';
  }});
  document.getElementById('detPane').innerHTML = html;
}}
function togSh(h) {{
  var r = h.nextElementSibling;
  r.classList.toggle('open');
  h.querySelector('span:last-child').textContent = r.classList.contains('open') ? '▼' : '▶';
}}
</script>
{json_scripts_html}
</body>
</html>"""

    return html


# ---------------------------------------------------------------------------
# JSON serialisation helper
# ---------------------------------------------------------------------------

def _json_default(obj):
    """Fallback JSON serialiser for numpy/pandas types."""
    try:
        import numpy as np
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            f = float(obj)
            return None if not math.isfinite(f) else f   # nan/inf → null
        if isinstance(obj, np.ndarray):
            return obj.tolist()
    except ImportError:
        pass
    raise TypeError(f'Object of type {type(obj).__name__} is not JSON serialisable')


def _sanitize_for_json(obj):
    """Recursively replace Python float NaN/Inf with None so output is valid JSON."""
    if isinstance(obj, float):
        return None if not math.isfinite(obj) else obj
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_for_json(v) for v in obj]
    return obj


def _esc(s: str) -> str:
    """Minimal HTML escape for text inserted into HTML attributes/content."""
    return (s
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


def _esc_js(s: str) -> str:
    """Escape a string for safe embedding inside a JS single-quoted string literal."""
    return (s
            .replace('\\', '\\\\')
            .replace("'", "\\'")
            .replace('\n', '\\n')
            .replace('\r', '\\r'))
