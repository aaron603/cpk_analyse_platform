"""
html_fail_report.py
-------------------
Generates a self-contained failure-analysis HTML report for the
folder_direct Scenario B mode.  No external dependencies.
"""

import json as _json
import os
from collections import Counter
from datetime import datetime


def generate_fail_report(
    fail_data: dict,
    output_path: str,
    title: str = '',
    generated_at: str = '',
) -> str:
    """
    Build and write a self-contained HTML failure analysis report.

    Parameters
    ----------
    fail_data : dict
        {stype: {barcode_stats, fail_barcodes, never_pass_barcodes, all_fail_items}}
    output_path : str
        Absolute path to write the .html file.
    title : str
        Product name / report title.
    generated_at : str
        Timestamp string for the report header.

    Returns
    -------
    str — absolute path of the written file.
    """
    if not generated_at:
        generated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if not title:
        title = 'Failure Analysis Report - Zillnk'
    else:
        title = f'{title} Failure Analysis Report - Zillnk'

    html = _build_html(fail_data, title, generated_at)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return output_path


# ---------------------------------------------------------------------------
# HTML builder
# ---------------------------------------------------------------------------

def _build_html(fail_data: dict, title: str, generated_at: str) -> str:
    # ── Aggregate statistics ──────────────────────────────────────────────
    total_barcodes  = 0
    total_fail_bc   = 0
    total_never_bc  = 0
    all_item_counts: Counter = Counter()   # (sheet, point_name) → count
    all_fail_barcodes = []   # [{bc, stype, fail_count, pass_count, latest, item_count}]
    all_never_barcodes = []  # [{bc, stype, tests, latest}]

    for stype, sdata in fail_data.items():
        total_barcodes  += len(sdata.get('barcode_stats', {}))
        total_fail_bc   += len(sdata.get('fail_barcodes', {}))
        total_never_bc  += len(sdata.get('never_pass_barcodes', []))

        for bc, st in sdata.get('fail_barcodes', {}).items():
            latest = max(st['times']) if st['times'] else ''
            all_fail_barcodes.append({
                'bc':         bc,
                'stype':      stype,
                'fail_count': st['fail_count'],
                'pass_count': st['pass_count'],
                'items':      len(st['fail_items']),
                'latest':     latest,
            })

        for bc in sdata.get('never_pass_barcodes', []):
            st = sdata['barcode_stats'].get(bc, {})
            latest = max(st.get('times', [])) if st.get('times') else ''
            all_never_barcodes.append({
                'bc':     bc,
                'stype':  stype,
                'tests':  st.get('fail_count', 0),
                'latest': latest,
            })

        for rec in sdata.get('all_fail_items', []):
            bc, time_str, sheet, point, data, lsl, usl, dev = rec
            key = f'{sheet} / {point}' if point else sheet
            all_item_counts[key] += 1

    fail_rate = (total_fail_bc / total_barcodes * 100) if total_barcodes else 0
    never_rate = (total_never_bc / total_barcodes * 100) if total_barcodes else 0

    # Top-20 failed items for Pareto chart
    pareto_items = all_item_counts.most_common(20)

    # ── Render HTML sections ──────────────────────────────────────────────
    pareto_html  = _render_pareto(pareto_items)
    fail_bc_html = _render_fail_barcodes(all_fail_barcodes)
    never_html   = _render_never_barcodes(all_never_barcodes)

    card_class_fail  = 'card-red'  if fail_rate  > 10 else ('card-orange' if fail_rate  > 0 else 'card-green')
    card_class_never = 'card-red'  if never_rate > 5  else ('card-orange' if never_rate > 0 else 'card-green')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_esc(title)}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#212121;font-size:13px}}
header{{background:linear-gradient(135deg,#1a237e,#3949ab);color:#fff;padding:16px 24px}}
header h1{{font-size:18px;font-weight:600;margin-bottom:4px}}
header p{{font-size:11px;opacity:.8}}
.container{{max-width:1200px;margin:0 auto;padding:16px 20px}}
.cards{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}}
.card{{flex:1;min-width:160px;background:#fff;border-radius:8px;padding:14px 18px;
       box-shadow:0 1px 4px rgba(0,0,0,.1);border-top:4px solid #ccc}}
.card-green{{border-top-color:#43a047}}
.card-orange{{border-top-color:#fb8c00}}
.card-red{{border-top-color:#e53935}}
.card-blue{{border-top-color:#1e88e5}}
.card .val{{font-size:26px;font-weight:700;margin:6px 0}}
.card .lbl{{font-size:11px;color:#666}}
section{{background:#fff;border-radius:8px;padding:16px 20px;
         box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:16px}}
section h2{{font-size:14px;font-weight:600;color:#1a237e;margin-bottom:12px;
            border-bottom:2px solid #e8eaf6;padding-bottom:6px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{background:#1a237e;color:#fff;padding:7px 10px;text-align:center;font-weight:500}}
td{{padding:6px 10px;border-bottom:1px solid #e0e0e0;text-align:center}}
tr:hover td{{background:#f5f5f5}}
.bc-cell{{text-align:left;font-family:monospace;font-size:11px}}
.fail-row td{{background:#fff8f8}}
.never-row td{{background:#fffde7}}
.badge-fail{{display:inline-block;background:#ffebee;color:#c62828;
             border-radius:4px;padding:1px 6px;font-size:10px;font-weight:600}}
.badge-never{{display:inline-block;background:#fff3e0;color:#e65100;
              border-radius:4px;padding:1px 6px;font-size:10px;font-weight:600}}
.chart-wrap{{overflow-x:auto;padding:4px 0}}
svg.pareto{{min-width:600px}}
.pareto-bar{{fill:#3949ab}}
.pareto-bar:hover{{fill:#1a237e}}
.pareto-line{{fill:none;stroke:#e53935;stroke-width:2}}
.pareto-dot{{fill:#e53935}}
.no-data{{color:#aaa;font-style:italic;padding:20px;text-align:center}}
.collapse-btn{{background:none;border:none;cursor:pointer;color:#3949ab;
               font-size:12px;font-weight:600;padding:2px 0;margin-bottom:8px}}
.collapse-btn:hover{{text-decoration:underline}}
</style>
</head>
<body>
<header>
  <h1>{_esc(title)}</h1>
  <p>Generated: {_esc(generated_at)}</p>
</header>
<div class="container">

<!-- Summary cards -->
<div class="cards">
  <div class="card card-blue">
    <div class="lbl">Total Barcodes</div>
    <div class="val">{total_barcodes}</div>
  </div>
  <div class="card {card_class_fail}">
    <div class="lbl">With Failures</div>
    <div class="val">{total_fail_bc}</div>
    <div class="lbl">{fail_rate:.1f}%</div>
  </div>
  <div class="card {card_class_never}">
    <div class="lbl">Never Passed</div>
    <div class="val">{total_never_bc}</div>
    <div class="lbl">{never_rate:.1f}%</div>
  </div>
  <div class="card card-orange">
    <div class="lbl">Failed Item Types</div>
    <div class="val">{len(all_item_counts)}</div>
  </div>
</div>

<!-- Pareto chart -->
<section>
  <h2>High-Frequency Failed Test Items — Pareto Analysis (Top {len(pareto_items)})</h2>
  {pareto_html}
</section>

<!-- Failed barcodes table -->
<section>
  <h2>Failed Barcode List <span class="badge-fail">{total_fail_bc}</span></h2>
  {fail_bc_html}
</section>

<!-- Never-passed barcodes -->
<section>
  <h2>Never-Passed Barcodes <span class="badge-never">{total_never_bc}</span></h2>
  {never_html}
</section>

</div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Section renderers
# ---------------------------------------------------------------------------

def _render_pareto(items: list) -> str:
    if not items:
        return '<p class="no-data">No failed test item data</p>'

    labels = [_esc(k) for k, _ in items]
    counts = [v for _, v in items]
    total  = sum(counts)
    cumsum = []
    running = 0
    for c in counts:
        running += c
        cumsum.append(running / total * 100)

    n      = len(counts)
    w      = max(600, n * 44)
    pad_l  = 60
    pad_r  = 60
    pad_t  = 20
    pad_b  = 110
    h      = 280
    chart_w = w - pad_l - pad_r
    chart_h = h - pad_t - pad_b
    bar_w   = max(10, chart_w // n - 4)
    max_c   = max(counts) if counts else 1

    bars_svg = []
    for i, (c, cum, lbl) in enumerate(zip(counts, cumsum, labels)):
        x      = pad_l + i * (chart_w // n) + (chart_w // n - bar_w) // 2
        bar_h  = int(c / max_c * chart_h)
        y      = pad_t + chart_h - bar_h
        bars_svg.append(
            f'<rect class="pareto-bar" x="{x}" y="{y}" width="{bar_w}" height="{bar_h}">'
            f'<title>{lbl}: {c}</title></rect>'
        )
        # Label below x-axis (rotated)
        lx = x + bar_w // 2
        bars_svg.append(
            f'<text x="{lx}" y="{pad_t + chart_h + 14}" '
            f'font-size="9" fill="#333" text-anchor="end" '
            f'transform="rotate(-35,{lx},{pad_t + chart_h + 14})">{lbl[:30]}</text>'
        )
        # Count on bar top
        if bar_h > 14:
            bars_svg.append(
                f'<text x="{x + bar_w // 2}" y="{y - 3}" '
                f'font-size="9" fill="#1a237e" text-anchor="middle">{c}</text>'
            )

    # Cumulative line
    pts = []
    for i, cum in enumerate(cumsum):
        x = pad_l + i * (chart_w // n) + (chart_w // n) // 2
        y = pad_t + int(chart_h * (1 - cum / 100))
        pts.append(f'{x},{y}')
    line_svg = (f'<polyline class="pareto-line" points="{" ".join(pts)}"/>'
                if pts else '')
    dots_svg = ''.join(
        f'<circle class="pareto-dot" cx="{pt.split(",")[0]}" cy="{pt.split(",")[1]}" r="3">'
        f'<title>{cumsum[i]:.1f}%</title></circle>'
        for i, pt in enumerate(pts)
    )

    # Y-axis labels
    y_labels = ''
    for pct in (0, 25, 50, 75, 100):
        yy = pad_t + chart_h - int(chart_h * pct / 100)
        cnt_at = int(max_c * pct / 100)
        y_labels += (
            f'<line x1="{pad_l - 4}" y1="{yy}" x2="{w - pad_r}" y2="{yy}" '
            f'stroke="#e0e0e0" stroke-dasharray="3,3"/>'
            f'<text x="{pad_l - 8}" y="{yy + 4}" font-size="9" fill="#666" text-anchor="end">'
            f'{cnt_at}</text>'
        )
    # Right axis (cumulative %)
    for pct in (0, 25, 50, 75, 100):
        yy = pad_t + chart_h - int(chart_h * pct / 100)
        y_labels += (
            f'<text x="{w - pad_r + 6}" y="{yy + 4}" font-size="9" fill="#e53935">'
            f'{pct}%</text>'
        )

    svg = (
        f'<div class="chart-wrap">'
        f'<svg class="pareto" width="{w}" height="{h}" xmlns="http://www.w3.org/2000/svg">'
        f'{y_labels}'
        f'{"".join(bars_svg)}'
        f'{line_svg}{dots_svg}'
        f'<text x="{pad_l - 8}" y="{pad_t - 6}" font-size="9" fill="#1a237e">Count</text>'
        f'<text x="{w - pad_r + 6}" y="{pad_t - 6}" font-size="9" fill="#e53935">Cum%</text>'
        f'</svg></div>'
    )
    return svg


def _render_fail_barcodes(rows: list) -> str:
    if not rows:
        return '<p class="no-data">No failed barcodes</p>'
    rows_sorted = sorted(rows, key=lambda r: r['fail_count'], reverse=True)
    trs = ''.join(
        f'<tr class="fail-row">'
        f'<td class="bc-cell">{_esc(r["bc"])}</td>'
        f'<td>{_esc(r["stype"])}</td>'
        f'<td>{r["fail_count"]}</td>'
        f'<td>{r["pass_count"]}</td>'
        f'<td>{r["items"]}</td>'
        f'<td>{_esc(r["latest"])}</td>'
        f'</tr>'
        for r in rows_sorted
    )
    return (
        '<table>'
        '<thead><tr>'
        '<th style="text-align:left">Barcode</th>'
        '<th>Station</th><th>Fail Count</th><th>Pass Count</th>'
        '<th>Failed Items</th><th>Latest Test Time</th>'
        '</tr></thead>'
        f'<tbody>{trs}</tbody>'
        '</table>'
    )


def _render_never_barcodes(rows: list) -> str:
    if not rows:
        return '<p class="no-data">No never-passed barcodes</p>'
    rows_sorted = sorted(rows, key=lambda r: r['tests'], reverse=True)
    trs = ''.join(
        f'<tr class="never-row">'
        f'<td class="bc-cell">{_esc(r["bc"])}</td>'
        f'<td>{_esc(r["stype"])}</td>'
        f'<td>{r["tests"]}</td>'
        f'<td>{_esc(r["latest"])}</td>'
        f'</tr>'
        for r in rows_sorted
    )
    return (
        '<table>'
        '<thead><tr>'
        '<th style="text-align:left">Barcode</th>'
        '<th>Station</th><th>Total Tests</th><th>Latest Test Time</th>'
        '</tr></thead>'
        f'<tbody>{trs}</tbody>'
        '</table>'
    )


def _esc(s) -> str:
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))
