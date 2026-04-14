"""
run_test.py  –  Command-line test runner (not part of the platform, for testing only)
Usage:
  python run_test.py         → N3 barcodes + B3 station folders
  python run_test.py N40     → N40 barcodes + B40 station folders
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(__file__))

from core.data_extractor import read_barcodes, run_extraction
from core.cpk_calculator import analyze_xlsx_folder
from core.html_report import generate_report

# ── Configuration ────────────────────────────────────────────────────────────
PRODUCT = sys.argv[1] if len(sys.argv) > 1 else 'N3'

if PRODUCT == 'N40':
    BARCODE_FILE = 'N40发货条码.xlsx'
    OUTPUT_DIR   = 'output_N40'
    STATION_CONFIGS = [
        {'type': 'FT1',  'folder': 'B40/FT1 1'},
        {'type': 'FT1',  'folder': 'B40/FT1 2'},
        {'type': 'FT1',  'folder': 'B40/FT1 3'},
        {'type': 'FT1',  'folder': 'B40/FT1 4'},
        {'type': 'FT1',  'folder': 'B40/FT1 5'},
        {'type': 'VSWR', 'folder': 'B40/VSWR'},
    ]
    REPORT_NAME = 'cpk_report_N40.html'
else:
    BARCODE_FILE = 'N3发货条码.xlsx'
    OUTPUT_DIR   = 'output_N3'
    STATION_CONFIGS = [
        {'type': 'FT1',  'folder': 'B3/FT1 1'},
        {'type': 'FT1',  'folder': 'B3/FT1 2'},
        {'type': 'FT1',  'folder': 'B3/FT1 3'},
        {'type': 'FT1',  'folder': 'B3/FT1 4'},
        {'type': 'FT1',  'folder': 'B3/FT1 5'},
        {'type': 'VSWR', 'folder': 'B3/VSWR'},
    ]
    REPORT_NAME = 'cpk_report_N3.html'
# ─────────────────────────────────────────────────────────────────────────────

def log(msg):
    print(msg, flush=True)

def progress(done, total, bc):
    if done % 50 == 0 or done == total:
        pct = 100 * done / max(total, 1)
        print(f'  Progress: {done}/{total}  ({pct:.0f}%)  latest: {bc}', flush=True)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Step 1 ─ Read barcodes
barcodes = read_barcodes(BARCODE_FILE)
print(f'\n[1/4] Read {len(barcodes)} barcodes from {BARCODE_FILE}')

# Step 2 ─ Extract files
print('\n[2/4] Extracting latest successful test files ...')
t0 = time.time()
summary = run_extraction(
    barcodes=barcodes,
    station_configs=STATION_CONFIGS,
    output_base_dir=OUTPUT_DIR,
    log_cb=log,
    progress_cb=progress,
)
print(f'\nExtraction done in {time.time()-t0:.1f}s')

for stype, info in summary.items():
    results = info['results']
    ok  = sum(1 for r in results if r['status'] == 'success')
    np_ = sum(1 for r in results if r['status'] == 'no_pass')
    nf  = sum(1 for r in results if r['status'] == 'not_found')
    nox = sum(1 for r in results if r['status'] == 'no_xlsx')
    xlsx_n = len(os.listdir(info['xlsx_dir']))
    print(f'  [{stype}] OK={ok}  no_pass={np_}  not_found={nf}  no_xlsx={nox}'
          f'  → {xlsx_n} xlsx files copied')

# Step 3 ─ CPK analysis
print('\n[3/4] Running CPK analysis ...')
t1 = time.time()
all_analysis = {}
for stype, info in summary.items():
    print(f'  Analysing [{stype}] ...')
    result = analyze_xlsx_folder(info['xlsx_dir'], log_cb=log)
    if result:
        all_analysis[stype] = result
        sheets = list(result.keys())
        total_pts = sum(len(pts) for pts in result.values())
        print(f'    → {len(sheets)} sheet(s), {total_pts} test points analysed')
    else:
        print(f'    → No data')
print(f'CPK analysis done in {time.time()-t1:.1f}s')

# Step 4 ─ Generate report
print('\n[4/4] Generating HTML report ...')
report_path = os.path.join(OUTPUT_DIR, REPORT_NAME)
generate_report(all_analysis, report_path)
print(f'Report written to: {os.path.abspath(report_path)}')

import webbrowser
webbrowser.open('file:///' + os.path.abspath(report_path).replace('\\', '/'))
print('\nDone! Browser should open automatically.')