"""
fault_db.py
-----------
SQLite persistence layer for fault analysis.

Tables:
  fault_rules   — keyword-based classification rules
  fault_records — one row per analysed test record (rich fields)
  fault_stats   — aggregated counts by fault_type
"""

import sqlite3
from datetime import datetime
from pathlib import Path


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

_DDL = """
PRAGMA journal_mode=WAL;

CREATE TABLE IF NOT EXISTS fault_rules (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    keywords    TEXT    NOT NULL,
    fault_type  TEXT    NOT NULL,
    suggestion  TEXT    DEFAULT '',
    created_at  TEXT    NOT NULL,
    updated_at  TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS fault_records (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    -- identification
    barcode          TEXT    DEFAULT '',   -- primary barcode (first of dual pair)
    barcode_full     TEXT    DEFAULT '',   -- original folder name (dual: BC1_BC2)
    station          TEXT    DEFAULT '',   -- user-configured station label (FT1, Aging…)
    station_machine  TEXT    DEFAULT '',   -- physical machine id (FT1_1 / EQP_ID from yml)
    product_category TEXT    DEFAULT '',   -- from TestResult sub-dir (ORBI_B3, ORBI_B40…)
    product_code     TEXT    DEFAULT '',   -- board revision/fixture slot (X11_X11, R1B…)
    test_time        TEXT    DEFAULT '',
    status           TEXT    DEFAULT '',   -- 'pass' / 'fail' / 'unknown'
    -- fault classification
    fault_type       TEXT    DEFAULT '未分类故障',
    matched_rule_id  INTEGER DEFAULT NULL,
    -- rich context extracted from record
    first_fail_desc  TEXT    DEFAULT '',   -- DutInfo.FirstFailCaseDescription (MEASUREMENT JSON)
    failed_items     TEXT    DEFAULT NULL, -- JSON: [{item,value,unit,lsl,usl,direction,deviation}]
    equip_errors     TEXT    DEFAULT NULL, -- JSON: [{pattern,detail,raw_line}]
    instruments      TEXT    DEFAULT NULL, -- JSON: {SA1:"TCPIP0::…", EQP_ID:"FT_1", …}
    log_excerpt      TEXT    DEFAULT '',
    log_path         TEXT    DEFAULT '',
    llm_analysis     TEXT    DEFAULT NULL,
    created_at       TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS fault_stats (
    fault_type  TEXT PRIMARY KEY,
    count       INTEGER DEFAULT 0,
    last_seen   TEXT    DEFAULT ''
);
"""

# Columns added after initial schema — applied as migrations on older DBs
_NEW_COLUMNS = [
    ('fault_records', 'barcode_full',     "TEXT DEFAULT ''"),
    ('fault_records', 'station_machine',  "TEXT DEFAULT ''"),
    ('fault_records', 'product_category', "TEXT DEFAULT ''"),
    ('fault_records', 'product_code',     "TEXT DEFAULT ''"),
    ('fault_records', 'first_fail_desc',  "TEXT DEFAULT ''"),
    ('fault_records', 'failed_items',     'TEXT DEFAULT NULL'),
    ('fault_records', 'equip_errors',     'TEXT DEFAULT NULL'),
    ('fault_records', 'instruments',      'TEXT DEFAULT NULL'),
]

_SEED_RULES = [
    ('power supply,voltage,current,电源,供电',          '程控电源',       '检查程控电源输出电压/电流是否正常，确认接线可靠'),
    ('RF switch,switch,开关,SW,COM',                   '射频开关',       '检查RF开关控制信号及连接，确认切换逻辑正确'),
    ('calibration,cal,校准,insertion loss,链路补偿',   '校准/链路插损',  '重新校准相关仪器，检查校准文件日期和补偿值'),
    ('signal generator,信号源,SG,source',              '信号源',         '检查信号源输出幅度/频率，确认已解锁'),
    ('spectrum analyzer,频谱,SA,MXA',                  '频谱仪',         '检查频谱仪连接和参数配置'),
    ('power meter,功率计,PM,power measurement',        '功率计',         '检查功率计零校及量程设置'),
    ('VNA,network analyzer,S参数,S11,S21',             '矢量网络分析仪', '检查VNA校准和端口连接'),
    ('RF cable,cable,射频线,连接器,VSWR',              '射频电缆',       '检查射频线缆连接是否松动或损坏'),
    ('network,timeout,超时,连接失败,socket',           '网络仪器连接',   '检查测试PC网络连接，重启网络服务'),
    ('USB,usb,VID,PID,device not found',               'USB接口',        '重新插拔USB，检查驱动程序'),
    ('serial,COM,UART,串口,SerialException',           '串口通信',       '确认串口号及波特率配置，检查线缆'),
    ('ATE,automation,测试程序,script,exception,traceback', 'ATE程序',    '查看ATE异常日志，联系测试程序开发人员'),
    ('firmware,固件,ZBOOT,SLOT,版本,download fail',   'DUT软件/固件',   '重新烧录固件，检查版本匹配'),
    ('hardware,硬件,元件,component,PAM接触',           'DUT硬件',        '检查DUT硬件状态，必要时返修'),
    ('SSH,ssh,retry,connect refuse',                   'DUT通信(SSH)',   '检查DUT IP地址及SSH服务，确认网络连通'),
    ('PA CURR,IDLE CURR,电流异常',                     'PA电流异常',     '检查PA供电及器件状态，排查过流或欠流'),
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _conn(db_path) -> sqlite3.Connection:
    con = sqlite3.connect(str(db_path))
    con.row_factory = sqlite3.Row
    return con


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _migrate(con: sqlite3.Connection) -> None:
    """Add new columns to existing tables idempotently."""
    existing = {row[1] for row in con.execute('PRAGMA table_info(fault_records)')}
    for table, col, coldef in _NEW_COLUMNS:
        if col not in existing:
            con.execute(f'ALTER TABLE {table} ADD COLUMN {col} {coldef}')


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def init_db(db_path) -> None:
    """Create tables, run migrations, and pre-populate seed rules if empty."""
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with _conn(db_path) as con:
        con.executescript(_DDL)
        _migrate(con)
        cur = con.execute('SELECT COUNT(*) FROM fault_rules')
        if cur.fetchone()[0] == 0:
            ts = _now()
            con.executemany(
                'INSERT INTO fault_rules (keywords, fault_type, suggestion, created_at, updated_at) '
                'VALUES (?, ?, ?, ?, ?)',
                [(kw, ft, sg, ts, ts) for kw, ft, sg in _SEED_RULES]
            )


def get_rules(db_path) -> list:
    with _conn(db_path) as con:
        rows = con.execute(
            'SELECT id, keywords, fault_type, suggestion, created_at, updated_at '
            'FROM fault_rules ORDER BY id'
        ).fetchall()
    return [dict(r) for r in rows]


def add_rule(db_path, keywords: str, fault_type: str, suggestion: str = '') -> int:
    ts = _now()
    with _conn(db_path) as con:
        cur = con.execute(
            'INSERT INTO fault_rules (keywords, fault_type, suggestion, created_at, updated_at) '
            'VALUES (?, ?, ?, ?, ?)',
            (keywords, fault_type, suggestion, ts, ts)
        )
        return cur.lastrowid


def update_rule(db_path, rule_id: int, keywords: str = None,
                fault_type: str = None, suggestion: str = None) -> None:
    fields, vals = [], []
    if keywords  is not None: fields.append('keywords = ?');   vals.append(keywords)
    if fault_type is not None: fields.append('fault_type = ?'); vals.append(fault_type)
    if suggestion is not None: fields.append('suggestion = ?'); vals.append(suggestion)
    if not fields:
        return
    fields.append('updated_at = ?'); vals.append(_now())
    vals.append(rule_id)
    with _conn(db_path) as con:
        con.execute(f'UPDATE fault_rules SET {", ".join(fields)} WHERE id = ?', vals)


def delete_rule(db_path, rule_id: int) -> None:
    with _conn(db_path) as con:
        con.execute('DELETE FROM fault_rules WHERE id = ?', (rule_id,))


def add_record(db_path, *, barcode: str, station: str, test_time: str, status: str,
               fault_type: str, matched_rule_id=None,
               # new rich fields
               barcode_full: str = '',
               station_machine: str = '',
               product_category: str = '',
               product_code: str = '',
               first_fail_desc: str = '',
               failed_items=None,      # list or None → stored as JSON
               equip_errors=None,      # list or None → stored as JSON
               instruments=None,       # dict or None → stored as JSON
               log_excerpt: str = '',
               log_path: str = '',
               llm_analysis: str = None) -> int:
    """Insert a fault record and update fault_stats. Returns new row id."""
    import json as _json
    ts = _now()
    fi_json  = _json.dumps(failed_items,  ensure_ascii=False) if failed_items  is not None else None
    ee_json  = _json.dumps(equip_errors,  ensure_ascii=False) if equip_errors  is not None else None
    ins_json = _json.dumps(instruments,   ensure_ascii=False) if instruments   is not None else None

    with _conn(db_path) as con:
        cur = con.execute(
            '''INSERT INTO fault_records
               (barcode, barcode_full, station, station_machine,
                product_category, product_code, test_time, status,
                fault_type, matched_rule_id,
                first_fail_desc, failed_items, equip_errors, instruments,
                log_excerpt, log_path, llm_analysis, created_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (barcode, barcode_full, station, station_machine,
             product_category, product_code, test_time, status,
             fault_type, matched_rule_id,
             first_fail_desc, fi_json, ee_json, ins_json,
             log_excerpt, log_path, llm_analysis, ts)
        )
        row_id = cur.lastrowid
        con.execute(
            'INSERT INTO fault_stats (fault_type, count, last_seen) VALUES (?, 1, ?) '
            'ON CONFLICT(fault_type) DO UPDATE SET '
            '  count = count + 1, last_seen = excluded.last_seen',
            (fault_type, test_time or ts)
        )
    return row_id


def get_records(db_path, limit: int = 500, fault_type: str = None,
                station: str = None, barcode: str = None) -> list:
    sql = 'SELECT * FROM fault_records'
    conds, params = [], []
    if fault_type: conds.append('fault_type = ?');   params.append(fault_type)
    if station:    conds.append('station = ?');       params.append(station)
    if barcode:    conds.append('barcode = ?');       params.append(barcode)
    if conds:
        sql += ' WHERE ' + ' AND '.join(conds)
    sql += ' ORDER BY id DESC LIMIT ?'
    params.append(limit)
    with _conn(db_path) as con:
        rows = con.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def get_cross_station_barcodes(db_path) -> list:
    """
    Return barcodes that appear in more than one physical station machine.
    Each row: {barcode, machines (comma-sep), record_count, has_fail}
    Useful for diagnosing DUTs retested at different stations after failure.
    """
    sql = '''
        SELECT barcode,
               GROUP_CONCAT(DISTINCT station_machine) AS machines,
               COUNT(*)                               AS record_count,
               MAX(CASE WHEN status = 'fail' THEN 1 ELSE 0 END) AS has_fail
        FROM fault_records
        WHERE barcode != ''
        GROUP BY barcode
        HAVING COUNT(DISTINCT station_machine) > 1
        ORDER BY record_count DESC
    '''
    with _conn(db_path) as con:
        rows = con.execute(sql).fetchall()
    return [dict(r) for r in rows]


def get_unclassified_records(db_path, limit: int = 200) -> list:
    with _conn(db_path) as con:
        rows = con.execute(
            "SELECT * FROM fault_records WHERE fault_type = '未分类故障' "
            'ORDER BY id DESC LIMIT ?',
            (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def update_record_fault_type(db_path, record_id: int, fault_type: str) -> None:
    with _conn(db_path) as con:
        row = con.execute(
            'SELECT fault_type, test_time FROM fault_records WHERE id = ?', (record_id,)
        ).fetchone()
        if not row:
            return
        old_type = row['fault_type']
        ts = row['test_time'] or _now()
        con.execute('UPDATE fault_records SET fault_type = ? WHERE id = ?',
                    (fault_type, record_id))
        con.execute(
            'UPDATE fault_stats SET count = MAX(0, count - 1) WHERE fault_type = ?',
            (old_type,)
        )
        con.execute(
            'INSERT INTO fault_stats (fault_type, count, last_seen) VALUES (?, 1, ?) '
            'ON CONFLICT(fault_type) DO UPDATE SET '
            '  count = count + 1, last_seen = excluded.last_seen',
            (fault_type, ts)
        )


def get_stats(db_path) -> list:
    with _conn(db_path) as con:
        rows = con.execute(
            'SELECT fault_type, count, last_seen FROM fault_stats ORDER BY count DESC'
        ).fetchall()
    return [dict(r) for r in rows]


def clear_records(db_path) -> None:
    """Delete all fault_records and reset fault_stats (keeps rules intact)."""
    with _conn(db_path) as con:
        con.execute('DELETE FROM fault_records')
        con.execute('DELETE FROM fault_stats')
