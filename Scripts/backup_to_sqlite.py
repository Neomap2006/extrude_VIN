"""
====================================================================
VELASTO PRODUCTION SYSTEM — DATABASE BACKUP SCRIPT
====================================================================
SAFE: Script ini HANYA READ dari SQL Server.
Tidak ada DELETE, DROP, TRUNCATE, atau UPDATE apapun.
====================================================================
Backup dari:
  - prd_extrude_hose  (ApplicationDbContext)
  - ELWP_PRD          (ElwpDbContext)

Output:
  - backup_velasto_YYYYMMDD_HHMMSS.sqlite
====================================================================
"""

import sqlite3
import sys
import os
from datetime import datetime

try:
    import pyodbc
except ImportError:
    print("[INSTALL] Menginstall pyodbc...")
    os.system(f"{sys.executable} -m pip install pyodbc")
    import pyodbc

# ── KONFIGURASI ──────────────────────────────────────────────────
SQL_SERVER   = "10.14.149.34,1433"
SQL_USER     = "usrvelasto"
SQL_PASSWORD = "H1s@na2025!!"

DB_MAIN = "prd_extrude_hose"   # ApplicationDbContext
DB_ELWP = "ELWP_PRD"           # ElwpDbContext

TIMESTAMP     = datetime.now().strftime("%Y%m%d_%H%M%S")
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
OUTPUT_SQLITE = os.path.join(SCRIPT_DIR, f"backup_velasto_{TIMESTAMP}.sqlite")

# ── TABEL YANG AKAN DI-BACKUP ────────────────────────────────────
# Format: (nama_tabel_sqlserver, nama_tabel_di_sqlite)
TABLES_MAIN = [
    ("ProductionReports",      "ProductionReports"),
    ("ProductionReadings",     "ProductionReadings"),
    ("SpsNoDocs",              "SpsNoDocs"),
    ("SpsItemLists",           "SpsItemLists"),
    ("DimensionReports",       "DimensionReports"),
    ("DimensionMeasurements",  "DimensionMeasurements"),
    ("DimensionSummaries",     "DimensionSummaries"),
    ("PlanningMasters",        "PlanningMasters"),
    ("Machines",               "Machines"),
    ("PartMasters",            "PartMasters"),
    ("ShiftMasters",           "ShiftMasters"),
    ("ProductionMaterialLots", "ProductionMaterialLots"),
    ("SensorIngestLogs",       "SensorIngestLogs"),
    ("ApprovalRequests",       "ApprovalRequests"),
    ("ApprovalRequestLogs",    "ApprovalRequestLogs"),
]

TABLES_ELWP = [
    ("[produksi].[tb_elwp_produksi_plannings]", "elwp_plannings"),
    ("[produksi].[tb_elwp_produksi_mesins]",    "elwp_machines"),
    ("[produksi].[tb_elwp_produksi_users]",     "elwp_users"),
    ("[produksi].[tb_elwp_produksi_areas]",     "elwp_areas"),
]

BATCH_SIZE = 500  # rows per commit (hemat memori)


def get_odbc_driver() -> str:
    """Cari driver ODBC SQL Server yang tersedia."""
    available = [d for d in pyodbc.drivers() if "SQL Server" in d]
    if not available:
        raise RuntimeError(
            "Tidak ada ODBC Driver SQL Server yang terinstall!\n"
            "Download dari: https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server"
        )
    for preferred in [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server",
    ]:
        if preferred in available:
            return preferred
    return available[-1]


def make_conn_str(database: str, driver: str) -> str:
    return (
        f"DRIVER={{{driver}}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={database};"
        f"UID={SQL_USER};"
        f"PWD={SQL_PASSWORD};"
        f"TrustServerCertificate=yes;"
        f"Encrypt=no;"
    )


def backup_table(sql_cursor, sqlite_conn: sqlite3.Connection,
                 src_table: str, dst_table: str) -> int:
    """
    Baca seluruh isi src_table dari SQL Server (READ ONLY),
    lalu tulis ke dst_table di SQLite.
    """
    print(f"  -> Backup {src_table} ...", end="", flush=True)

    try:
        sql_cursor.execute(f"SELECT * FROM {src_table}")
    except Exception as e:
        print(f" [SKIP: {e}]")
        return 0

    columns = [col[0] for col in sql_cursor.description]
    col_defs = ", ".join([f'"{c}" TEXT' for c in columns])
    col_names = ", ".join([f'"{c}"' for c in columns])
    placeholders = ", ".join(["?" for _ in columns])

    sqlite_conn.execute(f'CREATE TABLE IF NOT EXISTS "{dst_table}" ({col_defs})')
    sqlite_conn.execute(f'DELETE FROM "{dst_table}"')  # hapus data lama di sqlite saja

    insert_sql = f'INSERT INTO "{dst_table}" ({col_names}) VALUES ({placeholders})'

    total = 0
    batch = []
    for row in sql_cursor:
        batch.append(tuple(str(v) if v is not None else None for v in row))
        if len(batch) >= BATCH_SIZE:
            sqlite_conn.executemany(insert_sql, batch)
            sqlite_conn.commit()
            total += len(batch)
            batch.clear()

    if batch:
        sqlite_conn.executemany(insert_sql, batch)
        sqlite_conn.commit()
        total += len(batch)

    print(f" {total:,} rows OK")
    return total


def run_backup():
    print("=" * 60)
    print("  VELASTO DATABASE BACKUP")
    print("  MODE: READ ONLY (SQL Server tidak dimodifikasi)")
    print(f"  Output : {OUTPUT_SQLITE}")
    print(f"  Waktu  : {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print("=" * 60)

    try:
        driver = get_odbc_driver()
    except RuntimeError as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)

    print(f"\n[INFO] ODBC Driver: {driver}")

    sqlite_conn = sqlite3.connect(OUTPUT_SQLITE)
    sqlite_conn.execute("PRAGMA journal_mode=WAL")

    grand_total = 0
    errors = []

    # ── 1. BACKUP prd_extrude_hose ───────────────────────────────
    print(f"\n[1/2] Connecting ke: {DB_MAIN}")
    try:
        with pyodbc.connect(make_conn_str(DB_MAIN, driver), timeout=15) as mssql_conn:
            cursor = mssql_conn.cursor()
            print(f"  OK: Terhubung ke {DB_MAIN}\n")
            for src, dst in TABLES_MAIN:
                try:
                    grand_total += backup_table(cursor, sqlite_conn, src, dst)
                except Exception as e:
                    msg = f"{src}: {e}"
                    errors.append(msg)
                    print(f" [ERROR: {e}]")
    except pyodbc.Error as e:
        msg = f"Koneksi {DB_MAIN} gagal: {e}"
        print(f"\n[ERROR] {msg}")
        errors.append(msg)

    # ── 2. BACKUP ELWP_PRD ───────────────────────────────────────
    print(f"\n[2/2] Connecting ke: {DB_ELWP}")
    try:
        with pyodbc.connect(make_conn_str(DB_ELWP, driver), timeout=15) as mssql_conn:
            cursor = mssql_conn.cursor()
            print(f"  OK: Terhubung ke {DB_ELWP}\n")
            for src, dst in TABLES_ELWP:
                try:
                    grand_total += backup_table(cursor, sqlite_conn, src, dst)
                except Exception as e:
                    msg = f"{src}: {e}"
                    errors.append(msg)
                    print(f" [ERROR: {e}]")
    except pyodbc.Error as e:
        msg = f"Koneksi {DB_ELWP} gagal: {e}"
        print(f"\n[ERROR] {msg}")
        errors.append(msg)

    # ── SIMPAN METADATA BACKUP ───────────────────────────────────
    sqlite_conn.execute("""
        CREATE TABLE IF NOT EXISTS _backup_info (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    sqlite_conn.executemany(
        "INSERT OR REPLACE INTO _backup_info (key, value) VALUES (?, ?)",
        [
            ("backup_time",    datetime.now().isoformat()),
            ("source_server",  SQL_SERVER),
            ("source_db_main", DB_MAIN),
            ("source_db_elwp", DB_ELWP),
            ("total_rows",     str(grand_total)),
            ("errors",         "; ".join(errors) if errors else "none"),
        ]
    )
    sqlite_conn.commit()
    sqlite_conn.close()

    # ── RINGKASAN ────────────────────────────────────────────────
    size_mb = os.path.getsize(OUTPUT_SQLITE) / (1024 * 1024)
    print("\n" + "=" * 60)
    print("  BACKUP SELESAI")
    print(f"  File   : {OUTPUT_SQLITE}")
    print(f"  Size   : {size_mb:.2f} MB")
    print(f"  Total  : {grand_total:,} rows")
    if errors:
        print(f"  WARN   : {len(errors)} tabel gagal — cek log di atas")
    else:
        print("  STATUS : Semua tabel berhasil, tidak ada error")
    print("=" * 60)


if __name__ == "__main__":
    run_backup()
