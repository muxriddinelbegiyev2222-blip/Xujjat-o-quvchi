import sqlite3
from datetime import datetime

DB_NAME = "arxiv_hujjatlar.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            file_path TEXT,
            doc_year TEXT,
            doc_month TEXT,
            doc_day TEXT,
            region TEXT,
            organization TEXT,
            doc_type TEXT,
            doc_number TEXT,
            file_hash TEXT,
            saved_date TEXT,
            content TEXT
        )
    """)
    conn.commit()
    conn.close()

def is_duplicate(file_hash):
    if not file_hash:
        return None
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, file_name, file_path FROM documents WHERE file_hash = ?", (file_hash,))
    row = cursor.fetchone()
    conn.close()
    return row

def save_document_record(file_name, file_path, year, month, day, region, org, doc_type, doc_num, file_hash, content):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    saved_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO documents 
        (file_name, file_path, doc_year, doc_month, doc_day, region, org, doc_type, doc_number, file_hash, saved_date, content)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (file_name, file_path, str(year), str(month), str(day), region, org, doc_type, doc_num, file_hash, saved_date, content[:5000]))
    conn.commit()
    conn.close()
    return saved_date

def search_documents(query):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    param = f"%{query}%"
    cursor.execute("""
        SELECT file_name, region, organization, doc_type, doc_year || '-' || doc_month || '-' || doc_day, doc_number, file_path
        FROM documents
        WHERE content LIKE ? OR file_name LIKE ? OR doc_number LIKE ?
        ORDER BY id DESC
    """, (param, param, param))
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_statistics():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM documents")
    total = cursor.fetchone()[0]

    cursor.execute("SELECT region, COUNT(*) FROM documents GROUP BY region ORDER BY COUNT(*) DESC")
    regions = cursor.fetchall()

    cursor.execute("SELECT organization, COUNT(*) FROM documents GROUP BY organization ORDER BY COUNT(*) DESC")
    orgs = cursor.fetchall()

    cursor.execute("SELECT doc_type, COUNT(*) FROM documents GROUP BY doc_type ORDER BY COUNT(*) DESC")
    types = cursor.fetchall()

    cursor.execute("SELECT doc_year, COUNT(*) FROM documents GROUP BY doc_year ORDER BY doc_year DESC")
    years = cursor.fetchall()

    conn.close()
    return {
        "total": total,
        "regions": regions,
        "orgs": orgs,
        "types": types,
        "years": years
    }

def clear_all_data():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM documents")
    conn.commit()
    conn.close()
