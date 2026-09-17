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
            region TEXT,
            organization TEXT,
            doc_type TEXT,
            saved_date TEXT,
            content TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_document_record(file_name, file_path, region, org, doc_type, content):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    saved_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO documents (file_name, file_path, region, org, doc_type, saved_date, content)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (file_name, file_path, region, org, doc_type, saved_date, content))
    conn.commit()
    conn.close()
    return saved_date

def search_documents(query):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    search_term = f"%{query}%"
    cursor.execute("""
        SELECT file_name, file_path, region, org, doc_type, saved_date 
        FROM documents 
        WHERE content LIKE ? OR file_name LIKE ?
        ORDER BY id DESC
    """, (search_term, search_term))
    results = cursor.fetchall()
    conn.close()
    return results

def get_statistics():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM documents")
    total_docs = cursor.fetchone()[0]

    cursor.execute("SELECT region, COUNT(*) FROM documents GROUP BY region ORDER BY COUNT(*) DESC")
    region_stats = cursor.fetchall()

    cursor.execute("SELECT organization, COUNT(*) FROM documents GROUP BY organization ORDER BY COUNT(*) DESC")
    org_stats = cursor.fetchall()

    cursor.execute("SELECT doc_type, COUNT(*) FROM documents GROUP BY doc_type ORDER BY COUNT(*) DESC")
    type_stats = cursor.fetchall()

    conn.close()
    return {
        "total": total_docs,
        "regions": region_stats,
        "orgs": org_stats,
        "types": type_stats
    }

def clear_all_data():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM documents")
    conn.commit()
    conn.close()
