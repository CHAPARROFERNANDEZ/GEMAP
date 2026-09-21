"""Capa de datos para GEMAP — SQLite, sin dependencias externas."""
import sqlite3
import os
import uuid
from datetime import datetime, date

DB_PATH = os.environ.get("GEMAP_DB_PATH", "gemap.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS empresas (
            id TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            responsable TEXT DEFAULT '',
            particularidades TEXT DEFAULT '',
            creado TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS tareas (
            id TEXT PRIMARY KEY,
            empresa_id TEXT NOT NULL,
            descripcion TEXT NOT NULL,
            tipo TEXT NOT NULL,           -- contable | informacion | reunion
            estado TEXT NOT NULL DEFAULT 'pendiente',
            fecha TEXT,                   -- fecha límite o de reunión (YYYY-MM-DD)
            creado TEXT,
            completado TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id)
        )
    """)
    conn.commit()
    conn.close()


def slugify(nombre: str) -> str:
    import re
    import unicodedata
    s = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower()
    return s or str(uuid.uuid4())[:8]


def ensure_empresa(nombre: str) -> str:
    conn = get_conn()
    eid = slugify(nombre)
    row = conn.execute("SELECT id FROM empresas WHERE id = ?", (eid,)).fetchone()
    if not row:
        conn.execute(
            "INSERT INTO empresas (id, nombre, responsable, particularidades, creado) VALUES (?, ?, '', '', ?)",
            (eid, nombre, datetime.utcnow().isoformat()),
        )
        conn.commit()
    conn.close()
    return eid


def list_empresas():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM empresas ORDER BY nombre COLLATE NOCASE").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_empresa(empresa_id: str):
    conn = get_conn()
    row = conn.execute("SELECT * FROM empresas WHERE id = ?", (empresa_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def update_empresa(empresa_id: str, responsable: str, particularidades: str):
    conn = get_conn()
    conn.execute(
        "UPDATE empresas SET responsable = ?, particularidades = ? WHERE id = ?",
        (responsable, particularidades, empresa_id),
    )
    conn.commit()
    conn.close()


def add_tarea(empresa_id: str, descripcion: str, tipo: str, fecha: str | None):
    conn = get_conn()
    tid = str(uuid.uuid4())
    conn.execute(
        "INSERT INTO tareas (id, empresa_id, descripcion, tipo, estado, fecha, creado, completado) "
        "VALUES (?, ?, ?, ?, 'pendiente', ?, ?, NULL)",
        (tid, empresa_id, descripcion, tipo, fecha, datetime.utcnow().isoformat()),
    )
    conn.commit()
    conn.close()
    return tid


def list_tareas(empresa_id: str | None = None):
    conn = get_conn()
    if empresa_id:
        rows = conn.execute("SELECT * FROM tareas WHERE empresa_id = ?", (empresa_id,)).fetchall()
    else:
        rows = conn.execute("SELECT * FROM tareas").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def toggle_tarea(tarea_id: str, done: bool):
    conn = get_conn()
    conn.execute(
        "UPDATE tareas SET estado = ?, completado = ? WHERE id = ?",
        ("hecho" if done else "pendiente", datetime.utcnow().isoformat() if done else None, tarea_id),
    )
    conn.commit()
    conn.close()


def delete_tarea(tarea_id: str):
    conn = get_conn()
    conn.execute("DELETE FROM tareas WHERE id = ?", (tarea_id,))
    conn.commit()
    conn.close()
