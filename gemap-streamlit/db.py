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
            actividad TEXT DEFAULT '',
            analitica INTEGER DEFAULT 0,
            sii INTEGER DEFAULT 0,
            creado TEXT
        )
    """)
    # Migración suave para bases de datos creadas con el esquema anterior
    existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(empresas)")}
    for col, ddl in [
        ("actividad", "ALTER TABLE empresas ADD COLUMN actividad TEXT DEFAULT ''"),
        ("analitica", "ALTER TABLE empresas ADD COLUMN analitica INTEGER DEFAULT 0"),
        ("sii", "ALTER TABLE empresas ADD COLUMN sii INTEGER DEFAULT 0"),
    ]:
        if col not in existing_cols:
            conn.execute(ddl)
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
            outlook_event_id TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id)
        )
    """)
    tareas_cols = {row["name"] for row in conn.execute("PRAGMA table_info(tareas)")}
    if "outlook_event_id" not in tareas_cols:
        conn.execute("ALTER TABLE tareas ADD COLUMN outlook_event_id TEXT")
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


def update_empresa(empresa_id: str, responsable: str, particularidades: str,
                    actividad: str = "", analitica: bool = False, sii: bool = False):
    conn = get_conn()
    conn.execute(
        "UPDATE empresas SET responsable = ?, particularidades = ?, actividad = ?, "
        "analitica = ?, sii = ? WHERE id = ?",
        (responsable, particularidades, actividad, int(analitica), int(sii), empresa_id),
    )
    conn.commit()
    conn.close()


def add_tarea(empresa_id: str, descripcion: str, tipo: str, fecha: str | None):
    conn = get_conn()
    tid = str(uuid.uuid4())
    conn.execute(
        "INSERT INTO tareas (id, empresa_id, descripcion, tipo, estado, fecha, creado, completado, outlook_event_id) "
        "VALUES (?, ?, ?, ?, 'pendiente', ?, ?, NULL, NULL)",
        (tid, empresa_id, descripcion, tipo, fecha, datetime.utcnow().isoformat()),
    )
    conn.commit()
    conn.close()
    return tid


def set_outlook_event_id(tarea_id: str, event_id: str):
    conn = get_conn()
    conn.execute("UPDATE tareas SET outlook_event_id = ? WHERE id = ?", (event_id, tarea_id))
    conn.commit()
    conn.close()


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
