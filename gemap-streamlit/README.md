# GEMAP — Panel de empresas

App en Streamlit para llevar el seguimiento de las empresas de la gestoría:
tareas contables pendientes, información por pedir, reuniones, responsable
y particularidades de cada cliente. Incluye vista de calendario y alertas
de vencimientos.

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en `http://localhost:8501`. Los datos se guardan en un archivo
SQLite local (`gemap.db`), que se crea solo la primera vez.

## Subir a GitHub

```bash
git init
git add .
git commit -m "GEMAP panel inicial"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/gemap.git
git push -u origin main
```

## Desplegar

**Opción A — Streamlit Community Cloud (gratis, más simple)**
1. Sube el repo a GitHub (paso anterior).
2. Entra en https://share.streamlit.io, conecta el repo y selecciona `app.py`.
3. Listo — te da un enlace público.
   - Nota: en el plan gratuito el disco no es persistente entre reinicios del
     contenedor; para uso serio, cambia `db.py` a una base de datos externa
     (ver más abajo) o usa Railway con un volumen.

**Opción B — Railway (con datos persistentes)**
1. Sube el repo a GitHub.
2. En Railway, "New Project" → "Deploy from GitHub repo" → selecciona `gemap`.
3. Railway detecta el `Dockerfile` automáticamente.
4. Añade un volumen montado en `/app/data` para que `gemap.db` no se borre
   en cada despliegue.
5. Genera un dominio público desde la pestaña "Settings" del servicio.

## Estructura

```
app.py           # interfaz Streamlit (Resumen, Calendario, ficha de empresa)
db.py            # capa de datos SQLite (sin dependencias externas)
assets/logo.png  # logo de GEMAP
requirements.txt
Dockerfile
```

## Pasar a una base de datos en la nube (opcional)

Si vas a tener varias personas usando la app a la vez, cambia `db.py` para
usar Postgres (por ejemplo el add-on de Postgres de Railway) en vez de
SQLite — la interfaz de `app.py` no necesita tocarse, solo las funciones
de `db.py`.
