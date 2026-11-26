"""Bot SEGUNDO CEREBRO (GEMAP).

Recibe avisos de Google Drive (Apps Script), lanza la rutina de revisión de Claude
y reenvía los avisos a Telegram. Solo responde al chat de Yuri.
"""
import asyncio
import hmac
import io
import os
import time
from collections import defaultdict

import httpx
from fastapi import FastAPI, Header, HTTPException, Request

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"].strip().strip('"\'')
OWNER_CHAT_ID = os.environ.get("OWNER_CHAT_ID", "")           # chat de Yuri
SHARED_SECRET = os.environ["SHARED_SECRET"].strip()                   # Apps Script -> bot
WEBHOOK_PATH = os.environ["WEBHOOK_PATH"].strip().strip("/")                     # ruta secreta de Telegram
ROUTINE_FIRE_URL = os.environ.get("ROUTINE_FIRE_URL", "")     # URL del activador API
ROUTINE_TOKEN = os.environ.get("ROUTINE_TOKEN", "")           # token de la rutina
DEBOUNCE_SECONDS = int(os.environ.get("DEBOUNCE_SECONDS", "180"))
MIN_SECONDS_BETWEEN_RUNS = int(os.environ.get("MIN_SECONDS_BETWEEN_RUNS", "600"))

app = FastAPI()


@app.on_event("startup")
async def registrar_webhook():
    """Comprueba el token y registra el webhook de Telegram. No escribe el token en los logs."""
    dominio = os.environ.get("RAILWAY_PUBLIC_DOMAIN", "")
    try:
        async with httpx.AsyncClient(timeout=20) as c:
            me = await c.get(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getMe")
            print(f"[telegram] getMe -> HTTP {me.status_code} {me.json().get('description', 'ok')}", flush=True)
            if me.status_code != 200 or not dominio:
                print(f"[telegram] sin webhook (dominio={'si' if dominio else 'no'})", flush=True)
                return
            r = await c.post(
                f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/setWebhook",
                json={"url": f"https://{dominio}/telegram/{WEBHOOK_PATH}"},
            )
            print(f"[telegram] setWebhook -> HTTP {r.status_code} {r.json().get('description', '')}", flush=True)
    except Exception as e:  # el bot debe arrancar aunque falle esto
        print(f"[telegram] error al registrar webhook: {type(e).__name__}", flush=True)
pending = defaultdict(list)      # empresa -> archivos nuevos
pendientes = {}                  # empresa -> tareas abiertas (las publica el Apps Script)
asientos = {}                    # empresa -> filas de asientos propuestos (las publica el Apps Script)
seen = {}                        # id de archivo -> fecha de modificación ya procesada
last_event = 0.0
last_fire = 0.0
worker = None


def check_secret(value: str):
    if not value or not hmac.compare_digest(value, SHARED_SECRET):
        raise HTTPException(status_code=401)


async def telegram(text: str, chat_id: str = "", teclado=None):
    chat = chat_id or OWNER_CHAT_ID
    if not chat:
        return
    body = {"chat_id": chat, "text": text[:4000], "disable_web_page_preview": True}
    if teclado:
        body["reply_markup"] = {"inline_keyboard": teclado}
    async with httpx.AsyncClient(timeout=20) as c:
        await c.post(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage", json=body)


async def responder_boton(callback_id: str):
    async with httpx.AsyncClient(timeout=10) as c:
        await c.post(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/answerCallbackQuery",
                     json={"callback_query_id": callback_id})


def nombre_corto(emp: str) -> str:
    return emp.split(" - ", 1)[-1].strip()


def orden_prio(t):
    return {"alta": 0, "media": 1, "baja": 2}.get(str(t.get("prio", "")).lower(), 3)


def lista_empresas():
    """Empresas con tareas abiertas, las que tienen más prioridad alta primero."""
    def clave(item):
        emp, ts = item
        return (-sum(1 for t in ts if orden_prio(t) == 0), -len(ts), emp)
    return [e for e, _ in sorted(pendientes.items(), key=clave) if pendientes[e]]


def teclado_empresas(nombres):
    filas, fila = [], []
    for emp in nombres[:60]:
        idx = lista_empresas().index(emp) if emp in lista_empresas() else 0
        fila.append({"text": f"{nombre_corto(emp)[:28]} ({len(pendientes[emp])})", "callback_data": f"e:{idx}"})
        if len(fila) == 2:
            filas.append(fila)
            fila = []
    if fila:
        filas.append(fila)
    return filas


async def mostrar_empresa(emp: str, chat_id: str = ""):
    ts = sorted(pendientes.get(emp, []), key=orden_prio)
    if not ts:
        await telegram(f"{nombre_corto(emp)}: no tiene tareas abiertas.", chat_id)
        return
    partes, actual = [], f"{nombre_corto(emp)} - {len(ts)} tareas abiertas\n"
    grupo = None
    for t in ts:
        if t.get("prio") != grupo:
            grupo = t.get("prio")
            actual += f"\n[{grupo or 'Sin prioridad'}]\n"
        lim = f" (límite {t['limite']})" if t.get("limite") else ""
        linea = f"- {t['id']}: {t['tarea']}{lim}\n"
        if len(actual) + len(linea) > 3800:
            partes.append(actual)
            actual = ""
        actual += linea
    partes.append(actual)
    for i, p in enumerate(partes):
        if i == len(partes) - 1:
            idx = lista_empresas().index(emp) if emp in lista_empresas() else None
            fila = [{"text": "Otras empresas", "callback_data": "lista"}]
            if idx is not None and asientos.get(emp):
                fila.insert(0, {"text": "Excel de asientos", "callback_data": f"x:{idx}"})
            await telegram(p, chat_id, [fila])
        else:
            await telegram(p, chat_id)


def construir_excel(empresas):
    """Excel con una hoja por empresa: asientos propuestos, con totales por hoja."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    wb.remove(wb.active)
    cab = ["Asiento", "Fecha", "Cuenta", "Debe", "Haber", "Concepto", "Estado", "Tarea"]
    for emp in empresas:
        filas = asientos.get(emp, [])
        ws = wb.create_sheet(nombre_corto(emp)[:31].replace("/", "-"))
        ws.append(cab)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
        previo = None
        for r in filas:
            # r = [asiento, fecha, cuenta, debe, haber, concepto, estado, tarea]
            if previo is not None and r[0] != previo:
                ws.append([])
            previo = r[0]
            ws.append([r[0], r[1], r[2], r[3] if r[3] != "" else None, r[4] if r[4] != "" else None, r[5], r[6], r[7]])
        ultima = ws.max_row
        ws.append([])
        ws.append(["TOTAL", "", "", f"=SUM(D2:D{ultima})", f"=SUM(E2:E{ultima})", "", "", ""])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for row in ws.iter_rows(min_row=2, min_col=4, max_col=5):
            for c in row:
                c.number_format = '#,##0.00'
        for col, ancho in zip("ABCDEFGH", (11, 12, 46, 14, 14, 58, 26, 11)):
            ws.column_dimensions[col].width = ancho
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


async def enviar_excel_asientos(emp: str = "", chat_id: str = ""):
    """emp vacío = todas las empresas con asientos."""
    nombres = [emp] if emp else sorted(e for e, f in asientos.items() if f)
    nombres = [e for e in nombres if asientos.get(e)]
    if not nombres:
        await telegram("Todavía no hay asientos propuestos para eso. Se cargan cada 5 minutos desde la hoja Asientos.", chat_id)
        return
    datos = await asyncio.to_thread(construir_excel, nombres)
    etiqueta = nombre_corto(nombres[0]) if len(nombres) == 1 else "grupo"
    nombre_archivo = f"asientos_{etiqueta.lower().replace(' ', '_')}_{time.strftime('%Y%m%d')}.xlsx"
    n = sum(len(asientos[e]) for e in nombres)
    async with httpx.AsyncClient(timeout=60) as c:
        r = await c.post(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument",
                         data={"chat_id": chat_id or OWNER_CHAT_ID,
                               "caption": f"Asientos propuestos: {etiqueta} ({n} líneas). Estado: Provisional = contra 570.1, "
                                          "Propuesto = listo para contabilizar, el resto depende de la condición indicada."},
                         files={"document": (nombre_archivo, datos,
                                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    if r.status_code != 200:
        await telegram(f"No he podido enviar el Excel (estado {r.status_code}).", chat_id)


async def mostrar_lista(chat_id: str = ""):
    nombres = lista_empresas()
    if not nombres:
        await telegram("Todavía no tengo tareas cargadas. Espera unos minutos y prueba otra vez.", chat_id)
        return
    await telegram("¿De qué empresa quieres ver lo que queda por hacer?", chat_id, teclado_empresas(nombres))


def buscar_empresas(texto: str):
    t = texto.lower().strip()
    return [e for e in pendientes if t and t in e.lower()]


async def fire_routine(summary: str) -> str:
    if not (ROUTINE_FIRE_URL and ROUTINE_TOKEN):
        return "sin_configurar"
    async with httpx.AsyncClient(timeout=30) as c:
        r = await c.post(
            ROUTINE_FIRE_URL,
            headers={
                "Authorization": f"Bearer {ROUTINE_TOKEN}",
                "anthropic-version": "2023-06-01",
                "anthropic-beta": "experimental-cc-routine-2026-04-01",
                "Content-Type": "application/json",
            },
            json={"text": summary},
        )
    return f"{r.status_code}"


async def flush_when_quiet():
    """Espera a que Yuri termine de subir archivos y lanza UNA sola revisión."""
    global last_fire, worker
    while True:
        await asyncio.sleep(15)
        if not pending:
            break
        if time.time() - last_event < DEBOUNCE_SECONDS:
            continue
        if time.time() - last_fire < MIN_SECONDS_BETWEEN_RUNS:
            continue
        lines = []
        for empresa, files in pending.items():
            names = ", ".join(sorted({f["name"] for f in files}))
            lines.append(f"- {empresa}: {names}")
        pending.clear()
        text = (
            "Yuri ha subido archivos nuevos a 1_ENTRADA. Revisa solo estas empresas "
            "y deja los avisos en la hoja Avisos del TABLERO CENTRAL:\n" + "\n".join(lines)
        )
        status = await fire_routine(text)
        last_fire = time.time()
        if not status.startswith("2"):
            await telegram(f"No he podido lanzar la revisión (estado {status}). Archivos pendientes:\n" + "\n".join(lines))
        break
    worker = None


@app.get("/health")
async def health():
    return {"ok": True}


@app.post("/drive-event")
async def drive_event(request: Request, x_secret: str = Header(default="")):
    global last_event, worker
    check_secret(x_secret)
    data = await request.json()
    nuevos = 0
    for f in data.get("archivos", []):
        if seen.get(f.get("id")) == f.get("modified"):
            continue  # ya procesado: Drive avisa también al mover o abrir un archivo
        seen[f.get("id")] = f.get("modified")
        pending[data.get("empresa", "(sin empresa)")].append(f)
        nuevos += 1
    if len(seen) > 5000:
        seen.clear()
    if nuevos:
        last_event = time.time()
        if worker is None or worker.done():
            worker = asyncio.create_task(flush_when_quiet())
    return {"recibidos": nuevos}


@app.post("/pendientes")
async def recibir_pendientes(request: Request, x_secret: str = Header(default="")):
    check_secret(x_secret)
    data = await request.json()
    pendientes.clear()
    pendientes.update(data.get("empresas", {}))
    return {"empresas": len(pendientes)}


@app.post("/asientos")
async def recibir_asientos(request: Request, x_secret: str = Header(default="")):
    check_secret(x_secret)
    data = await request.json()
    asientos.clear()
    asientos.update(data.get("empresas", {}))
    return {"empresas": len(asientos), "lineas": sum(len(v) for v in asientos.values())}


@app.post("/aviso")
async def aviso(request: Request, x_secret: str = Header(default="")):
    check_secret(x_secret)
    data = await request.json()
    prio = data.get("prioridad", "")
    marca = {"Alta": "[URGENTE] ", "Media": "", "Baja": ""}.get(prio, "")
    await telegram(f"{marca}{data.get('empresa', '')}\n{data.get('texto', '')}")
    return {"ok": True}


@app.post("/telegram/{path}")
async def telegram_webhook(path: str, request: Request):
    if not hmac.compare_digest(path, WEBHOOK_PATH):
        raise HTTPException(status_code=404)
    update = await request.json()
    cb = update.get("callback_query")
    msg = update.get("message") or (cb or {}).get("message") or {}
    chat_id = str((msg.get("chat") or {}).get("id", ""))
    text = (msg.get("text") or "").strip() if not cb else ""
    if not chat_id:
        return {"ok": True}
    if not OWNER_CHAT_ID:
        # primer arranque: da a Yuri su chat id para guardarlo en Railway
        await telegram(f"Tu chat id es {chat_id}. Guárdalo en la variable OWNER_CHAT_ID de Railway.", chat_id)
        return {"ok": True}
    if chat_id != OWNER_CHAT_ID:
        return {"ok": True}               # ignora a cualquier otra persona

    if cb:
        await responder_boton(cb.get("id", ""))
        datos = cb.get("data", "")
        if datos == "lista":
            await mostrar_lista(chat_id)
        elif datos.startswith("x:"):
            nombres = lista_empresas()
            try:
                await enviar_excel_asientos(nombres[int(datos[2:])], chat_id)
            except (ValueError, IndexError):
                await mostrar_lista(chat_id)
        elif datos.startswith("e:"):
            nombres = lista_empresas()
            try:
                await mostrar_empresa(nombres[int(datos[2:])], chat_id)
            except (ValueError, IndexError):
                await mostrar_lista(chat_id)
        return {"ok": True}

    low = text.lower()
    if low.startswith("/estado"):
        await telegram(f"Pendientes de analizar: {sum(len(v) for v in pending.values())} archivos. "
                       f"Empresas con tareas abiertas: {len(lista_empresas())}.")
    elif low in ("/empresas", "/start", "empresas", "lista"):
        await mostrar_lista(chat_id)
    elif low.startswith("/asientos") or low.startswith("asientos"):
        nombre = text.split(None, 1)[1].strip() if len(text.split(None, 1)) > 1 else ""
        if not nombre or nombre.lower() in ("todos", "todas", "grupo"):
            await enviar_excel_asientos("", chat_id)
        else:
            coincide = [e for e in asientos if nombre.lower() in e.lower()]
            if len(coincide) == 1:
                await enviar_excel_asientos(coincide[0], chat_id)
            else:
                await telegram("No sé cuál es. Escribe /asientos y el nombre (ANJOBER, REAL ESTATE, PROINVER) "
                               "o solo /asientos para recibir todas." if not coincide else
                               f"Hay varias: {', '.join(nombre_corto(e) for e in coincide)}. Escribe un poco más del nombre.", chat_id)
    elif low.startswith("/cierre"):
        nombre = text[7:].strip()
        coincide = buscar_empresas(nombre)
        if len(coincide) != 1:
            await telegram("Dime la empresa: /cierre ANJOBER" if not nombre else
                           f"No sé cuál es. Coincidencias: {', '.join(nombre_corto(e) for e in coincide) or 'ninguna'}", chat_id)
        else:
            estado = await fire_routine(
                f"MODO CIERRE: {coincide[0]}. Haz la revisión de cierre completa de esta empresa (solo esta) "
                "según el apartado MODO CIERRE de las instrucciones y deja los avisos en la hoja Avisos.")
            await telegram(f"Lanzo el cierre de {nombre_corto(coincide[0])}." if estado.startswith("2")
                           else f"No he podido lanzar el cierre (estado {estado}).", chat_id)
    elif text.startswith("/"):
        await telegram("Escribe el nombre de una empresa, o /empresas para elegirla de una lista. "
                       "Con /asientos NOMBRE te envío el Excel de asientos propuestos (o /asientos para todas). "
                       "Con /cierre NOMBRE lanzo el cierre completo de esa empresa.", chat_id)
    else:
        coincide = buscar_empresas(text)
        if len(coincide) == 1:
            await mostrar_empresa(coincide[0], chat_id)
        elif coincide:
            await telegram("Hay varias, ¿cuál?", chat_id, teclado_empresas(coincide))
        else:
            await mostrar_lista(chat_id)
    return {"ok": True}
