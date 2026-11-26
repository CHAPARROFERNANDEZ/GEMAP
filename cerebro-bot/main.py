"""Bot SEGUNDO CEREBRO (GEMAP).

Recibe avisos de Google Drive (Apps Script), lanza la rutina de revisión de Claude
y reenvía los avisos a Telegram. Solo responde al chat de Yuri.
"""
import asyncio
import hmac
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
last_event = 0.0
last_fire = 0.0
worker = None


def check_secret(value: str):
    if not value or not hmac.compare_digest(value, SHARED_SECRET):
        raise HTTPException(status_code=401)


async def telegram(text: str, chat_id: str = ""):
    chat = chat_id or OWNER_CHAT_ID
    if not chat:
        return
    async with httpx.AsyncClient(timeout=20) as c:
        await c.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={"chat_id": chat, "text": text[:4000], "disable_web_page_preview": True},
        )


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
        await telegram("Archivos recibidos. Analizo:\n" + "\n".join(lines) + f"\n(estado {status})")
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
    for f in data.get("archivos", []):
        pending[data.get("empresa", "(sin empresa)")].append(f)
    last_event = time.time()
    if worker is None or worker.done():
        worker = asyncio.create_task(flush_when_quiet())
    return {"recibidos": len(data.get("archivos", []))}


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
    msg = update.get("message") or {}
    chat_id = str((msg.get("chat") or {}).get("id", ""))
    text = (msg.get("text") or "").strip()
    if not chat_id:
        return {"ok": True}
    if not OWNER_CHAT_ID:
        # primer arranque: da a Yuri su chat id para guardarlo en Railway
        await telegram(f"Tu chat id es {chat_id}. Guárdalo en la variable OWNER_CHAT_ID de Railway.", chat_id)
        return {"ok": True}
    if chat_id != OWNER_CHAT_ID:
        return {"ok": True}               # ignora a cualquier otra persona
    if text.startswith("/estado"):
        await telegram(f"Pendientes de analizar: {sum(len(v) for v in pending.values())} archivos.")
    else:
        await telegram("Recibido. De momento solo envío avisos; las órdenes (mover reuniones, añadir tareas) llegan en la fase 2.")
    return {"ok": True}
