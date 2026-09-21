"""
Integración con Outlook Calendar vía Microsoft Graph.

Usa el flujo "device code" de MSAL: no hace falta client secret ni redirect
URI. El usuario autoriza una vez desde su navegador y el token (con refresh)
se guarda cifrado-en-disco en el volumen persistente de Railway, así que
no hay que volver a autorizar tras cada despliegue.
"""
import os

import msal
import requests

CLIENT_ID = os.environ.get("GEMAP_MS_CLIENT_ID", "")
AUTHORITY = os.environ.get("GEMAP_MS_AUTHORITY", "https://login.microsoftonline.com/common")
SCOPES = ["Calendars.ReadWrite"]
CACHE_PATH = os.environ.get("GEMAP_MS_CACHE_PATH", "/app/data/msal_cache.bin")


def configured() -> bool:
    return bool(CLIENT_ID)


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r") as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())


def _app():
    cache = _load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    return app, cache


def is_connected() -> bool:
    if not configured():
        return False
    app, _ = _app()
    return len(app.get_accounts()) > 0


def connected_account_name() -> str | None:
    app, _ = _app()
    accounts = app.get_accounts()
    return accounts[0].get("username") if accounts else None


def start_device_flow() -> dict:
    app, _ = _app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(flow.get("error_description", "No se pudo iniciar la conexión con Microsoft."))
    return flow


def try_complete_device_flow(flow: dict):
    """Un único intento de canje (no bloquea en bucle). Devuelve (estado, detalle).
    estado: 'ok' | 'pending' | 'error'
    """
    app, cache = _app()
    result = app.acquire_token_by_device_flow(flow, exit_condition=lambda f: True)
    _save_cache(cache)
    if "access_token" in result:
        return "ok", None
    err = result.get("error")
    if err in ("authorization_pending", "slow_down"):
        return "pending", None
    return "error", result.get("error_description", err or "Error desconocido")


def disconnect():
    app, cache = _app()
    for acc in app.get_accounts():
        app.remove_account(acc)
    _save_cache(cache)


def _get_token():
    app, cache = _app()
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)
    if result and "access_token" in result:
        return result["access_token"]
    return None


def create_event(subject: str, date_str: str, body: str = "", location: str = "",
                  start_time: str = "09:00", duration_minutes: int = 30):
    """Crea un evento de un día en el calendario del usuario conectado.
    date_str: 'YYYY-MM-DD'. Devuelve (ok: bool, event_id_o_error: str).
    """
    token = _get_token()
    if not token:
        return False, "No conectado"

    from datetime import datetime, timedelta
    start_dt = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
    end_dt = start_dt + timedelta(minutes=duration_minutes)

    payload = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "start": {"dateTime": start_dt.isoformat(), "timeZone": "Romance Standard Time"},
        "end": {"dateTime": end_dt.isoformat(), "timeZone": "Romance Standard Time"},
        "location": {"displayName": location},
    }
    try:
        resp = requests.post(
            "https://graph.microsoft.com/v1.0/me/events",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload, timeout=15,
        )
    except requests.RequestException as e:
        return False, str(e)

    if resp.status_code in (200, 201):
        return True, resp.json().get("id")
    return False, f"{resp.status_code}: {resp.text[:200]}"


def delete_event(event_id: str):
    token = _get_token()
    if not token or not event_id:
        return
    try:
        requests.delete(
            f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
            headers={"Authorization": f"Bearer {token}"}, timeout=15,
        )
    except requests.RequestException:
        pass
