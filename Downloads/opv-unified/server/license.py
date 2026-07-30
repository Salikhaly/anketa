# -*- coding: utf-8 -*-
"""
Проверка доступа к серверному парсеру.
Сессия из веб-хаба проверяется через твой Apps Script (doPost, action=checkToken).
Локально можно включить DEV_ALLOW=1, чтобы обойти проверку при разработке.
"""

import os
import requests

APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
DEV_ALLOW = os.environ.get("DEV_ALLOW", "") == "1"


def check_token(token: str) -> dict:
    token = (token or "").strip()
    if DEV_ALLOW:
        return {"ok": True, "clientName": "DEV"}
    if not token:
        return {"ok": False, "message": "Нет ключа сессии — войдите в сервис заново"}
    if not APPS_SCRIPT_URL:
        return {"ok": False, "message": "Сервер не настроен (APPS_SCRIPT_URL)"}
    try:
        r = requests.post(
            APPS_SCRIPT_URL,
            json={"action": "checkToken", "token": token},
            timeout=10,
        )
        data = r.json()
        return data if isinstance(data, dict) else {"ok": False, "message": "Некорректный ответ проверки"}
    except Exception:
        return {"ok": False, "message": "Не удалось проверить доступ (нет связи с сервисом подписки)"}
