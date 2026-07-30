# -*- coding: utf-8 -*-
"""
FastAPI-сервис парсера ПКБ/ГКБ.
POST /parse  — multipart: token, file (кредитный PDF), pension (опц.)
             → JSON с превью договоров + Excel в base64.
GET  /health — проверка живости.
Ничего не хранит: файлы обрабатываются в памяти/во временных и сразу удаляются.
"""

import logging
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

import parser_core
import license as lic

logging.basicConfig(level=logging.INFO)
MAX_BYTES = 15 * 1024 * 1024  # 15 МБ

app = FastAPI(title="Парсер ПКБ/ГКБ")

# CORS: страница сервиса (Apps Script) обращается с другого домена
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # можно сузить до домена Apps Script
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    return {"ok": True}


@app.get("/debug-env")
def debug_env():
    """Временный диагностический эндпоинт: не раскрывает сам URL, только факт наличия."""
    import os
    val = os.environ.get("APPS_SCRIPT_URL", "")
    return {
        "apps_script_url_is_set": bool(val.strip()),
        "length": len(val),
        "starts_with_https": val.strip().startswith("https://") if val else False,
        "ends_with_exec": val.strip().endswith("/exec") if val else False,
    }


@app.post("/parse")
async def parse(
    token: str = Form(""),
    file: UploadFile = File(...),
    pension: UploadFile = File(None),
):
    # 1) доступ по подписке (сессия из веб-хаба)
    auth = lic.check_token(token)
    if not auth.get("ok"):
        raise HTTPException(status_code=401, detail=auth.get("message", "Доступ запрещён"))

    # 2) читаем файлы
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Пустой файл")
    if len(data) > MAX_BYTES:
        raise HTTPException(status_code=413, detail="Файл больше 15 МБ")
    pdata = await pension.read() if pension is not None else None
    if pdata and len(pdata) > MAX_BYTES:
        raise HTTPException(status_code=413, detail="Пенсионный файл больше 15 МБ")

    # 3) парсинг (ничего не сохраняем)
    try:
        result = parser_core.parse_report(data, pdata)
    except Exception as e:
        logging.exception("parse error")
        raise HTTPException(status_code=500, detail="Ошибка обработки файла")

    if not result.get("ok"):
        return JSONResponse(status_code=422, content=result)

    result["client"] = auth.get("clientName")
    return result
