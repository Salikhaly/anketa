# -*- coding: utf-8 -*-
"""
Точка входа для Vercel (serverless Python).
Vercel сам обслуживает ASGI-приложение `app`. Основной код — в service.py
(в корне server/), сюда он подтягивается через sys.path.
"""
import os
import sys

# корень server/ (на уровень выше api/) — там service.py, parser_core.py, main.py, assets/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from service import app  # noqa: E402  (ASGI-приложение FastAPI)
