# -*- coding: utf-8 -*-
"""
license_gate.py — проверка подписки парсера «по времени».

Использует тот же лист «Клиенты» и тот же ключ/срок, что и веб-сервис ОПВ.
Десктоп шлёт POST на веб-приложение Apps Script (doPost → apiLicense_) и
получает статус подписки. Привязка к устройству и офлайн-грейс по сроку.

БЕЗОПАСНО ПО УМОЛЧАНИЮ: если в настройках не задан license_url —
проверка полностью отключена, программа работает как раньше.
"""

import base64
import hashlib
import hmac
import json
import re
import uuid
import urllib.request
import tkinter as tk
from tkinter import messagebox
from datetime import date

# ─────────────── ЗАЩИТА / КОНФИГ (заполни ПЕРЕД сборкой .exe) ───────────────
# URL веб-приложения Apps Script (…/exec). Если вписать сюда — его нельзя
# отключить через settings.json (в отличие от пустого значения по умолчанию).
LICENSE_URL = ""

# Секрет подписи лицензии — ДОЛЖЕН совпадать с LICENSE_SECRET в Code.gs.
# На нём держится защита от подделки срока. Если пусто — подпись не проверяется
# (режим разработки). Для продажи ОБЯЗАТЕЛЬНО впиши тот же секрет, что и на сервере.
LICENSE_SECRET = ""
# ─────────────────────────────────────────────────────────────────────────────

# Цвета под тему приложения (вино/крем)
WINE = "#6B2737"; WINE_D = "#4E1C28"; BG = "#F5F0E8"; SURF = "#FFFFFF"
INK = "#1C1917"; INK2 = "#57534E"; ERR = "#991B1B"; GOLD = "#B5860D"; BORDER = "#E7E0D4"
FONT = "Segoe UI"


def get_device_id(settings, save_cb):
    d = (settings.get('device_id') or '').strip()
    if not d:
        d = 'PKB-' + uuid.uuid4().hex[:16].upper()
        settings['device_id'] = d
        save_cb(settings)
    return d


def _parse_date(s):
    s = (s or '').strip()
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$', s)
    if m:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r'^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$', s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _server_check(url, key, device_id, timeout=10):
    payload = json.dumps({'action': 'login', 'key': key, 'deviceId': device_id}).encode('utf-8')
    req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode('utf-8', 'replace')
    return json.loads(raw)


def _sign(payload_b64):
    """HMAC-подпись токена тем же секретом, что и на сервере."""
    return base64.b64encode(
        hmac.new(LICENSE_SECRET.encode('utf-8'), payload_b64.encode('utf-8'),
                 hashlib.sha256).digest()).decode('ascii')


def _verify_token(token_b64, sig_b64, device_id):
    """
    Проверяет подпись токена и привязку к устройству.
    Возвращает дату окончания (date) или None, если токен поддельный/чужой.
    Подделать срок нельзя: без LICENSE_SECRET правильную подпись не собрать.
    """
    if not token_b64 or not sig_b64:
        return None
    if not hmac.compare_digest(_sign(token_b64), str(sig_b64)):
        return None  # подпись не сходится — токен подделан
    try:
        payload = json.loads(base64.b64decode(token_b64).decode('utf-8'))
    except Exception:
        return None
    if payload.get('d') != device_id:
        return None  # токен выписан на другое устройство
    return _parse_date(payload.get('exp'))   # 'yyyy-mm-dd'


def _offline_ok(settings, device_id):
    """
    Офлайн-доступ: только по подписанному токену из прошлой успешной проверки.
    Если секрет не настроен (режим разработки) — мягкий откат на кэш даты.
    """
    if LICENSE_SECRET:
        exp = _verify_token(settings.get('license_token'),
                            settings.get('license_sig'), device_id)
        return exp is not None and date.today() <= exp
    # режим разработки (секрет пуст): старое поведение по кэшу даты
    if not settings.get('license_ok_once'):
        return False
    until = settings.get('license_until', None)
    if until in (None, ''):
        return until == ''
    d = _parse_date(until)
    return d is not None and date.today() <= d


def _ask_key(root, prefill='', info=None):
    """Модальное окно ввода ключа. Возвращает ключ (str) или None (выход)."""
    dlg = tk.Toplevel(root)
    dlg.title('Доступ к программе')
    dlg.configure(bg=BG)
    dlg.resizable(False, False)
    dlg.transient(root)
    dlg.grab_set()
    result = {'key': None}

    wrap = tk.Frame(dlg, bg=SURF, bd=0)
    wrap.pack(padx=18, pady=18, fill='both')

    tk.Label(wrap, text='🔐  Доступ по подписке', bg=SURF, fg=WINE,
             font=(FONT, 15, 'bold')).pack(anchor='w', padx=22, pady=(20, 2))
    tk.Label(wrap, text='Введите ключ доступа (тот же, что для сервиса ОПВ)',
             bg=SURF, fg=INK2, font=(FONT, 9)).pack(anchor='w', padx=22)

    if info:
        box = tk.Frame(wrap, bg='#FEF2F2', bd=0)
        box.pack(fill='x', padx=22, pady=(12, 0))
        msg = info.get('message') or 'Подписка неактивна.'
        tk.Label(box, text='⛔ ' + msg, bg='#FEF2F2', fg=ERR,
                 font=(FONT, 9, 'bold'), justify='left', wraplength=360).pack(anchor='w', padx=10, pady=8)
        price = info.get('price'); phone = info.get('kaspiPhone'); name = info.get('kaspiName')
        if phone:
            det = 'Оплата на Kaspi: {}   Получатель: {}'.format(phone, name or '')
            if price:
                det = 'Подписка {} ₸/мес.  '.format(price) + det
            tk.Label(box, text=det, bg='#FEF2F2', fg=INK2,
                     font=(FONT, 8), justify='left', wraplength=360).pack(anchor='w', padx=10, pady=(0, 8))

    ent = tk.Entry(wrap, font=(FONT, 12), bg='#FBF8F3', fg=INK,
                   relief='flat', highlightthickness=1.5, highlightbackground=BORDER,
                   highlightcolor=WINE, show='')
    ent.pack(fill='x', padx=22, pady=(14, 4), ipady=7)
    ent.insert(0, prefill or '')
    ent.focus_set()

    errlbl = tk.Label(wrap, text='', bg=SURF, fg=ERR, font=(FONT, 8))
    errlbl.pack(anchor='w', padx=22)

    def submit():
        k = ent.get().strip()
        if not k:
            errlbl.config(text='Введите ключ')
            return
        result['key'] = k
        dlg.destroy()

    def cancel():
        result['key'] = None
        dlg.destroy()

    btns = tk.Frame(wrap, bg=SURF)
    btns.pack(fill='x', padx=22, pady=(12, 20))
    tk.Button(btns, text='Войти', command=submit, bg=WINE, fg='white',
              activebackground=WINE_D, activeforeground='white', relief='flat',
              font=(FONT, 10, 'bold'), padx=22, pady=8, cursor='hand2').pack(side='left')
    tk.Button(btns, text='Выйти', command=cancel, bg=SURF, fg=INK2,
              activebackground=BG, relief='flat', font=(FONT, 10),
              padx=18, pady=8, cursor='hand2').pack(side='left', padx=(8, 0))

    ent.bind('<Return>', lambda e: submit())
    dlg.protocol('WM_DELETE_WINDOW', cancel)

    dlg.update_idletasks()
    try:
        x = root.winfo_rootx() + (root.winfo_width() - dlg.winfo_width()) // 2
        y = root.winfo_rooty() + (root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry('+{}+{}'.format(max(x, 40), max(y, 40)))
    except Exception:
        pass

    root.wait_window(dlg)
    return result['key']


def require_license(root, settings, save_cb, logger=None):
    """
    Возвращает True — доступ разрешён, False — заблокировать и выйти.
    Если license_url не задан — проверка отключена (True).
    """
    # Хардкод-URL важнее настроек: его нельзя отключить правкой settings.json.
    url = (LICENSE_URL or settings.get('license_url') or '').strip()
    if not url:
        return True  # лицензирование не настроено — как раньше

    device_id = get_device_id(settings, save_cb)
    key = (settings.get('access_key') or '').strip()
    info = None

    while True:
        if not key:
            key = _ask_key(root, prefill='', info=info)
            info = None
            if not key:
                return False
            settings['access_key'] = key
            save_cb(settings)

        try:
            r = _server_check(url, key, device_id)
        except Exception as e:
            if logger:
                logger.warning('Лицензия: нет связи (%s)', e)
            if _offline_ok(settings, device_id):
                return True  # офлайн: действует подписанный токен
            if messagebox.askretrycancel(
                    'Нет связи',
                    'Не удалось проверить подписку (нет интернета), '
                    'а сохранённый срок недоступен или истёк.\n\n'
                    'Подключитесь к интернету и повторите.'):
                continue
            return False

        if r.get('ok'):
            settings['license_ok_once'] = True
            settings['license_until'] = r.get('until', '') or ''
            settings['license_name'] = r.get('clientName', '')
            settings['access_key'] = key
            # подписанный токен для офлайн-проверки (нельзя подделать без секрета)
            settings['license_token'] = r.get('token', '')
            settings['license_sig'] = r.get('sig', '')
            save_cb(settings)
            # если секрет настроен — доверяем токену, а не «голому» ok сервера
            if LICENSE_SECRET:
                exp = _verify_token(r.get('token'), r.get('sig'), device_id)
                if exp is None or date.today() > exp:
                    info = {'message': 'Подпись лицензии недействительна. Обратитесь к администратору.'}
                    key = ''
                    settings['access_key'] = ''
                    save_cb(settings)
                    continue
            days = r.get('daysLeft')
            if logger:
                logger.info('Лицензия: OK, клиент=%s, осталось=%s',
                            r.get('clientName'), days)
            if isinstance(days, (int, float)) and days <= 7:
                messagebox.showinfo(
                    'Подписка',
                    'Подписка активна. Осталось дней: {}.'.format(int(days)))
            return True

        # не ok — сбросить ключ и предложить ввести другой
        info = r
        key = ''
        settings['access_key'] = ''
        save_cb(settings)
        if logger:
            logger.warning('Лицензия: отказ (%s) %s', r.get('reason'), r.get('message'))
