# -*- coding: utf-8 -*-
"""
parser_core.py — обёртка над main.py для веб-сервиса.
Принимает байты PDF, возвращает (превью-словарь, байты Excel).
Ничего не хранит: все временные файлы удаляются сразу.
"""

import os
import tempfile
from pathlib import Path
from typing import Optional, Dict, Any, List

import main as parser

ASSETS = Path(__file__).parent / "assets"
TEMPLATE = ASSETS / "template.xlsx"


def _contract_to_dict(c) -> Dict[str, Any]:
    return {
        "creditor": c.creditor,
        "financing_type": c.financing_type,
        "contract_amount": c.contract_amount,
        "periodic_payment": c.periodic_payment,
        "outstanding": c.outstanding,
        "current_overdue_days": c.current_overdue_days,
        "current_overdue_amount": c.current_overdue_amount,
        "max_overdue_days": c.max_overdue_days,
        "max_overdue_amount": c.max_overdue_amount,
        "end_date": c.end_date,
        "status": c.status,
        "cessionary": c.cessionary,
    }


def _write_temp(data: bytes, suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path


def parse_report(credit_pdf: bytes,
                 pension_pdf: Optional[bytes] = None,
                 enable_ocr: bool = False) -> Dict[str, Any]:
    """
    Возвращает: {
      ok, source, iin, fio, pkr, counts, monthly_load,
      active[], completed_recent[], completed_old[],
      excel_b64  # готовый файл в base64
    }
    """
    import base64

    tmp_credit = _write_temp(credit_pdf, ".pdf")
    tmp_pension = _write_temp(pension_pdf, ".pdf") if pension_pdf else None
    tmp_out = _write_temp(b"", ".xlsx")
    os.unlink(tmp_out)  # нужен только путь; write создаст заново

    try:
        cd = parser.extract_any_credit_report(
            pdf_path=Path(tmp_credit), enable_ocr=enable_ocr, ocr_lang="rus",
            poppler_path=None, tesseract_cmd=None, tessdata_dir=None,
            plumber_layout=False,
        )
        parser.validate_credit_data(cd, strict=False)

        if not cd.get("iin") and not cd.get("pkr") and not cd.get("active"):
            return {"ok": False, "error": "Не удалось прочитать отчёт. "
                    "Проверьте, что это отчёт ПКБ (1cb.kz) или ГКБ/МКБ (mkb.kz) и что PDF текстовый."}

        pension = None
        srzp = 0
        if tmp_pension:
            try:
                pension = parser.parse_pension_pdf_tables(Path(tmp_pension))
                parser.validate_pension_data(pension, strict=False)
                amts = [r.amount for r in (pension.get("rows") or [])
                        if isinstance(r, parser.PensionRow) and r.amount]
                srzp = parser.calc_pension_avg_vals(amts) if amts else 0
            except Exception:
                pension = None

        active = cd.get("active") or []
        recent = cd.get("completed_recent") or []
        old = cd.get("completed_old") or []
        revoked = cd.get("revoked") or []

        parser.write_output_legacy_excel(
            template_path=TEMPLATE, output_path=Path(tmp_out),
            pkr=cd.get("pkr"), active=active, recent=recent, old=old,
            revoked=revoked, blank_zero=True, only_loans_active=False,
        )
        if pension:
            parser.fill_anketa_from_pension(Path(tmp_out), pension, sort_rows=False)

        with open(tmp_out, "rb") as f:
            excel_bytes = f.read()

        monthly = sum((x.periodic_payment or 0) for x in active)

        return {
            "ok": True,
            "source": cd.get("source"),
            "iin": cd.get("iin"),
            "fio": (pension.get("fio") if pension else None) or cd.get("fio"),
            "pkr": cd.get("pkr"),
            "srzp": srzp,
            "monthly_load": monthly,
            "counts": {
                "active": len(active),
                "completed_recent": len(recent),
                "completed_old": len(old),
            },
            "active": [_contract_to_dict(c) for c in active],
            "completed_recent": [_contract_to_dict(c) for c in recent],
            "completed_old": [_contract_to_dict(c) for c in old],
            "excel_b64": base64.b64encode(excel_bytes).decode("ascii"),
        }
    finally:
        for p in (tmp_credit, tmp_pension, tmp_out):
            if p and os.path.exists(p):
                try:
                    os.unlink(p)
                except OSError:
                    pass
