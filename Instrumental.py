import os
import sys
import shutil
import subprocess
from datetime import date
from pathlib import Path
from typing import Dict, Any, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import yaml
from docxtpl import DocxTemplate

DEFAULT_FIELDS = [
    ("project_name", "Название проекта"),
    ("goal", "Цель проекта"),
    ("scope", "Область применения"),
    ("feature_list", "Ключевые функции"),
    ("ui_requirements", "Требования к интерфейсу"),
    ("deadlines", "Сроки выполнения"),
    ("customer", "Заказчик"),
    ("performers", "Исполнители"),
]


def render_docx(template_path: Path, output_docx: Path, context: Dict[str, Any]) -> None:
    tpl = DocxTemplate(str(template_path))
    ctx = dict(context)
    ctx.setdefault("today", date.today().strftime("%d.%m.%Y"))
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    tpl.render(ctx)
    tpl.save(str(output_docx))


def try_convert_to_pdf(input_docx: Path, output_pdf: Path) -> bool:
    #docx2pdf
    try:
        from docx2pdf import convert as docx2pdf_convert  # type: ignore
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        temp_out_dir = output_pdf.parent / f"__tmp_pdf_{output_pdf.stem}"
        temp_out_dir.mkdir(exist_ok=True)
        docx2pdf_convert(str(input_docx), str(temp_out_dir))
        candidate = temp_out_dir / (input_docx.stem + ".pdf")
        if candidate.exists():
            shutil.move(str(candidate), str(output_pdf))
            shutil.rmtree(temp_out_dir, ignore_errors=True)
            return True
        shutil.rmtree(temp_out_dir, ignore_errors=True)
    except Exception:
        pass

    #LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            cmd = [
                soffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_pdf.parent),
                str(input_docx),
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            produced = input_docx.with_suffix(".pdf")
            if produced.exists():
                produced.rename(output_pdf)
                return True
        except Exception:
            pass

    return False


#Вспомогательные функции для OS
def open_file(path: Path) -> None:
    if not path.exists():
        messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


def open_folder(path: Path) -> None:
    if not path.exists():
        messagebox.showerror("Ошибка", f"Папка не найдена:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])
