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

#основа для Tkinter
class TzApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DEAL Генератор технических заданий")
        self.geometry("900x600")
        self.minsize(900, 600)

        # состояние приложения
        self.template_path: Optional[Path] = None
        self.output_dir: Path = Path("output")
        self.base_name: str = "tz"
        self.form_data: Dict[str, str] = {k: "" for k, _ in DEFAULT_FIELDS}
        self.generated_docx: Optional[Path] = None
        self.generated_pdf: Optional[Path] = None

        # контейнер для экранов
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        # создаём кадры-экраны
        self.frames = {}
        for FrameCls in (MainMenuFrame, FormFrame, PreviewFrame, ResultFrame):
            frame = FrameCls(parent=container, controller=self)
            frame.grid(row=0, column=0, sticky="nsew")
            self.frames[FrameCls.__name__] = frame

        self.show_frame("MainMenuFrame")

    def show_frame(self, name: str):
        frame = self.frames[name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()  # type: ignore[call-arg]

#фреймы
class MainMenuFrame(ttk.Frame):
    def __init__(self, parent, controller: TzApp):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # Заголовок
        title = ttk.Label(
            self,
            text="DEAL Генератор технических заданий",
            font=("Arial", 24, "bold"),
            anchor="center",
        )
        title.grid(row=0, column=0, pady=(40, 20))

        # центральная панель с кнопками
        center = ttk.Frame(self)
        center.grid(row=1, column=0)
        for i in range(3):
            center.rowconfigure(i, pad=10)

        btn_new = ttk.Button(center, text="Создать новое ТЗ", command=self.on_new_tz)
        btn_new.grid(row=0, column=0, ipady=10, ipadx=40, pady=10)

        btn_template = ttk.Button(
            center, text="Загрузить шаблон DOCX", command=self.on_load_template
        )
        btn_template.grid(row=1, column=0, ipady=10, ipadx=40, pady=10)

        # подпись о текущем шаблоне
        self.template_label = ttk.Label(
            self,
            text="Текущий шаблон: [по умолчанию: templates/ts_template.docx]",
            font=("Arial", 10),
        )
        self.template_label.grid(row=2, column=0, pady=(40, 20))

        # кнопка Выход в правом нижнем углу
        bottom = ttk.Frame(self)
        bottom.grid(row=3, column=0, sticky="se", padx=20, pady=20)
        btn_exit = ttk.Button(bottom, text="Выход", command=self.controller.destroy)
        btn_exit.pack()

    def on_show(self):
        # обновить надпись о шаблоне
        if self.controller.template_path:
            text = f"Текущий шаблон: {self.controller.template_path}"
        else:
            text = "Текущий шаблон: [по умолчанию: templates/ts_template.docx]"
        self.template_label.config(text=text)

    def on_load_template(self):
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон DOCX",
            filetypes=[("DOCX файлы", "*.docx")],
        )
        if file_path:
            self.controller.template_path = Path(file_path)
            self.on_show()

    def reset_state(self):
        self.form_data = {k: "" for k, _ in DEFAULT_FIELDS}
        self.template_path = None
        self.generated_docx = None
        self.generated_pdf = None

    def on_new_tz(self):
        self.controller.reset_state()
        self.controller.show_frame("FormFrame")

name__ == "__main__":
    app = TzApp()
    app.mainloop()