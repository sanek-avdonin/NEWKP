from __future__ import annotations

import threading
import random
from decimal import Decimal
from datetime import datetime
from pathlib import Path
from typing import List
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from kp_generator.config import APP_NAME, COMPANIES_JSON, DEFAULT_OUTPUT_DIR, ROUNDING_STEPS
from kp_generator.logger import setup_file_logger
from kp_generator.company_store import CompanyStore
from kp_generator.models import VariantSettings
from kp_generator.pricing import apply_pricing
from kp_generator.extract.excel_reader import read_items_from_excel
from kp_generator.extract.pdf_reader import extract_text_from_pdf
from kp_generator.extract.table_parser import parse_items_from_text
from kp_generator.render.excel_template import render_kp
from kp_generator.render.docx_template import render_kp_docx


class AppGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("900x650")

        self.logger = setup_file_logger(DEFAULT_OUTPUT_DIR)
        self.store = CompanyStore(COMPANIES_JSON)
        try:
            self.companies = self.store.load()
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Ошибка загрузки", str(e))
            self.companies = []

        self.input_path: str | None = None
        self.template_excel_path: str | None = None
        self.template_docx_path: str | None = None

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        top = ttk.LabelFrame(frm, text="Исходный файл", padding=10)
        top.pack(fill=tk.X)

        self.lbl_file = ttk.Label(top, text="Файл не выбран")
        self.lbl_file.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(top, text="Выбрать PDF/Excel", command=self.on_choose_file).pack(side=tk.RIGHT)

        tpl = ttk.LabelFrame(frm, text="Шаблон DOCX (опционально)", padding=10)
        tpl.pack(fill=tk.X, pady=5)

        self.lbl_tpl = ttk.Label(tpl, text="DOCX-шаблон не выбран (если не выбрать — будет Excel/дефолт)")
        self.lbl_tpl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(tpl, text="Выбрать DOCX", command=self.on_choose_docx_template).pack(side=tk.RIGHT)

        variants_box = ttk.LabelFrame(frm, text="Варианты КП (2–3)", padding=10)
        variants_box.pack(fill=tk.X, pady=10)

        self.variant_vars = []

        for i in range(3):
            vf = ttk.Frame(variants_box)
            vf.pack(fill=tk.X, pady=5)
            self._build_variant_row(vf, i + 1)

        actions = ttk.Frame(frm)
        actions.pack(fill=tk.X, pady=5)

        ttk.Label(actions, text="Папка вывода:").pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        ttk.Entry(actions, textvariable=self.out_dir_var, width=60).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions, text="...", command=self.on_choose_out_dir).pack(side=tk.LEFT)

        self.btn_generate = ttk.Button(actions, text="Сгенерировать", command=self.on_generate)
        self.btn_generate.pack(side=tk.RIGHT)

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=5)

        logbox = ttk.LabelFrame(frm, text="Лог", padding=10)
        logbox.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(logbox, height=18, wrap="word")
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        self._log("Готово. Выберите исходный файл и (при необходимости) DOCX-шаблон.")

    def _build_variant_row(self, parent: ttk.Frame, idx: int):
        enabled = tk.BooleanVar(value=(idx <= 2))
        company = tk.StringVar(value=self.companies[0].id if self.companies else "")
        percent = tk.StringVar(value="1.5")
        fixed = tk.StringVar(value="0")
        spread = tk.StringVar(value="50")
        rounding = tk.StringVar(value="10")

        self.variant_vars.append((enabled, company, percent, fixed, spread, rounding))

        ttk.Checkbutton(parent, text=f"Вариант {idx}", variable=enabled).pack(side=tk.LEFT, padx=5)

        ttk.Label(parent, text="Компания:").pack(side=tk.LEFT)
        cb = ttk.Combobox(parent, state="readonly", width=28, values=[f"{c.id} — {c.name}" for c in self.companies])
        cb.pack(side=tk.LEFT, padx=5)
        if self.companies:
            cb.current(0)

        def on_cb_select(_):
            cur = cb.get().split("—", 1)[0].strip()
            company.set(cur)

        cb.bind("<<ComboboxSelected>>", on_cb_select)

        ttk.Label(parent, text="%:").pack(side=tk.LEFT)
        ttk.Entry(parent, textvariable=percent, width=6).pack(side=tk.LEFT, padx=2)

        ttk.Label(parent, text="+руб:").pack(side=tk.LEFT)
        ttk.Entry(parent, textvariable=fixed, width=8).pack(side=tk.LEFT, padx=2)

        ttk.Label(parent, text="±руб:").pack(side=tk.LEFT)
        ttk.Entry(parent, textvariable=spread, width=8).pack(side=tk.LEFT, padx=2)

        ttk.Label(parent, text="Округл.:").pack(side=tk.LEFT)
        rcb = ttk.Combobox(parent, state="readonly", width=6, values=ROUNDING_STEPS, textvariable=rounding)
        rcb.pack(side=tk.LEFT, padx=2)
        rcb.set("10")

    def _log(self, msg: str):
        self.logger.info(msg)
        # Обновление текстового лога в главном потоке (безопасно при вызове из worker)
        self.root.after(0, lambda m=msg: self._log_ui(m))

    def _log_ui(self, msg: str):
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)

    def on_choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите исходное КП",
            filetypes=[("PDF или Excel", "*.pdf *.xlsx *.xls"), ("PDF", "*.pdf"), ("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return
        self.input_path = path
        self.lbl_file.config(text=path)
        self._log(f"Выбран файл: {path}")

        if path.lower().endswith((".xlsx", ".xls")):
            self.template_excel_path = path
        else:
            self.template_excel_path = None

    def on_choose_docx_template(self):
        path = filedialog.askopenfilename(
            title="Выберите DOCX-шаблон",
            filetypes=[("Word DOCX", "*.docx"), ("Все файлы", "*.*")]
        )
        if not path:
            return
        if path.lower().endswith(".doc"):
            messagebox.showwarning(
                "Формат .doc не поддерживается",
                "Сохраните файл в Word как .docx и выберите его.",
            )
            return
        self.template_docx_path = path
        self.lbl_tpl.config(text=path)
        self._log(f"Выбран DOCX-шаблон: {path}")

    def on_choose_out_dir(self):
        d = filedialog.askdirectory(title="Выберите папку вывода")
        if d:
            self.out_dir_var.set(d)

    def _parse_decimal(self, s: str) -> Decimal:
        return Decimal(str(s).replace(",", ".").strip())

    def _collect_variant_settings(self) -> List[VariantSettings]:
        enabled_variants: List[VariantSettings] = []

        for idx, (enabled, company_id, percent, fixed, spread, rounding) in enumerate(self.variant_vars, start=1):
            if not enabled.get():
                continue

            try:
                p = self._parse_decimal(percent.get())
                f = self._parse_decimal(fixed.get())
                s = self._parse_decimal(spread.get())
                r = self._parse_decimal(rounding.get())
            except Exception as e:
                raise ValueError(f"Вариант {idx}: неверный формат числа (процент/надбавка/разброс/округление).") from e

            if p < 0:
                raise ValueError(f"Вариант {idx}: процент не может быть отрицательным.")
            if f < 0:
                raise ValueError(f"Вариант {idx}: надбавка (руб) не может быть отрицательной.")
            if s < 0:
                raise ValueError(f"Вариант {idx}: разброс (±руб) не может быть отрицательным.")
            if str(r) not in ROUNDING_STEPS and r not in [Decimal(x) for x in ROUNDING_STEPS]:
                raise ValueError(f"Вариант {idx}: шаг округления должен быть одним из {', '.join(ROUNDING_STEPS)}.")

            enabled_variants.append(VariantSettings(
                company_id=company_id.get(),
                percent_up=p,
                fixed_add=f,
                random_spread=s,
                rounding_step=r,
            ))

        if len(enabled_variants) < 2:
            raise ValueError("Нужно включить минимум 2 варианта КП.")

        return enabled_variants

    def on_generate(self):
        if not self.input_path:
            messagebox.showwarning("Нет файла", "Сначала выберите исходный PDF/Excel.")
            return
        if not self.companies:
            messagebox.showwarning("Нет компаний", "Не удалось загрузить компании. Проверьте assets/companies.json.")
            return

        out_dir = Path(self.out_dir_var.get()).expanduser()
        out_dir.mkdir(parents=True, exist_ok=True)

        try:
            _ = self._collect_variant_settings()
        except ValueError as e:
            self._log(f"Ошибка в настройках вариантов: {e}")
            messagebox.showerror("Ошибка", str(e))
            return

        self.btn_generate.config(state=tk.DISABLED)
        self.progress.start(10)

        th = threading.Thread(target=self._generate_worker, args=(str(out_dir),), daemon=True)
        th.start()

    def _generate_worker(self, out_dir: str):
        try:
            enabled_variants = self._collect_variant_settings()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if self.template_docx_path:
                self._log("Используется DOCX-шаблон, Excel-шаблон не применяется.")

            self._log("Извлечение позиций...")

            items = None
            if self.input_path.lower().endswith((".xlsx", ".xls")):
                if self.input_path.lower().endswith(".xls"):
                    raise ValueError("Поддержка .xls отключена. Сохраните файл как .xlsx.")
                items, sheet = read_items_from_excel(self.input_path)
                self._log(f"Извлечено позиций из Excel: {len(items)} (лист: {sheet})")
            else:
                text = extract_text_from_pdf(self.input_path, ocr_lang="rus")
                items = parse_items_from_text(text)
                self._log(f"Извлечено позиций из PDF: {len(items)}")

            rng = random.Random()
            rng.seed()

            for i, st in enumerate(enabled_variants, start=1):
                company = self.store.get(st.company_id)
                self._log(
                    f"Вариант {i}: компания '{company.name}', "
                    f"настройки: %={st.percent_up}, +руб={st.fixed_add}, ±={st.random_spread}, округл={st.rounding_step}"
                )

                new_items = apply_pricing(items, st, rng)

                safe_name = "".join(ch for ch in company.name if ch.isalnum() or ch in " _-").strip().replace(" ", "_")
                base_out = Path(out_dir) / f"КП_{safe_name}_{timestamp}_v{i}"

                if self.template_docx_path:
                    out_path = str(base_out.with_suffix(".docx"))
                    render_kp_docx(
                        template_path=self.template_docx_path,
                        company=company,
                        items=new_items,
                        output_path=out_path
                    )
                else:
                    out_path = str(base_out.with_suffix(".xlsx"))
                    render_kp(
                        template_path=self.template_excel_path,
                        company=company,
                        items=new_items,
                        output_path=out_path
                    )

                self._log(f"Сохранено: {out_path}")

            self._log("Готово ✅")

        except ValueError as e:
            self._log(f"Ошибка извлечения данных: {e}")
            msg = (
                "Не удалось извлечь товары из файла.\n"
                "Проверьте, что файл содержит таблицу с заголовками: Наименование, Количество, Цена и т.п.\n\n"
                f"Техническая информация: {e}"
            )
            self.root.after(0, lambda m=msg: messagebox.showerror("Ошибка", m))
        except Exception as e:
            self._log(f"Ошибка: {e}")
            err_msg = str(e)
            self.root.after(0, lambda e=err_msg: messagebox.showerror("Ошибка", e))
        finally:
            self.root.after(0, self._ui_finish)

    def _ui_finish(self):
        self.progress.stop()
        self.btn_generate.config(state=tk.NORMAL)
