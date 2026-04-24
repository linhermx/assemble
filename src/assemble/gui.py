from __future__ import annotations

import ctypes
import math
import os
import subprocess
import threading
from ctypes import wintypes
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from assemble import __version__
from assemble.core import RunIssue, RunResult, calculate_capacity


class LinherAssembleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title(f"LINHER Assemble (v{__version__})")

        self.checklist_path = tk.StringVar(value="")
        self.inventory_path = tk.StringVar(value="")
        self.out_dir = tk.StringVar(value="")
        self.target_ovens_var = tk.IntVar(value=1)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Listo.")
        self.banner_title_var = tk.StringVar(value="Carga checklist e inventario para iniciar el análisis.")
        self.banner_info_vars = [
            tk.StringVar(value="La herramienta calcula cuántos hornos completos pueden armarse hoy."),
            tk.StringVar(value="También muestra qué material bloquea el siguiente horno completo."),
            tk.StringVar(value="Y resume los faltantes para el siguiente horno y para tu objetivo."),
        ]
        self.output_info_var = tk.StringVar(value="Aún no se ha generado ningún reporte.")
        self.last_output_dir: Path | None = None
        self.last_selected_dir = str(Path.home())

        self.metric_vars: dict[str, tk.StringVar] = {}
        self.metric_note_vars: dict[str, tk.StringVar] = {}
        self.trees: dict[str, ttk.Treeview] = {}

        self._apply_style()
        self._build_ui()
        self._reset_results()
        self._configure_window()
        self.deiconify()

    def _apply_style(self):
        style = ttk.Style(self)
        for theme in ("vista", "xpnative", "clam"):
            try:
                style.theme_use(theme)
                break
            except Exception:
                continue

        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", font=("Segoe UI", 9), foreground="#5f6b7a")
        style.configure("Value.TLabel", font=("Segoe UI", 22, "bold"))
        style.configure("Primary.TButton", padding=(14, 8))

    def _configure_window(self):
        self.update_idletasks()

        left, top, right, bottom = self._get_work_area()
        work_width = max(right - left, 900)
        work_height = max(bottom - top, 620)

        requested_width = self.root_content.winfo_reqwidth() + 40
        requested_height = self.root_content.winfo_reqheight() + 70

        available_width = max(work_width - 20, 900)
        available_height = max(work_height - 20, 620)

        width = min(max(requested_width, 1040), available_width)
        height = min(max(requested_height, 760), available_height)

        min_width = min(width, 980)
        min_height = min(height, 700)

        position_x = left + max((work_width - width) // 2, 0)
        position_y = top + max(int((work_height - height) * 0.42), 0)

        self.minsize(min_width, min_height)
        self.geometry(f"{width}x{height}+{position_x}+{position_y}")
        if requested_height > available_height or requested_width > available_width:
            try:
                self.state("zoomed")
            except Exception:
                pass

    def _get_work_area(self) -> tuple[int, int, int, int]:
        try:
            rect = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right, rect.bottom
        except Exception:
            pass
        return 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=16)
        root.grid(row=0, column=0, sticky="nsew")
        self.root_content = root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)
        root.grid_rowconfigure(3, minsize=210)

        self._build_header(root)
        self._build_controls(root)
        self._build_summary(root)
        self._build_tabs(root)

        status_wrap = ttk.Frame(root)
        status_wrap.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        status_wrap.columnconfigure(0, weight=1)

        ttk.Separator(status_wrap).grid(row=0, column=0, sticky="ew", columnspan=2, pady=(0, 6))
        ttk.Label(status_wrap, textvariable=self.status_var, style="Hint.TLabel", anchor="w").grid(
            row=1, column=0, sticky="ew"
        )
        ttk.Label(status_wrap, textvariable=self.output_info_var, style="Hint.TLabel", anchor="e").grid(
            row=1, column=1, sticky="e", padx=(12, 0)
        )

    def _build_header(self, parent: ttk.Frame):
        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Capacidad de Producción de Hornos", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=(
                "Carga el checklist estándar y el inventario estándar para saber cuántos hornos completos salen hoy, "
                "qué material bloquea el siguiente y qué faltantes existen para alcanzar un objetivo."
            ),
            style="Hint.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

    def _build_controls(self, parent: ttk.Frame):
        controls = ttk.Frame(parent)
        controls.grid(row=1, column=0, sticky="ew")
        controls.columnconfigure(0, weight=3)
        controls.columnconfigure(1, weight=2)

        files_box = ttk.LabelFrame(controls, text="Archivos de entrada", style="Section.TLabelframe")
        files_box.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        files_box.columnconfigure(1, weight=1)

        ttk.Label(files_box, text="Checklist estándar").grid(row=0, column=0, sticky="w", padx=(10, 8), pady=(12, 8))
        ttk.Entry(files_box, textvariable=self.checklist_path).grid(row=0, column=1, sticky="ew", pady=(12, 8))
        ttk.Button(files_box, text="Seleccionar...", command=self.pick_checklist).grid(
            row=0, column=2, padx=(8, 10), pady=(12, 8)
        )

        ttk.Label(files_box, text="Inventario estándar").grid(row=1, column=0, sticky="w", padx=(10, 8), pady=8)
        ttk.Entry(files_box, textvariable=self.inventory_path).grid(row=1, column=1, sticky="ew", pady=8)
        ttk.Button(files_box, text="Seleccionar...", command=self.pick_inventory).grid(
            row=1, column=2, padx=(8, 10), pady=8
        )

        ttk.Label(files_box, text="Carpeta de salida").grid(row=2, column=0, sticky="w", padx=(10, 8), pady=8)
        ttk.Entry(files_box, textvariable=self.out_dir).grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Button(files_box, text="Seleccionar...", command=self.pick_outdir).grid(
            row=2, column=2, padx=(8, 10), pady=8
        )

        ttk.Label(
            files_box,
            text="Checklist: hoja 'checklist'. Inventario: hoja 'inventario'. Puedes pegar las rutas manualmente si prefieres no usar el selector.",
            style="Hint.TLabel",
            wraplength=680,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=10, pady=(4, 12))

        action_box = ttk.LabelFrame(controls, text="Análisis", style="Section.TLabelframe")
        action_box.grid(row=0, column=1, sticky="new")
        action_box.columnconfigure(0, weight=1)

        objective_row = ttk.Frame(action_box)
        objective_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(12, 8))
        objective_row.columnconfigure(1, weight=1)

        ttk.Label(objective_row, text="Objetivo de hornos").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.target_spin = ttk.Spinbox(
            objective_row,
            from_=1,
            to=9999,
            increment=1,
            textvariable=self.target_ovens_var,
            width=8,
            justify="center",
        )
        self.target_spin.grid(row=0, column=1, sticky="w")

        ttk.Checkbutton(action_box, text="Sobrescribir si existe", variable=self.overwrite_var).grid(
            row=1, column=0, sticky="w", padx=10, pady=(0, 8)
        )

        self.run_btn = ttk.Button(action_box, text="Analizar capacidad", style="Primary.TButton", command=self.run)
        self.run_btn.grid(row=2, column=0, sticky="ew", padx=10, pady=(6, 8))

        secondary_actions = ttk.Frame(action_box)
        secondary_actions.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))
        secondary_actions.columnconfigure(0, weight=1)
        secondary_actions.columnconfigure(1, weight=1)

        self.reset_btn = ttk.Button(secondary_actions, text="Restablecer", command=self.reset_form)
        self.reset_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.open_btn = ttk.Button(
            secondary_actions,
            text="Abrir carpeta...",
            command=self.open_output,
            state="disabled",
        )
        self.open_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ttk.Label(
            action_box,
            text="Resultado esperado: hornos completos hoy, faltantes para el siguiente horno y simulación para tu objetivo.",
            style="Hint.TLabel",
            wraplength=300,
            justify="left",
        ).grid(row=4, column=0, sticky="w", padx=10, pady=(0, 12))

    def _build_summary(self, parent: ttk.Frame):
        summary = ttk.LabelFrame(parent, text="Resultado principal", style="Section.TLabelframe")
        summary.grid(row=2, column=0, sticky="ew", pady=(14, 14))
        summary.columnconfigure(0, weight=1)

        banner = tk.Frame(summary, bg="#eef4ff", bd=1, relief="solid")
        banner.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 10))
        banner.columnconfigure(0, weight=1)

        tk.Label(
            banner,
            textvariable=self.banner_title_var,
            bg="#eef4ff",
            fg="#173a63",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
            justify="left",
            wraplength=1060,
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        for index, info_var in enumerate(self.banner_info_vars, start=1):
            tk.Label(
                banner,
                textvariable=info_var,
                bg="#eef4ff",
                fg="#33506f",
                font=("Segoe UI", 10),
                anchor="w",
                justify="left",
                wraplength=1060,
            ).grid(row=index, column=0, sticky="ew", padx=14, pady=(0, 4 if index < 3 else 12))

        metrics = ttk.Frame(summary)
        metrics.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        for index in range(4):
            metrics.columnconfigure(index, weight=1)

        self._create_metric_card(
            metrics,
            row=0,
            column=0,
            key="capacity",
            title="Hornos completos hoy",
            note="Capacidad completa con inventario actual",
        )
        self._create_metric_card(
            metrics,
            row=0,
            column=1,
            key="next",
            title="Siguiente horno",
            note="Número del siguiente horno a completar",
        )
        self._create_metric_card(
            metrics,
            row=0,
            column=2,
            key="next_shortages",
            title="Faltantes del siguiente",
            note="Materiales faltantes para el siguiente horno",
        )
        self._create_metric_card(
            metrics,
            row=0,
            column=3,
            key="target_shortages",
            title="Faltantes del objetivo",
            note="Materiales faltantes para el objetivo capturado",
        )

    def _create_metric_card(self, parent: ttk.Frame, row: int, column: int, key: str, title: str, note: str):
        card = tk.Frame(parent, bg="#f7f8fb", bd=1, relief="solid")
        padx = (0, 10) if column < 3 else (0, 0)
        card.grid(row=row, column=column, sticky="nsew", padx=padx)

        tk.Label(card, text=title, bg="#f7f8fb", fg="#4d5a69", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 2)
        )

        value_var = tk.StringVar(value="--")
        note_var = tk.StringVar(value=note)
        self.metric_vars[key] = value_var
        self.metric_note_vars[key] = note_var

        tk.Label(card, textvariable=value_var, bg="#f7f8fb", fg="#111827", font=("Segoe UI", 22, "bold")).grid(
            row=1, column=0, sticky="w", padx=12, pady=(0, 4)
        )
        tk.Label(
            card,
            textvariable=note_var,
            bg="#f7f8fb",
            fg="#5f6b7a",
            font=("Segoe UI", 9),
            justify="left",
            wraplength=230,
        ).grid(row=2, column=0, sticky="w", padx=12, pady=(0, 12))

    def _build_tabs(self, parent: ttk.Frame):
        notebook = ttk.Notebook(parent)
        notebook.grid(row=3, column=0, sticky="nsew")

        self.trees["summary"] = self._create_table_tab(
            notebook,
            title="Resumen",
            columns=[
                ("Indicador", 250, "w"),
                ("Valor", 760, "w"),
            ],
        )
        self.trees["next_shortages"] = self._create_table_tab(
            notebook,
            title="Faltantes para el siguiente horno",
            columns=[
                ("Material", 360, "w"),
                ("Unidad", 90, "center"),
                ("Cantidad por horno", 120, "e"),
                ("Existencia actual", 120, "e"),
                ("Requerido total", 120, "e"),
                ("Faltante para el siguiente horno", 180, "e"),
            ],
        )
        self.trees["target_shortages"] = self._create_table_tab(
            notebook,
            title="Faltantes para el objetivo",
            columns=[
                ("Material", 360, "w"),
                ("Unidad", 90, "center"),
                ("Cantidad por horno", 120, "e"),
                ("Existencia actual", 120, "e"),
                ("Requerido total", 120, "e"),
                ("Faltante para el objetivo", 160, "e"),
            ],
        )
        self.trees["limiting"] = self._create_table_tab(
            notebook,
            title="Materiales limitantes",
            columns=[
                ("Material", 340, "w"),
                ("Unidad", 90, "center"),
                ("Cantidad por horno", 120, "e"),
                ("Existencia actual", 120, "e"),
                ("Capacidad por material", 140, "e"),
                ("Secciones", 230, "w"),
            ],
        )
        self.trees["excluded"] = self._create_table_tab(
            notebook,
            title="Excluidos",
            columns=[
                ("Material", 340, "w"),
                ("Unidad", 90, "center"),
                ("Cantidad por horno", 120, "e"),
                ("Secciones", 220, "w"),
                ("Observaciones", 260, "w"),
            ],
        )
        self.trees["issues"] = self._create_table_tab(
            notebook,
            title="Incidencias",
            columns=[
                ("Nivel", 90, "center"),
                ("Mensaje", 930, "w"),
            ],
        )

    def _create_table_tab(self, notebook: ttk.Notebook, title: str, columns: list[tuple[str, int, str]]) -> ttk.Treeview:
        frame = ttk.Frame(notebook, padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        notebook.add(frame, text=title)

        tree = ttk.Treeview(frame, columns=[name for name, _, _ in columns], show="headings", height=6)
        tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=y_scroll.set)

        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(xscrollcommand=x_scroll.set)

        for name, width, anchor in columns:
            tree.heading(name, text=name)
            tree.column(name, width=width, minwidth=70, stretch=True, anchor=anchor)

        return tree

    def pick_checklist(self):
        self._pick_file(title="Selecciona el checklist estándar", target_var=self.checklist_path)

    def pick_inventory(self):
        self._pick_file(title="Selecciona el inventario estándar", target_var=self.inventory_path)

    def pick_outdir(self):
        start_dir = self._suggested_start_dir(self.out_dir.get().strip())
        try:
            self._prepare_dialog()
            path = filedialog.askdirectory(
                parent=self,
                title="Selecciona carpeta de salida",
                initialdir=start_dir,
                mustexist=False,
            )
        except KeyboardInterrupt:
            self.status_var.set("Selección interrumpida.")
            return
        except Exception as exc:
            self.status_var.set("No se pudo abrir el selector.")
            messagebox.showerror(
                "No se pudo abrir el selector",
                f"No se pudo abrir el selector de carpeta.\n\nDetalle:\n{exc}",
            )
            return
        finally:
            self._restore_after_dialog()

        if path:
            self.out_dir.set(path)
            self.last_selected_dir = str(Path(path))
            self.status_var.set("Carpeta seleccionada.")

    def _pick_file(self, title: str, target_var: tk.StringVar):
        start_dir = self._suggested_start_dir(target_var.get().strip())
        try:
            self._prepare_dialog()
            path = filedialog.askopenfilename(
                parent=self,
                title=title,
                initialdir=start_dir,
                filetypes=[("Excel files", ("*.xlsx", "*.xlsm", "*.xls")), ("All files", "*.*")],
            )
        except KeyboardInterrupt:
            self.status_var.set("Selección interrumpida.")
            return
        except Exception as exc:
            self.status_var.set("No se pudo abrir el selector.")
            messagebox.showerror(
                "No se pudo abrir el selector",
                f"No se pudo abrir el selector de archivos.\n\nDetalle:\n{exc}",
            )
            return
        finally:
            self._restore_after_dialog()

        if path:
            target_var.set(path)
            self.last_selected_dir = str(Path(path).parent)
            self.status_var.set("Archivo seleccionado.")

    def reset_form(self):
        self.checklist_path.set("")
        self.inventory_path.set("")
        self.out_dir.set("")
        self.target_ovens_var.set(1)
        self.overwrite_var.set(False)
        self.last_output_dir = None
        self.open_btn.configure(state="disabled")
        self.status_var.set("Listo.")
        self.output_info_var.set("Aún no se ha generado ningún reporte.")
        self._reset_results()

    def _suggested_start_dir(self, current_value: str) -> str:
        if current_value:
            current_path = Path(current_value)
            if current_path.is_file():
                return str(current_path.parent)
            if current_path.is_dir():
                return str(current_path)

        if self.last_selected_dir:
            last_dir = Path(self.last_selected_dir)
            if last_dir.exists():
                return str(last_dir)

        return str(Path.home())

    def _prepare_dialog(self):
        self.update_idletasks()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)

    def _restore_after_dialog(self):
        try:
            self.attributes("-topmost", False)
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def open_output(self):
        if not self.last_output_dir:
            return
        try:
            os.startfile(str(self.last_output_dir))  # type: ignore[attr-defined]
        except Exception:
            subprocess.run(["explorer", str(self.last_output_dir)], check=False)

    def run(self):
        checklist = self.checklist_path.get().strip()
        inventory = self.inventory_path.get().strip()
        out_dir = self.out_dir.get().strip()

        if not checklist:
            messagebox.showerror("Falta checklist", "Selecciona o pega la ruta del checklist estándar.")
            return
        if not inventory:
            messagebox.showerror("Falta inventario", "Selecciona o pega la ruta del inventario estándar.")
            return
        if not out_dir:
            messagebox.showerror("Falta carpeta", "Selecciona o pega la carpeta de salida.")
            return

        try:
            target = int(self.target_ovens_var.get())
            if target <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Objetivo inválido", "El objetivo de hornos debe ser un entero positivo.")
            return

        self.run_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.last_output_dir = None
        self.status_var.set("Analizando...")
        self.output_info_var.set("Generando reporte...")
        self._set_banner(
            title="Analizando inventario y checklist...",
            lines=[
                "Estamos validando columnas obligatorias y revisando el formato.",
                "También estamos agregando materiales repetidos antes de comparar contra inventario.",
                "Enseguida calculamos capacidad actual, faltantes para el siguiente horno y faltantes para tu objetivo.",
            ],
        )
        self._set_metric("capacity", "--", "Procesando")
        self._set_metric("next", "--", "Procesando")
        self._set_metric("next_shortages", "--", "Procesando")
        self._set_metric("target_shortages", "--", "Procesando")
        self._clear_tables()

        def task():
            try:
                result = calculate_capacity(
                    checklist_path=Path(checklist),
                    inventory_path=Path(inventory),
                    output_dir=Path(out_dir),
                    target_ovens=target,
                    overwrite=bool(self.overwrite_var.get()),
                )
                self.after(0, lambda: self._on_success(result))
            except Exception as exc:
                self.after(0, lambda: self._on_error(exc))

        threading.Thread(target=task, daemon=True).start()

    def _on_success(self, result: RunResult):
        self.last_output_dir = result.output_dir
        current_capacity = "No determinado" if result.current_capacity is None else str(result.current_capacity)
        next_target = "No determinado" if result.next_oven_target is None else str(result.next_oven_target)

        self._set_metric("capacity", current_capacity, "Hornos completos con inventario actual")
        self._set_metric("next", next_target, "Número del siguiente horno a completar")
        self._set_metric(
            "next_shortages",
            str(len(result.next_shortages_frame)),
            "Materiales faltantes para el siguiente horno",
        )
        self._set_metric(
            "target_shortages",
            str(len(result.target_shortages_frame)),
            f"Materiales faltantes para llegar a {result.target_ovens} horno(s)",
        )

        self._set_banner_from_result(result)
        self.output_info_var.set(f"Reporte: {result.report_file.name} | Log: {result.log_file.name}")
        self.status_var.set("Análisis completado.")
        self.run_btn.configure(state="normal")
        self.open_btn.configure(state="normal")

        self._fill_tree(self.trees["summary"], result.summary_frame)
        self._fill_tree(self.trees["next_shortages"], result.next_shortages_frame)
        self._fill_tree(self.trees["target_shortages"], result.target_shortages_frame)
        self._fill_tree(
            self.trees["limiting"],
            result.limiting_frame[
                ["Material", "Unidad", "Cantidad por horno", "Existencia actual", "Capacidad por material", "Secciones"]
            ],
        )
        self._fill_tree(
            self.trees["excluded"],
            result.excluded_frame[["Material", "Unidad", "Cantidad por horno", "Secciones", "Observaciones"]],
        )
        self._fill_tree(self.trees["issues"], self._issues_to_frame(result.issues))

    def _on_error(self, exc: Exception):
        self.run_btn.configure(state="normal")
        self.open_btn.configure(state="disabled")
        self.status_var.set("Error.")
        self.output_info_var.set("No se generó reporte.")
        self._set_banner(
            title="El análisis no pudo completarse.",
            lines=[
                "Se detectó un problema durante el análisis.",
                str(exc),
                "Revisa la pestaña Incidencias para ver el detalle.",
            ],
        )
        self._set_metric("capacity", "--", "No disponible")
        self._set_metric("next", "--", "No disponible")
        self._set_metric("next_shortages", "--", "No disponible")
        self._set_metric("target_shortages", "--", "No disponible")
        self._fill_tree(self.trees["issues"], self._issues_to_frame([RunIssue(level="error", message=str(exc))]))
        messagebox.showerror("Error", str(exc))

    def _set_banner_from_result(self, result: RunResult):
        if result.current_capacity is None:
            self._set_banner(
                title="No se pudo determinar la capacidad completa.",
                lines=[
                    "No fue posible calcular la capacidad completa con los datos actuales.",
                    "Revisa columnas faltantes, unidades incompatibles o valores no numéricos.",
                    "Consulta la pestaña Incidencias para corregir el archivo.",
                ],
            )
            return

        if result.next_shortages_frame.empty:
            self._set_banner(
                title=f"Con el inventario actual salen {result.current_capacity} horno(s) y no hay faltantes para el siguiente.",
                lines=[
                    "Material limitante: ninguno identificado para el siguiente horno.",
                    "Disponibilidad: todos los materiales obligatorios alcanzan para el siguiente horno completo.",
                    f"Faltantes para el objetivo {result.target_ovens}: {len(result.target_shortages_frame)} material(es).",
                ],
            )
            return

        material = "Sin material limitante identificado."
        required = ""
        available = ""
        unit = ""
        if not result.limiting_frame.empty:
            row = result.limiting_frame.iloc[0]
            material = str(row.get("Material", "") or material)
            required = self._format_cell(row.get("Cantidad por horno", ""))
            available = self._format_cell(row.get("Existencia actual", ""))
            unit = str(row.get("Unidad", "") or "")

        next_number = result.current_capacity + 1
        self._set_banner(
            title=f"Con el inventario actual salen {result.current_capacity} horno(s). El horno {next_number} ya no sale completo.",
            lines=[
                f"Material limitante: {material}",
                f"Requerido por horno: {required} {unit}    Disponible: {available} {unit}".strip(),
                (
                    f"Faltantes para el siguiente horno: {len(result.next_shortages_frame)} material(es)    |    "
                    f"Faltantes para el objetivo {result.target_ovens}: {len(result.target_shortages_frame)} material(es)"
                ),
            ],
        )

    def _set_banner(self, title: str, lines: list[str]):
        self.banner_title_var.set(title)
        normalized = list(lines[:3])
        while len(normalized) < 3:
            normalized.append("")
        for info_var, line in zip(self.banner_info_vars, normalized):
            info_var.set(line)

    def _set_metric(self, key: str, value: str, note: str):
        self.metric_vars[key].set(value)
        self.metric_note_vars[key].set(note)

    def _reset_results(self):
        self._set_banner(
            title="Carga checklist e inventario para iniciar el análisis.",
            lines=[
                "El resumen mostrará cuántos hornos completos salen hoy.",
                "También verás qué material limita la producción del siguiente horno.",
                "Y se resumirán los faltantes para el siguiente horno y para tu objetivo.",
            ],
        )
        self._set_metric("capacity", "--", "Hornos completos con inventario actual")
        self._set_metric("next", "--", "Número del siguiente horno a completar")
        self._set_metric("next_shortages", "--", "Materiales faltantes para el siguiente horno")
        self._set_metric("target_shortages", "--", "Materiales faltantes para el objetivo")
        self._clear_tables()

    def _clear_tables(self):
        for tree in self.trees.values():
            for item in tree.get_children():
                tree.delete(item)

    def _fill_tree(self, tree: ttk.Treeview, frame: pd.DataFrame | None):
        for item in tree.get_children():
            tree.delete(item)

        if frame is None or frame.empty:
            return

        columns = list(tree["columns"])
        for _, row in frame.iterrows():
            values = [self._format_cell(row.get(column, "")) for column in columns]
            tree.insert("", "end", values=values)

    def _issues_to_frame(self, issues: list[RunIssue]) -> pd.DataFrame:
        return pd.DataFrame([{"Nivel": issue.level.upper(), "Mensaje": issue.message} for issue in issues])

    def _format_cell(self, value):
        if value is None:
            return ""
        if isinstance(value, float):
            if math.isnan(value):
                return ""
            if value.is_integer():
                return str(int(value))
            return f"{value:.4f}".rstrip("0").rstrip(".")
        return str(value)


def main():
    app = LinherAssembleApp()
    app.mainloop()
