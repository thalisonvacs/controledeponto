from __future__ import annotations

import csv
import sqlite3
import tempfile
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
from urllib import error, request
import json
import re

APP_NAME = "NexusLab Ponto Pro"
DB_PATH = Path("nexuslab_ponto.db")


@dataclass
class Employee:
    id: int
    name: str
    registration: str
    department: str
    pis: str


class Database:
    def __init__(self, db_path: Path) -> None:
        self.conn = sqlite3.connect(db_path)
        self.conn.execute("PRAGMA foreign_keys = ON;")
        self._init_schema()

    def _init_schema(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                registration TEXT NOT NULL UNIQUE,
                department TEXT NOT NULL,
                pis TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS punches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                punch_type TEXT NOT NULL,
                punch_time TEXT NOT NULL,
                origin TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
        self._migrate_employees_table()
        self.conn.commit()

    def _migrate_employees_table(self) -> None:
        cols = {row[1] for row in self.conn.execute("PRAGMA table_info(employees)").fetchall()}
        if "pis" not in cols:
            self.conn.execute("ALTER TABLE employees ADD COLUMN pis TEXT NOT NULL DEFAULT ''")

    def add_employee(self, name: str, registration: str, department: str, pis: str) -> None:
        self.conn.execute(
            """
            INSERT INTO employees (name, registration, department, pis, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (name.strip(), registration.strip(), department.strip(), pis.strip(), datetime.now().isoformat()),
        )
        self.conn.commit()

    def update_employee(self, employee_id: int, name: str, registration: str, department: str, pis: str) -> None:
        self.conn.execute(
            """
            UPDATE employees
            SET name = ?, registration = ?, department = ?, pis = ?
            WHERE id = ?
            """,
            (name.strip(), registration.strip(), department.strip(), pis.strip(), employee_id),
        )
        self.conn.commit()

    def delete_employee(self, employee_id: int) -> None:
        self.conn.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
        self.conn.commit()

    def list_employees(self) -> list[Employee]:
        cursor = self.conn.execute(
            "SELECT id, name, registration, department, pis FROM employees ORDER BY name ASC"
        )
        return [Employee(*row) for row in cursor.fetchall()]

    def add_punch(self, employee_id: int, punch_type: str, origin: str, punch_time: str | None = None) -> None:
        if punch_time is None:
            punch_time = datetime.now().isoformat(timespec="seconds")
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            """
            INSERT INTO punches (employee_id, punch_type, punch_time, origin, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (employee_id, punch_type, punch_time, origin, now),
        )
        self.conn.commit()

    def list_punches(self) -> list[tuple]:
        cursor = self.conn.execute(
            """
            SELECT p.id, e.name, e.registration, p.punch_type, p.punch_time, p.origin
            FROM punches p
            JOIN employees e ON e.id = p.employee_id
            ORDER BY p.punch_time DESC
            """
        )
        return cursor.fetchall()

    def monthly_summary(self, month: str) -> list[tuple]:
        cursor = self.conn.execute(
            """
            SELECT e.name, e.registration, e.department, COUNT(p.id) as total
            FROM employees e
            LEFT JOIN punches p ON p.employee_id = e.id AND strftime('%Y-%m', p.punch_time) = ?
            GROUP BY e.id
            ORDER BY e.name
            """,
            (month,),
        )
        return cursor.fetchall()

    def punches_for_employee_month(self, employee_id: int, month: str) -> list[tuple]:
        cursor = self.conn.execute(
            """
            SELECT punch_type, punch_time, origin
            FROM punches
            WHERE employee_id = ?
              AND strftime('%Y-%m', punch_time) = ?
            ORDER BY punch_time ASC
            """,
            (employee_id, month),
        )
        return cursor.fetchall()

    def upsert_setting(self, key: str, value: str) -> None:
        self.conn.execute(
            """
            INSERT INTO settings(key, value)
            VALUES(?, ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value
            """,
            (key, value),
        )
        self.conn.commit()

    def get_setting(self, key: str, default: str = "") -> str:
        cursor = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row[0] if row else default


class App(tk.Tk):
    def __init__(self, db: Database) -> None:
        super().__init__()
        self.db = db
        self.title(f"{APP_NAME} • Controle de Ponto para Control iD")
        self.geometry("1200x760")
        self.configure(bg="#0f172a")

        self.selected_employee_id: int | None = None

        self._setup_style()

        header = tk.Frame(self, bg="#0f172a", pady=16)
        header.pack(fill="x")
        tk.Label(
            header,
            text=APP_NAME,
            fg="#38bdf8",
            bg="#0f172a",
            font=("Segoe UI", 24, "bold"),
        ).pack()
        tk.Label(
            header,
            text="Gestão completa de ponto eletrônico • by NexusLab",
            fg="#cbd5e1",
            bg="#0f172a",
            font=("Segoe UI", 10),
        ).pack()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        self.tab_employees = ttk.Frame(notebook)
        self.tab_punch = ttk.Frame(notebook)
        self.tab_reports = ttk.Frame(notebook)
        self.tab_controlid = ttk.Frame(notebook)
        self.tab_settings = ttk.Frame(notebook)

        notebook.add(self.tab_employees, text="Colaboradores")
        notebook.add(self.tab_punch, text="Lançamento")
        notebook.add(self.tab_reports, text="Relatórios")
        notebook.add(self.tab_controlid, text="Integração Control iD")
        notebook.add(self.tab_settings, text="Configurações")

        self._build_employees_tab()
        self._build_punch_tab()
        self._build_reports_tab()
        self._build_controlid_tab()
        self._build_settings_tab()

        self.refresh_employees()
        self.refresh_punches()
        self.refresh_summary()

    def _setup_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#111827")
        style.configure("TNotebook", background="#0f172a", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1e293b", foreground="#f8fafc", padding=(12, 8))
        style.map("TNotebook.Tab", background=[("selected", "#334155")])
        style.configure("TLabel", background="#111827", foreground="#e2e8f0", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8)
        style.configure("Treeview", background="#1f2937", fieldbackground="#1f2937", foreground="#e5e7eb")
        style.configure("Treeview.Heading", background="#334155", foreground="#f8fafc")

    def _build_employees_tab(self) -> None:
        frame = ttk.Frame(self.tab_employees, padding=16)
        frame.pack(fill="both", expand=True)

        form = ttk.Frame(frame)
        form.pack(fill="x", pady=(0, 14))

        ttk.Label(form, text="Nome").grid(row=0, column=0, sticky="w")
        ttk.Label(form, text="Matrícula").grid(row=0, column=1, sticky="w")
        ttk.Label(form, text="Departamento").grid(row=0, column=2, sticky="w")
        ttk.Label(form, text="PIS").grid(row=0, column=3, sticky="w")

        self.name_var = tk.StringVar()
        self.reg_var = tk.StringVar()
        self.dep_var = tk.StringVar()
        self.pis_var = tk.StringVar()

        ttk.Entry(form, textvariable=self.name_var, width=28).grid(row=1, column=0, padx=(0, 8), sticky="we")
        ttk.Entry(form, textvariable=self.reg_var, width=16).grid(row=1, column=1, padx=(0, 8), sticky="we")
        ttk.Entry(form, textvariable=self.dep_var, width=20).grid(row=1, column=2, padx=(0, 8), sticky="we")
        ttk.Entry(form, textvariable=self.pis_var, width=18).grid(row=1, column=3, padx=(0, 8), sticky="we")
        ttk.Button(form, text="Adicionar", command=self.create_employee).grid(row=1, column=4, padx=(0, 8))
        ttk.Button(form, text="Salvar edição", command=self.update_employee).grid(row=1, column=5, padx=(0, 8))
        ttk.Button(form, text="Excluir", command=self.delete_employee).grid(row=1, column=6)

        self.employee_tree = ttk.Treeview(frame, columns=("id", "name", "reg", "dep", "pis"), show="headings", height=18)
        for col, title, width in [
            ("id", "ID", 60),
            ("name", "Nome", 300),
            ("reg", "Matrícula", 140),
            ("dep", "Departamento", 180),
            ("pis", "PIS", 160),
        ]:
            self.employee_tree.heading(col, text=title)
            self.employee_tree.column(col, width=width, anchor="w")
        self.employee_tree.pack(fill="both", expand=True)
        self.employee_tree.bind("<<TreeviewSelect>>", self._on_employee_select)

    def _build_punch_tab(self) -> None:
        frame = ttk.Frame(self.tab_punch, padding=16)
        frame.pack(fill="both", expand=True)

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(0, 10))

        ttk.Label(toolbar, text="Colaborador").grid(row=0, column=0, sticky="w")
        ttk.Label(toolbar, text="Tipo").grid(row=0, column=1, sticky="w")
        ttk.Label(toolbar, text="Origem").grid(row=0, column=2, sticky="w")

        self.punch_employee_var = tk.StringVar()
        self.punch_type_var = tk.StringVar(value="Entrada")
        self.punch_origin_var = tk.StringVar(value="Manual")

        self.employee_combo = ttk.Combobox(toolbar, textvariable=self.punch_employee_var, state="readonly", width=42)
        self.employee_combo.grid(row=1, column=0, padx=(0, 10), sticky="we")

        ttk.Combobox(
            toolbar,
            textvariable=self.punch_type_var,
            state="readonly",
            values=["Entrada", "Saída para almoço", "Retorno do almoço", "Saída"],
            width=24,
        ).grid(row=1, column=1, padx=(0, 10), sticky="we")

        ttk.Combobox(
            toolbar,
            textvariable=self.punch_origin_var,
            state="readonly",
            values=["Manual", "Control iD", "Importação AFD", "Importação"],
            width=20,
        ).grid(row=1, column=2, padx=(0, 10), sticky="we")

        ttk.Button(toolbar, text="Registrar ponto agora", command=self.create_punch).grid(row=1, column=3, padx=(0, 8))
        ttk.Button(toolbar, text="Importar AFD", command=self.import_afd).grid(row=1, column=4)

        self.punch_tree = ttk.Treeview(
            frame,
            columns=("id", "name", "reg", "type", "time", "origin"),
            show="headings",
            height=18,
        )
        for col, title, width in [
            ("id", "ID", 60),
            ("name", "Nome", 220),
            ("reg", "Matrícula", 120),
            ("type", "Tipo", 180),
            ("time", "Data/Hora", 180),
            ("origin", "Origem", 140),
        ]:
            self.punch_tree.heading(col, text=title)
            self.punch_tree.column(col, width=width, anchor="w")
        self.punch_tree.pack(fill="both", expand=True)

    def _build_reports_tab(self) -> None:
        frame = ttk.Frame(self.tab_reports, padding=16)
        frame.pack(fill="both", expand=True)

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(0, 10))

        ttk.Label(toolbar, text="Mês (AAAA-MM)").pack(side="left")
        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        ttk.Entry(toolbar, textvariable=self.month_var, width=12).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Atualizar", command=self.refresh_summary).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Exportar CSV", command=self.export_summary_csv).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Imprimir relatório", command=self.print_monthly_report).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Espelho de ponto", command=self.print_employee_timesheet).pack(side="left")

        self.summary_tree = ttk.Treeview(frame, columns=("name", "reg", "dep", "total"), show="headings", height=20)
        self.summary_tree.heading("name", text="Nome")
        self.summary_tree.heading("reg", text="Matrícula")
        self.summary_tree.heading("dep", text="Departamento")
        self.summary_tree.heading("total", text="Total de registros")
        self.summary_tree.column("name", width=340)
        self.summary_tree.column("reg", width=160)
        self.summary_tree.column("dep", width=180)
        self.summary_tree.column("total", width=180, anchor="center")
        self.summary_tree.pack(fill="both", expand=True)

        aviso = "A conferência jurídica/contábil final deve seguir regras vigentes da Portaria 671/MTP e normas coletivas."
        ttk.Label(frame, text=aviso).pack(anchor="w", pady=(8, 0))

    def _build_controlid_tab(self) -> None:
        frame = ttk.Frame(self.tab_controlid, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Conector Control iD Access",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        info = (
            "Configure os dados de comunicação para importar batidas do equipamento Control iD.\n"
            "Também é possível enviar cadastro de colaborador para endpoint HTTP do integrador."
        )
        ttk.Label(frame, text=info).pack(anchor="w", pady=(0, 18))

        form = ttk.Frame(frame)
        form.pack(anchor="w")

        self.control_host = tk.StringVar(value=self.db.get_setting("controlid_host", ""))
        self.control_port = tk.StringVar(value=self.db.get_setting("controlid_port", "80"))
        self.control_token = tk.StringVar(value=self.db.get_setting("controlid_token", ""))
        self.control_employee_endpoint = tk.StringVar(value=self.db.get_setting("controlid_employee_endpoint", "/api/employees"))

        ttk.Label(form, text="Host/IP").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.control_host, width=40).grid(row=1, column=0, padx=(0, 10), pady=(0, 10))

        ttk.Label(form, text="Porta").grid(row=0, column=1, sticky="w")
        ttk.Entry(form, textvariable=self.control_port, width=10).grid(row=1, column=1, padx=(0, 10), pady=(0, 10))

        ttk.Label(form, text="Token/API Key").grid(row=2, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.control_token, width=54, show="*").grid(row=3, column=0, columnspan=2, pady=(0, 10), sticky="we")

        ttk.Label(form, text="Endpoint de colaborador").grid(row=4, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.control_employee_endpoint, width=54).grid(row=5, column=0, columnspan=2, pady=(0, 10), sticky="we")

        actions = ttk.Frame(frame)
        actions.pack(anchor="w", pady=8)
        ttk.Button(actions, text="Salvar configuração", command=self.save_controlid_settings).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Importar CSV do equipamento", command=self.import_punches_csv).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Enviar colaborador selecionado", command=self.send_selected_employee_to_controlid).pack(side="left")

        self.sync_status = tk.StringVar(value="Status: aguardando configuração")
        ttk.Label(frame, textvariable=self.sync_status).pack(anchor="w", pady=(14, 0))

    def _build_settings_tab(self) -> None:
        frame = ttk.Frame(self.tab_settings, padding=16)
        frame.pack(fill="both", expand=True)

        self.company_name = tk.StringVar(value=self.db.get_setting("company_name", "Minha Empresa"))
        self.company_cnpj = tk.StringVar(value=self.db.get_setting("company_cnpj", ""))

        ttk.Label(frame, text="Empresa").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.company_name, width=40).grid(row=1, column=0, padx=(0, 10), pady=(0, 10))

        ttk.Label(frame, text="CNPJ").grid(row=0, column=1, sticky="w")
        ttk.Entry(frame, textvariable=self.company_cnpj, width=20).grid(row=1, column=1, padx=(0, 10), pady=(0, 10))

        ttk.Button(frame, text="Salvar configurações gerais", command=self.save_general_settings).grid(row=2, column=0, sticky="w")

    def _on_employee_select(self, _event: tk.Event) -> None:
        selected = self.employee_tree.selection()
        if not selected:
            return
        row = self.employee_tree.item(selected[0], "values")
        self.selected_employee_id = int(row[0])
        self.name_var.set(row[1])
        self.reg_var.set(row[2])
        self.dep_var.set(row[3])
        self.pis_var.set(row[4])

    def create_employee(self) -> None:
        name = self.name_var.get().strip()
        reg = self.reg_var.get().strip()
        dep = self.dep_var.get().strip()
        pis = self.pis_var.get().strip()

        if not name or not reg or not dep:
            messagebox.showwarning(APP_NAME, "Preencha nome, matrícula e departamento.")
            return

        try:
            self.db.add_employee(name, reg, dep, pis)
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_NAME, "Já existe um colaborador com esta matrícula.")
            return

        self._clear_employee_form()
        self.refresh_employees()
        messagebox.showinfo(APP_NAME, "Colaborador cadastrado com sucesso.")

    def update_employee(self) -> None:
        if self.selected_employee_id is None:
            messagebox.showwarning(APP_NAME, "Selecione um colaborador para editar.")
            return
        try:
            self.db.update_employee(
                self.selected_employee_id,
                self.name_var.get(),
                self.reg_var.get(),
                self.dep_var.get(),
                self.pis_var.get(),
            )
        except sqlite3.IntegrityError:
            messagebox.showerror(APP_NAME, "Matrícula já usada por outro colaborador.")
            return
        self.refresh_employees()
        messagebox.showinfo(APP_NAME, "Colaborador atualizado.")

    def delete_employee(self) -> None:
        if self.selected_employee_id is None:
            messagebox.showwarning(APP_NAME, "Selecione um colaborador para excluir.")
            return
        if not messagebox.askyesno(APP_NAME, "Confirma exclusão do colaborador e suas batidas?"):
            return
        self.db.delete_employee(self.selected_employee_id)
        self.selected_employee_id = None
        self._clear_employee_form()
        self.refresh_employees()
        self.refresh_punches()
        self.refresh_summary()
        messagebox.showinfo(APP_NAME, "Colaborador excluído.")

    def _clear_employee_form(self) -> None:
        self.name_var.set("")
        self.reg_var.set("")
        self.dep_var.set("")
        self.pis_var.set("")

    def refresh_employees(self) -> None:
        rows = self.db.list_employees()
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        combo_values = []
        for emp in rows:
            self.employee_tree.insert("", "end", values=(emp.id, emp.name, emp.registration, emp.department, emp.pis))
            combo_values.append(f"{emp.id} - {emp.name} ({emp.registration})")
        self.employee_combo["values"] = combo_values

    def create_punch(self) -> None:
        selected = self.punch_employee_var.get()
        if not selected:
            messagebox.showwarning(APP_NAME, "Selecione um colaborador.")
            return
        employee_id = int(selected.split(" - ", maxsplit=1)[0])
        self.db.add_punch(employee_id, self.punch_type_var.get(), self.punch_origin_var.get())
        self.refresh_punches()
        self.refresh_summary()
        messagebox.showinfo(APP_NAME, "Ponto registrado.")

    def refresh_punches(self) -> None:
        for item in self.punch_tree.get_children():
            self.punch_tree.delete(item)

        for punch in self.db.list_punches():
            self.punch_tree.insert("", "end", values=punch)

    def refresh_summary(self) -> None:
        month = self.month_var.get().strip()
        if not month:
            month = datetime.now().strftime("%Y-%m")
            self.month_var.set(month)

        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)

        for row in self.db.monthly_summary(month):
            self.summary_tree.insert("", "end", values=row)

    def export_summary_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar relatório",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path:
            return

        month = self.month_var.get().strip() or datetime.now().strftime("%Y-%m")
        rows = self.db.monthly_summary(month)
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["nome", "matricula", "departamento", "total_registros", "referencia"])
            for row in rows:
                writer.writerow([row[0], row[1], row[2], row[3], month])

        messagebox.showinfo(APP_NAME, "Relatório exportado com sucesso.")

    def save_controlid_settings(self) -> None:
        self.db.upsert_setting("controlid_host", self.control_host.get().strip())
        self.db.upsert_setting("controlid_port", self.control_port.get().strip())
        self.db.upsert_setting("controlid_token", self.control_token.get().strip())
        self.db.upsert_setting("controlid_employee_endpoint", self.control_employee_endpoint.get().strip())
        self.sync_status.set("Status: configuração salva pela NexusLab")
        messagebox.showinfo(APP_NAME, "Configurações salvas.")

    def save_general_settings(self) -> None:
        self.db.upsert_setting("company_name", self.company_name.get().strip())
        self.db.upsert_setting("company_cnpj", self.company_cnpj.get().strip())
        messagebox.showinfo(APP_NAME, "Configurações gerais salvas.")

    def import_punches_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Importar batidas",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        employees = {emp.registration: emp.id for emp in self.db.list_employees()}
        imported = 0
        skipped = 0
        with open(path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                reg = (row.get("matricula") or "").strip()
                typ = (row.get("tipo") or "Entrada").strip()
                if reg not in employees:
                    skipped += 1
                    continue
                self.db.add_punch(employees[reg], typ, "Importação")
                imported += 1

        self.refresh_punches()
        self.refresh_summary()
        self.sync_status.set(f"Status: importação concluída ({imported} importados, {skipped} ignorados)")
        messagebox.showinfo(APP_NAME, f"Importação finalizada.\nImportados: {imported}\nIgnorados: {skipped}")

    def import_afd(self) -> None:
        path = filedialog.askopenfilename(
            title="Importar AFD",
            filetypes=[("AFD/TXT", "*.afd *.txt"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        employees = {emp.registration: emp.id for emp in self.db.list_employees()}
        imported = 0
        skipped = 0

        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                parsed = self._parse_afd_line(line)
                if parsed is None:
                    continue
                registration, punch_dt = parsed
                if registration not in employees:
                    skipped += 1
                    continue
                self.db.add_punch(employees[registration], "Entrada", "Importação AFD", punch_time=punch_dt)
                imported += 1

        self.refresh_punches()
        self.refresh_summary()
        messagebox.showinfo(
            APP_NAME,
            f"AFD processado.\nImportados: {imported}\nIgnorados: {skipped}\n\n"
            "Obs.: valide o layout AFD do seu REP/Control iD conforme Portaria 671.",
        )

    def _parse_afd_line(self, line: str) -> tuple[str, str] | None:
        # Suporta dois formatos comuns:
        # 1) Delimitado por ';' com colunas: matricula;data;hora
        # 2) Linha fixa contendo ...YYYYMMDDHHMM...matricula ao final
        text = line.strip()
        if not text:
            return None
        if ";" in text:
            parts = [p.strip() for p in text.split(";")]
            if len(parts) >= 3 and parts[0] and parts[1] and parts[2]:
                try:
                    dt = datetime.strptime(parts[1] + parts[2], "%Y%m%d%H%M").strftime("%Y-%m-%dT%H:%M:%S")
                    return parts[0], dt
                except ValueError:
                    return None
        digits = "".join(ch for ch in text if ch.isdigit())
        if len(digits) >= 12:
            match = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])([0-1]\d|2[0-3])([0-5]\d)", digits)
            if match:
                date_part = "".join(match.groups()[:3])
                time_part = "".join(match.groups()[3:])
                suffix = digits[match.end():]
                registration_raw = suffix[-12:] if suffix else digits[-12:]
                registration = registration_raw[-6:].lstrip("0") or registration_raw[-6:]
                try:
                    dt = datetime.strptime(date_part + time_part, "%Y%m%d%H%M").strftime("%Y-%m-%dT%H:%M:%S")
                    return registration, dt
                except ValueError:
                    return None
        return None

    def send_selected_employee_to_controlid(self) -> None:
        if self.selected_employee_id is None:
            messagebox.showwarning(APP_NAME, "Selecione um colaborador na aba de colaboradores.")
            return

        selected_emp = next((e for e in self.db.list_employees() if e.id == self.selected_employee_id), None)
        if selected_emp is None:
            messagebox.showerror(APP_NAME, "Colaborador não encontrado.")
            return

        host = self.control_host.get().strip()
        port = self.control_port.get().strip()
        token = self.control_token.get().strip()
        endpoint = self.control_employee_endpoint.get().strip() or "/api/employees"
        if not host:
            messagebox.showwarning(APP_NAME, "Configure host do Control iD.")
            return

        url = f"http://{host}:{port}{endpoint}"
        payload = {
            "name": selected_emp.name,
            "registration": selected_emp.registration,
            "department": selected_emp.department,
            "pis": selected_emp.pis,
        }
        data = json.dumps(payload).encode("utf-8")
        req = request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/json")
        if token:
            req.add_header("Authorization", f"Bearer {token}")

        try:
            with request.urlopen(req, timeout=8) as resp:
                status = resp.status
        except error.URLError as exc:
            messagebox.showerror(APP_NAME, f"Falha ao enviar colaborador: {exc}")
            return

        self.sync_status.set(f"Status: colaborador enviado (HTTP {status})")
        messagebox.showinfo(APP_NAME, f"Colaborador enviado ao integrador Control iD (HTTP {status}).")

    def print_monthly_report(self) -> None:
        month = self.month_var.get().strip() or datetime.now().strftime("%Y-%m")
        rows = self.db.monthly_summary(month)
        company = self.company_name.get().strip() or "Empresa"

        trs = "".join(
            f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td style='text-align:center'>{r[3]}</td></tr>" for r in rows
        )
        html = f"""
        <html><head><meta charset='utf-8'><title>Relatório {month}</title></head>
        <body style='font-family:Segoe UI, sans-serif'>
            <h2>{company} — Relatório Mensal de Ponto ({month})</h2>
            <table border='1' cellspacing='0' cellpadding='6'>
                <tr><th>Nome</th><th>Matrícula</th><th>Departamento</th><th>Total de Registros</th></tr>
                {trs}
            </table>
            <p>Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} por NexusLab Ponto Pro.</p>
        </body></html>
        """
        self._open_html_for_print(html, f"relatorio_{month}.html")

    def print_employee_timesheet(self) -> None:
        selected = self.summary_tree.selection()
        if not selected:
            messagebox.showwarning(APP_NAME, "Selecione um colaborador no relatório para gerar espelho.")
            return

        row = self.summary_tree.item(selected[0], "values")
        registration = row[1]
        month = self.month_var.get().strip() or datetime.now().strftime("%Y-%m")

        employees = self.db.list_employees()
        emp = next((e for e in employees if e.registration == registration), None)
        if not emp:
            messagebox.showerror(APP_NAME, "Colaborador não encontrado.")
            return

        punches = self.db.punches_for_employee_month(emp.id, month)
        lines = "".join(
            f"<tr><td>{p[0]}</td><td>{p[1].replace('T', ' ')}</td><td>{p[2]}</td></tr>" for p in punches
        )

        html = f"""
        <html><head><meta charset='utf-8'><title>Espelho {emp.name}</title></head>
        <body style='font-family:Segoe UI, sans-serif'>
            <h2>Espelho de Ponto — {month}</h2>
            <p><strong>Empresa:</strong> {self.company_name.get().strip()}</p>
            <p><strong>Colaborador:</strong> {emp.name} | <strong>Matrícula:</strong> {emp.registration} | <strong>PIS:</strong> {emp.pis}</p>
            <table border='1' cellspacing='0' cellpadding='6'>
                <tr><th>Tipo</th><th>Data/Hora</th><th>Origem</th></tr>
                {lines}
            </table>
            <p>Documento para conferência interna. Validar regras legais vigentes e convenção coletiva aplicável.</p>
            <p>Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} por NexusLab.</p>
        </body></html>
        """
        self._open_html_for_print(html, f"espelho_{emp.registration}_{month}.html")

    def _open_html_for_print(self, html: str, filename_hint: str) -> None:
        with tempfile.NamedTemporaryFile("w", delete=False, suffix="_" + filename_hint, encoding="utf-8") as f:
            f.write(html)
            temp_path = f.name
        webbrowser.open(f"file://{temp_path}")
        messagebox.showinfo(APP_NAME, "Arquivo aberto no navegador para impressão/salvar em PDF.")


def main() -> None:
    db = Database(DB_PATH)
    app = App(db)
    app.mainloop()


if __name__ == "__main__":
    main()
