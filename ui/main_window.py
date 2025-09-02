# -*- coding: utf-8 -*-
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from typing import List, Dict, Any, Optional
from datetime import date, datetime
try:
    from tkcalendar import DateEntry
except Exception:
    # Fallback stub to avoid import error during static analysis (real app should have tkcalendar installed)
    class DateEntry(ttk.Entry):
        def get_date(self):
            return date.today()

from constants import (
    APP_TITLE, APP_GEOMETRY, CATEGORIES, REAGENT_TYPES, REAGENT_QUALIFICATIONS,
    UNITS, DEPARTMENTS, COLOR_EXPIRED, COLOR_SOON, COLOR_NORMAL, EXPIRY_SOON_THRESHOLD,
    NEEDS_DEPARTMENTS, STORAGE_DEPARTMENT, QA_DEPARTMENT
)
from models import Item
from storage import (
    load_items, save_items, get_next_seq_id,
    load_users, save_users, ensure_default_admin, hash_password,
    load_needs, save_needs, next_need_id, next_qa_request_id, next_issue_id, next_store_request_id
)
from exports import export_stock_to_excel, export_issue_docx

def parse_date(s: str):
    if not s: return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None

def run_app():
    ensure_default_admin()
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry(APP_GEOMETRY)
    app = MainApp(root)
    root.mainloop()

class MainApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.current_user: Dict[str, Any] = {}
        self.items: List[Item] = load_items()
        self.needs: Dict[str, Any] = load_needs()
        self.style = ttk.Style()
        self.style.configure(
            "Danger.TButton",
            foreground="white",
            background="#d43f3a",
            borderwidth=0,
            padding=(6, 2),
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", "#c9302c"), ("!disabled", "#d9534f")],
        )
        self.btn_qa_requests = None
        self.btn_store_requests = None
        self._build_login()

    # Utility: search resets (also wired via lambdas in buttons for robustness)
    def reset_inv_search(self):
        if hasattr(self, "var_inv_search"):
            self.var_inv_search.set("")
        self.apply_search()

    def reset_needs_search(self):
        if hasattr(self, "var_needs_search"):
            self.var_needs_search.set("")
        self.apply_search()

    # Login UI
    def _build_login(self):
        self.login_frame = ttk.Frame(self.root, padding=24)
        self.login_frame.pack(expand=True)
        ttk.Label(self.login_frame, text="Вход", font=("Segoe UI", 16)).grid(row=0, column=0, columnspan=2, pady=10)
        ttk.Label(self.login_frame, text="Логин").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(self.login_frame, text="Пароль").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.var_user = tk.StringVar(); self.var_pass = tk.StringVar()
        ttk.Entry(self.login_frame, textvariable=self.var_user, width=30).grid(row=1, column=1, padx=6, pady=6)
        ttk.Entry(self.login_frame, textvariable=self.var_pass, width=30, show="*").grid(row=2, column=1, padx=6, pady=6)
        ttk.Button(self.login_frame, text="Войти", command=self._do_login).grid(row=3, column=0, columnspan=2, pady=12)
        self.login_frame.columnconfigure(0, weight=1); self.login_frame.columnconfigure(1, weight=1)

    def _do_login(self):
        username = self.var_user.get().strip(); password = self.var_pass.get().strip()
        if not username or not password:
            messagebox.showwarning("Вход", "Введите логин и пароль"); return
        users = load_users(); pw_hash = hash_password(password)
        for u in users:
            if u.get("username")==username and u.get("password_hash")==pw_hash:
                self.current_user = u; self.login_frame.destroy(); self._build_main_ui(); return
        messagebox.showerror("Вход", "Неверный логин и пароль")

    # Main UI
    def _build_main_ui(self):
        role = self.current_user.get("role"); dept = self.current_user.get("department")
        plan_year = self.needs.get("plan_year"); locked = self.needs.get("locked")
        self.paned = ttk.Panedwindow(self.root, orient=tk.VERTICAL); self.paned.pack(fill="both", expand=True)

        # Top: Needs
        needs_wrapper = ttk.Frame(self.paned); self.paned.add(needs_wrapper, weight=1)
        needs_bar = ttk.Frame(needs_wrapper, padding=(10,6)); needs_bar.pack(fill="x")
        ttk.Label(needs_bar, text=f"План {plan_year} {'(утвержден)' if locked else '(черновик)'}").pack(side="left")
        ttk.Label(needs_bar, text=f" | Пользователь: {self.current_user.get('username')} ({role}, {dept})").pack(side="left", padx=(6,0))
        ttk.Label(needs_bar, text="Поиск по потребностям:").pack(side="left", padx=(16,6))
        self.var_needs_search = tk.StringVar()
        ent_n = ttk.Entry(needs_bar, textvariable=self.var_needs_search, width=36); ent_n.pack(side="left")
        ent_n.bind("<KeyRelease>", lambda e: self.apply_search())
        ttk.Button(needs_bar, text="Сбросить", command=lambda: (self.var_needs_search.set(""), self.apply_search())).pack(side="left", padx=(6,12))
        if dept==QA_DEPARTMENT or role=="admin":
            self.btn_qa_requests = ttk.Button(needs_bar, text="Входящие запросы (ОУК)", command=self.show_qa_requests)
            self.btn_qa_requests.pack(side="right", padx=6)
        else:
            self.btn_qa_requests = None
        if dept==STORAGE_DEPARTMENT or role=="admin":
            self.btn_store_requests = ttk.Button(needs_bar, text="Входящие запросы (склад)", command=self.show_store_requests)
            self.btn_store_requests.pack(side="right", padx=6)
        else:
            self.btn_store_requests = None
        if role=="admin":
            ttk.Button(needs_bar, text="Пользователи", command=self.manage_users).pack(side="right", padx=6)
            ttk.Button(needs_bar, text="Утвердить план", command=self.approve_plan).pack(side="right", padx=6)

        self.needs_nb = ttk.Notebook(needs_wrapper); self.needs_nb.pack(fill="both", expand=True, padx=8, pady=(0,8))
        self.needs_trees: Dict[str, ttk.Treeview] = {}
        self.needs_order: List[str] = []
        self._build_needs_tabs()

        # Bottom: Inventory
        inv_wrapper = ttk.Frame(self.paned); self.paned.add(inv_wrapper, weight=3)
        inv_bar = ttk.Frame(inv_wrapper, padding=(10,6)); inv_bar.pack(fill="x")
        ttk.Label(inv_bar, text="Учет прихода позиций (склад)").pack(side="left")
        ttk.Label(inv_bar, text="Поиск по складу:").pack(side="left", padx=(16,6))
        self.var_inv_search = tk.StringVar()
        ent_i = ttk.Entry(inv_bar, textvariable=self.var_inv_search, width=36); ent_i.pack(side="left")
        ent_i.bind("<KeyRelease>", lambda e: self.apply_search())
        ttk.Button(inv_bar, text="Сбросить", command=lambda: (self.var_inv_search.set(""), self.apply_search())).pack(side="left", padx=(6,0))
        ttk.Button(inv_bar, text="Добавить позицию", command=self.add_item_dialog).pack(side="right", padx=6)
        ttk.Button(inv_bar, text="Редактировать", command=self.edit_selected_item).pack(side="right", padx=6)
        ttk.Button(inv_bar, text="Удалить", command=self.delete_selected_item).pack(side="right", padx=6)
        ttk.Button(inv_bar, text="Экспорт Excel", command=self.export_excel_dialog).pack(side="right", padx=6)
        ttk.Button(inv_bar, text="Экспорт DOCX (выдача)", command=self.export_docx_dialog).pack(side="right", padx=6)
        if role=="admin":
            ttk.Button(inv_bar, text="Пользователи", command=self.manage_users).pack(side="right", padx=6)

        self.inv_nb = ttk.Notebook(inv_wrapper); self.inv_nb.pack(fill="both", expand=True, padx=8, pady=(0,8))
        self.inv_trees: Dict[str, ttk.Treeview] = {}
        self.inv_order: List[str] = ["Реактивы", "ГСО-ПГС-СО", "Расходные материалы"]
        self._build_inventory_tabs()
        self.reload_all_trees()

        self.root.update_idletasks()
        try:
            total_h = self.root.winfo_height(); top_h = max(220, int(total_h*0.30)); self.paned.sashplace(0, 0, top_h)
        except Exception:
            pass

    def _build_inventory_tabs(self):
        self._build_inv_tab("Реактивы", [
            ("seq_id","ID",80),("name","Наименование",260),("quantity","Количество",110),("unit","Ед.",60),
            ("packaging","Фасовка",120),("storage_place","Место хранения",140),("expiry_date","Срок годности",120),
            ("date_received","Дата поступления",130),("batch_number","Номер партии",120),("responsible","Ответственный",140),
            ("qualification","Квалификация",140),("reagent_type","Тип реактива",140)
        ])
        self._build_inv_tab("ГСО-ПГС-СО", [
            ("seq_id","ID",80),("name","Наименование",260),("quantity","Количество",110),("unit","Ед.",60),
            ("packaging","Фасовка",120),("storage_place","Место хранения",140),("expiry_date","Срок годности",120),
            ("date_received","Дата поступления",130),("batch_number","Номер партии",120),("responsible","Ответственный",140),
            ("state_register_no","№ в госреестре СО",160),("certified_value","Аттестованное значение",180),
            ("manufacture_date","Дата выпуска",120),("manufacturer","Производитель",140),("storage_conditions","Условия хранения",160)
        ])
        self._build_inv_tab("Расходные материалы", [
            ("seq_id","ID",80),("name","Наименование",300),("quantity","Количество",110),("unit","Ед.",60),
            ("packaging","Фасовка",120),("storage_place","Место хранения",160),("date_received","Дата поступления",130),
            ("batch_number","Номер партии",120),("responsible","Ответственный",160)
        ])

    def _build_inv_tab(self, category: str, columns: List[tuple]):
        frame = ttk.Frame(self.inv_nb)
        self.inv_nb.add(frame, text=category)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        for key, title, width in columns:
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor="w")
        self.inv_trees[category] = tree

    def _build_needs_tabs(self):
        role = self.current_user.get("role")
        dept = self.current_user.get("department")
        locked = self.needs.get("locked")
        show_all = (role=="admin") or (dept==STORAGE_DEPARTMENT) or (dept==QA_DEPARTMENT)
        depts_to_show = NEEDS_DEPARTMENTS if show_all else [dept]
        for d in depts_to_show:
            page = ttk.Frame(self.needs_nb)
            self.needs_nb.add(page, text=f"Потребности — {d}")
            self.needs_order.append(d)
            tb = ttk.Frame(page, padding=(6,4)); tb.pack(fill="x")
            btn_add = ttk.Button(tb, text="+ Добавить", command=lambda dep=d: self.add_need_dialog(dep))
            btn_edit = ttk.Button(tb, text="Редактировать", command=lambda dep=d: self.edit_need_dialog(dep))
            btn_del = ttk.Button(tb, text="Удалить", command=lambda dep=d: self.delete_need(dep))
            btn_issue = ttk.Button(tb, text="Выдать выбранную (склад)", command=lambda dep=d: self.issue_against_need(dep))
            btn_request = ttk.Button(tb, text="Запросить выдачу", command=lambda dep=d: self.request_issue_from_need(dep))
            btn_add.pack(side="left"); btn_edit.pack(side="left", padx=(6,0)); btn_del.pack(side="left", padx=(6,0))
            if self.current_user.get("department")==STORAGE_DEPARTMENT or self.current_user.get("role")=="admin":
                btn_issue.pack(side="left", padx=(12,0))
            if (self.current_user.get("role")!="admin") and (self.current_user.get("department")==d) and (self.current_user.get("department") not in [STORAGE_DEPARTMENT, QA_DEPARTMENT]):
                btn_request.pack(side="left", padx=(12,0))
            can_edit = (not locked) and ((self.current_user.get("role")!="admin") and (dept==d))
            if not can_edit:
                btn_add.config(state="disabled"); btn_edit.config(state="disabled"); btn_del.config(state="disabled")
            cols = [
                ("need_id","ID",70), ("category","Категория",140), ("item_name","Наименование",260),
                ("plan_qty","Кол-во (план)",120), ("remaining_qty","Кол-во (остаток)",140), ("unit","Ед.",70),
                ("qualification","Квалификация",140), ("state_register_no","№ в реестре СО",160),
                ("cylinder_volume","Объем баллона",140), ("certified_value","Аттестованное значение",180),
                ("purpose","Цель использования",220),
            ]
            tree = ttk.Treeview(page, columns=[c[0] for c in cols], show="headings", selectmode="browse")
            vsb = ttk.Scrollbar(page, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y"); tree.pack(side="left", fill="both", expand=True, padx=6, pady=(0,6))
            for key, title, width in cols:
                tree.heading(key, text=title); tree.column(key, width=width, anchor="w")
            self.needs_trees[d] = tree

    def reload_all_trees(self):
        # Inventory
        for cat, tree in self.inv_trees.items():
            for row in tree.get_children():
                tree.delete(row)
        for it in self.items:
            self._insert_item(it)

        # Needs
        for dep, tree in self.needs_trees.items():
            for row in tree.get_children():
                tree.delete(row)
            for n in self.needs.get("departments", {}).get(dep, []):
                tree.insert("", "end", iid=f"{dep}-{n.get('need_id')}", values=(
                    n.get("need_id"), n.get("category"), n.get("item_name"),
                    n.get("plan_qty"), n.get("remaining_qty"), n.get("unit"),
                    n.get("qualification") or "", n.get("state_register_no") or "",
                    n.get("cylinder_volume") or "", n.get("certified_value") or "",
                    n.get("purpose") or ""
                ))
        self.apply_search()
        self._update_request_buttons()

    def _update_request_buttons(self):
        pending_qa = any(r.get("status") == "pending" for r in self.needs.get("qa_overflow_requests", []))
        pending_store = any(r.get("status") == "pending" for r in self.needs.get("store_requests", []))
        if self.btn_qa_requests:
            self.btn_qa_requests.configure(style="Danger.TButton" if pending_qa else "TButton")
        if self.btn_store_requests:
            self.btn_store_requests.configure(style="Danger.TButton" if pending_store else "TButton")

    def _insert_item(self, it: Item):
        tree = self.inv_trees.get(it.category)
        if not tree:
            return
        col_keys = list(tree["columns"])
        mapping = {
            "seq_id": it.seq_id,
            "name": it.name,
            "quantity": it.quantity,
            "unit": it.unit,
            "packaging": it.packaging,
            "storage_place": it.storage_place,
            "expiry_date": it.expiry_date or "",
            "date_received": it.date_received or "",
            "batch_number": it.batch_number,
            "responsible": it.responsible,
            "qualification": it.qualification or "",
            "reagent_type": it.reagent_type or "",
            "state_register_no": it.state_register_no or "",
            "certified_value": it.certified_value or "",
            "manufacture_date": it.manufacture_date or "",
            "manufacturer": it.manufacturer or "",
            "storage_conditions": it.storage_conditions or "",
        }
        values = tuple(mapping.get(k, "") for k in col_keys)
        tree.insert("", "end", iid=f"{it.category}-{it.seq_id}", values=values)

    def get_selected_inventory_tree(self):
        idx = self.inv_nb.index(self.inv_nb.select())
        cat = self.inv_order[idx]
        return self.inv_trees[cat], cat

    def get_selected_needs_department(self):
        idx = self.needs_nb.index(self.needs_nb.select())
        dept = self.needs_order[idx]
        return self.needs_trees[dept], dept

    def apply_search(self):
        inv_query = (self.var_inv_search.get().strip().lower() if hasattr(self, "var_inv_search") else "")
        for cat, tree in self.inv_trees.items():
            for iid in tree.get_children():
                tree.reattach(iid, "", "end")
            if not inv_query:
                continue
            for iid in list(tree.get_children()):
                txt = " ".join(str(v) for v in tree.item(iid, "values")).lower()
                if inv_query not in txt:
                    tree.detach(iid)
        needs_query = (self.var_needs_search.get().strip().lower() if hasattr(self, "var_needs_search") else "")
        for dep, tree in self.needs_trees.items():
            for iid in tree.get_children():
                tree.reattach(iid, "", "end")
            if not needs_query:
                continue
            for iid in list(tree.get_children()):
                txt = " ".join(str(v) for v in tree.item(iid, "values")).lower()
                if needs_query not in txt:
                    tree.detach(iid)

    # Items CRUD
    def add_item_dialog(self):
        ItemDialog(self.root, title="Добавить позицию", on_save=self._add_item_save, default_responsible=self.current_user.get('username'))

    def _add_item_save(self, payload: dict):
        if not payload.get('responsible'):
            payload['responsible'] = self.current_user.get('username')
        seq = get_next_seq_id(self.items)
        payload["seq_id"] = seq
        it = Item.from_dict(payload)
        self.items.append(it)
        save_items(self.items)
        self._insert_item(it)
        self.apply_search()

    def edit_selected_item(self):
        tree, _ = self.get_selected_inventory_tree()
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Редактирование", "Выберите позицию")
            return
        seq_id = int(tree.item(sel[0], "values")[0])
        it = next((x for x in self.items if x.seq_id == seq_id), None)
        if not it:
            messagebox.showerror("Редактирование", "Позиция не найдена")
            return
        ItemDialog(self.root, title="Редактировать позицию", item=it, on_save=self._edit_item_save, default_responsible=self.current_user.get('username'))

    def _edit_item_save(self, payload: dict):
        seq_id = payload.get("seq_id")
        for i, it in enumerate(self.items):
            if it.seq_id == seq_id:
                self.items[i] = Item.from_dict(payload)
                save_items(self.items)
                self.reload_all_trees()
                return

    def delete_selected_item(self):
        tree, _ = self.get_selected_inventory_tree()
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Удаление", "Выберите позицию")
            return
        seq_id = int(tree.item(sel[0], "values")[0])
        if not messagebox.askyesno("Удаление", f"Удалить позицию ID {seq_id}?"):
            return
        self.items = [x for x in self.items if x.seq_id != seq_id]
        save_items(self.items)
        self.reload_all_trees()

    # Export
    def export_excel_dialog(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        export_stock_to_excel(self.items, path)
        messagebox.showinfo("Экспорт", "Экспорт завершен")

    def export_docx_dialog(self):
        tree, _ = self.get_selected_inventory_tree()
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Экспорт DOCX", "Выберите позицию")
            return
        seq_id = int(tree.item(sel[0], "values")[0])
        it = next((x for x in self.items if x.seq_id == seq_id), None)
        if not it: return
        template = filedialog.askopenfilename(title="Выберите DOCX шаблон", filetypes=[("DOCX", "*.docx")])
        if not template: return
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOCX", "*.docx")])
        if not path: return
        export_issue_docx(it, path, template)
        messagebox.showinfo("Экспорт DOCX", "Документ сформирован")

    # Users / QA / Store
    def manage_users(self):
        if self.current_user.get("role") != "admin":
            messagebox.showwarning("Пользователи", "Недостаточно прав"); return
        UsersWindow(self.root)

    def show_qa_requests(self):
        win = QARequestsWindow(self.root, self.needs)
        self.root.wait_window(win)
        self.reload_all_trees()

    def show_store_requests(self):
        win = StoreRequestsWindow(self.root, self)
        self.root.wait_window(win)
        self.reload_all_trees()

    def approve_plan(self):
        if self.current_user.get("role") != "admin": return
        if self.needs.get("locked"):
            messagebox.showinfo("План", "План уже утвержден"); return
        if not messagebox.askyesno("План", "Утвердить план? Внесение новых потребностей будет заблокировано."): return
        self.needs["locked"] = True
        save_needs(self.needs)
        messagebox.showinfo("План", "План утвержден")
        self.reload_all_trees()

    # Needs CRUD
    def add_need_dialog(self, department: str):
        if self.needs.get("locked"):
            messagebox.showwarning("Потребности", "План утвержден, добавление запрещено"); return
        if self.current_user.get("department") != department:
            messagebox.showwarning("Потребности", "Можно добавлять только в свой отдел"); return
        NeedDialog(self.root, title=f"Добавить потребность — {department}", on_save=lambda payload: self._add_need_save(department, payload), items=self.items)

    def _add_need_save(self, department: str, payload: dict):
        need_id = next_need_id(self.needs)
        payload["need_id"] = need_id
        payload["remaining_qty"] = payload["plan_qty"]
        payload["status"] = "planned"
        payload["approved_by_qa"] = False
        payload["created"] = date.today().strftime("%Y-%m-%d")
        self.needs["departments"].setdefault(department, []).append(payload)
        save_needs(self.needs)
        self.reload_all_trees()

    def edit_need_dialog(self, department: str):
        if self.needs.get("locked"):
            messagebox.showwarning("Потребности", "План утвержден, редактирование запрещено"); return
        if self.current_user.get("department") != department:
            messagebox.showwarning("Потребности", "Можно редактировать только свой отдел"); return
        tree = self.needs_trees[department]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Потребности", "Выберите строку"); return
        need_id = int(tree.item(sel[0], "values")[0])
        n = next((x for x in self.needs["departments"][department] if int(x.get("need_id"))==need_id), None)
        if not n:
            messagebox.showerror("Потребности", "Запись не найдена"); return
        NeedDialog(self.root, title=f"Редактировать потребность — {department}", on_save=lambda payload: self._edit_need_save(department, payload), need=n, items=self.items)

    def _edit_need_save(self, department: str, payload: dict):
        need_id = int(payload.get("need_id"))
        for i, n in enumerate(self.needs["departments"][department]):
            if int(n.get("need_id")) == need_id:
                already_issued = float(n.get("plan_qty",0)) - float(n.get("remaining_qty",0))
                new_plan = float(payload.get("plan_qty",0))
                new_remaining = max(0.0, new_plan - already_issued)
                payload["remaining_qty"] = new_remaining
                self.needs["departments"][department][i] = payload
                save_needs(self.needs)
                self.reload_all_trees()
                return

    def delete_need(self, department: str):
        if self.needs.get("locked"):
            messagebox.showwarning("Потребности", "План утвержден, удаление запрещено"); return
        if self.current_user.get("department") != department:
            messagebox.showwarning("Потребности", "Можно удалять только в своем отделе"); return
        tree = self.needs_trees[department]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Потребности", "Выберите строку"); return
        need_id = int(tree.item(sel[0], "values")[0])
        if not messagebox.askyesno("Удаление", f"Удалить запись ID {need_id}?"): return
        self.needs["departments"][department] = [n for n in self.needs["departments"][department] if int(n.get("need_id"))!=need_id]
        save_needs(self.needs)
        self.reload_all_trees()

    def issue_against_need(self, department: str):
        tree = self.needs_trees[department]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Выдача", "Выберите строку плана")
            return
        vals = tree.item(sel[0], "values")
        need_id = int(vals[0]); unit = vals[5]
        remaining = float(vals[4] or 0)
        qty = simpledialog.askfloat("Выдача", f"Введите количество для выдачи ({unit}). Остаток по плану: {remaining}", minvalue=0.0)
        if qty is None or qty <= 0:
            return
        result = self._process_issue(department, need_id, qty)
        messagebox.showinfo("Выдача", result)

    def request_issue_from_need(self, department: str):
        tree = self.needs_trees[department]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Запросить выдачу", "Выберите строку плана")
            return
        vals = tree.item(sel[0], "values")
        need_id = int(vals[0]); unit = vals[5]
        remaining = float(vals[4] or 0)
        qty = simpledialog.askfloat("Запросить выдачу", f"Сколько требуется выдать ({unit})? Остаток по плану: {remaining}", minvalue=0.0)
        if qty is None or qty <= 0:
            return
        req_id = next_store_request_id(self.needs)
        self.needs.setdefault("store_requests", []).append({
            "request_id": req_id, "department": department, "need_id": need_id,
            "requested_qty": qty, "unit": unit, "status": "pending",
            "created": date.today().strftime("%Y-%m-%d"),
            "requested_by": self.current_user.get("username")
        })
        save_needs(self.needs)
        messagebox.showinfo("Запросить выдачу", "Заявка отправлена на склад")

    def _process_issue(self, department: str, need_id: int, qty: float) -> str:
        # Process issuing an item against a department's plan
        n = None
        for d in self.needs.get("departments", {}).get(department, []):
            if int(d.get("need_id")) == int(need_id):
                n = d; break
        if not n:
            return "План не найден"
        category = n.get("category"); item_name = n.get("item_name"); unit = n.get("unit")
        it = next((x for x in self.items if x.category==category and x.name==item_name), None)
        if not it:
            return "Позиция не найдена на складе"
        if qty > it.quantity:
            return "Недостаточно остатка на складе"
        if qty > float(n.get("remaining_qty",0)):
            rq_id = next_qa_request_id(self.needs); extra = qty - float(n.get("remaining_qty",0))
            self.needs.setdefault("qa_overflow_requests", []).append({
                "request_id": rq_id, "department": department, "need_id": need_id,
                "category": category, "item_name": item_name, "requested_qty": qty, "excess_qty": extra,
                "unit": unit, "status": "pending", "created": date.today().strftime("%Y-%m-%d")
            })
            save_needs(self.needs)
            return "Превышение плана — заявка отправлена в ОУК"
        it.quantity -= qty; save_items(self.items)
        n["remaining_qty"] = float(n.get("remaining_qty",0)) - qty
        iss_id = next_issue_id(self.needs)
        self.needs.setdefault("issues", []).append({
            "issue_id": iss_id, "department": department, "need_id": need_id, "item_seq_id": it.seq_id,
            "item_name": it.name, "category": it.category, "qty": qty, "unit": it.unit,
            "date": date.today().strftime("%Y-%m-%d"), "issued_by": self.current_user.get("username")
        })
        save_needs(self.needs)
        self.reload_all_trees()
        return "Выдача выполнена"

# -------- Dialogs / Windows --------
class NeedDialog(tk.Toplevel):
    def __init__(self, master, title: str, on_save, items: List[Item], need: Optional[dict]=None):
        super().__init__(master); self.title(title); self.resizable(False, False)
        self.on_save = on_save; self.items = items; self.need = need or {}
        self.var_need_id = tk.StringVar(value=str(self.need.get("need_id","")))
        self.var_category = tk.StringVar(value=self.need.get("category","Реактивы"))
        self.var_item_name = tk.StringVar(value=self.need.get("item_name",""))
        self.var_plan_qty = tk.StringVar(value=str(self.need.get("plan_qty",0)))
        self.var_remaining_qty = tk.StringVar(value=str(self.need.get("remaining_qty", self.need.get("plan_qty",0))))
        self.var_unit = tk.StringVar(value=self.need.get("unit","шт"))
        self.var_qualification = tk.StringVar(value=self.need.get("qualification",""))
        self.var_state_register_no = tk.StringVar(value=self.need.get("state_register_no",""))
        self.var_cylinder_volume = tk.StringVar(value=self.need.get("cylinder_volume",""))
        self.var_certified_value = tk.StringVar(value=self.need.get("certified_value",""))
        self.var_purpose = tk.StringVar(value=self.need.get("purpose",""))

        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Категория").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        cb_cat = ttk.Combobox(frm, values=CATEGORIES, textvariable=self.var_category, state="readonly", width=28)
        cb_cat.grid(row=0, column=1, sticky="w", padx=6, pady=4); cb_cat.bind("<<ComboboxSelected>>", lambda e: self._toggle_extra())

        ttk.Label(frm, text="Наименование").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        entry_name = ttk.Entry(frm, textvariable=self.var_item_name, width=45); entry_name.grid(row=0, column=3, sticky="w", padx=6, pady=4)
        self.suggest_box = tk.Listbox(frm, height=5)
        def on_keyup(_):
            text = self.var_item_name.get().strip().lower()
            self.suggest_box.delete(0, tk.END)
            if not text: self.suggest_box.place_forget(); return
            options = sorted({it.name for it in self.items if text in it.name.lower()})
            if options:
                self.suggest_box.place(x=entry_name.winfo_x(), y=entry_name.winfo_y()+entry_name.winfo_height()+frm.winfo_y())
                for o in options[:50]: self.suggest_box.insert(tk.END, o)
            else:
                self.suggest_box.place_forget()
        entry_name.bind("<KeyRelease>", on_keyup)
        def on_pick(evt):
            sel = self.suggest_box.curselection()
            if sel:
                self.var_item_name.set(self.suggest_box.get(sel[0])); self.suggest_box.place_forget()
        self.suggest_box.bind("<<ListboxSelect>>", on_pick)

        ttk.Label(frm, text="Кол-во (план)").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_plan_qty, width=12).grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Кол-во (остаток)").grid(row=1, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_remaining_qty, width=12, state="disabled").grid(row=1, column=3, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Ед. изм.").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Combobox(frm, values=UNITS, textvariable=self.var_unit, width=12, state="readonly").grid(row=2, column=1, sticky="w", padx=6, pady=4)

        self.frm_reagents = ttk.Labelframe(frm, text="Доп. поля для категории 'Реактивы'", padding=8)
        ttk.Label(self.frm_reagents, text="Квалификация").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Combobox(self.frm_reagents, values=REAGENT_QUALIFICATIONS, textvariable=self.var_qualification, width=20, state="readonly").grid(row=0, column=1, sticky="w", padx=6, pady=4)

        self.frm_gso = ttk.Labelframe(frm, text="Доп. поля для 'ГСО-ПГС-СО'", padding=8)
        ttk.Label(self.frm_gso, text="№ в реестре СО").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_state_register_no, width=20).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(self.frm_gso, text="Объем баллона").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_cylinder_volume, width=20).grid(row=0, column=3, sticky="w", padx=6, pady=4)
        ttk.Label(self.frm_gso, text="Аттестованное значение").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_certified_value, width=20).grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Цель использования").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_purpose, width=60).grid(row=3, column=1, columnspan=3, sticky="w", padx=6, pady=4)

        btns = ttk.Frame(frm); btns.grid(row=4, column=0, columnspan=4, sticky="e", pady=8)
        ttk.Button(btns, text="Сохранить", command=self._save).pack(side="right", padx=6)
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right", padx=6)
        for i in range(4): frm.columnconfigure(i, weight=1)
        self._toggle_extra()

    def _toggle_extra(self):
        cat = self.var_category.get()
        if cat=="Реактивы":
            self.frm_reagents.grid(row=2, column=2, columnspan=2, sticky="ew", padx=6, pady=6)
            self.frm_gso.grid_remove()
        elif cat=="ГСО-ПГС-СО":
            self.frm_gso.grid(row=2, column=2, columnspan=2, sticky="ew", padx=6, pady=6)
            self.frm_reagents.grid_remove()
        else:
            self.frm_reagents.grid_remove(); self.frm_gso.grid_remove()

    def _save(self):
        name = self.var_item_name.get().strip()
        if not name:
            messagebox.showerror("Потребности", "Заполните поле 'Наименование'"); return
        try:
            plan = float(self.var_plan_qty.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Потребности", "Поле 'Кол-во (план)' должно быть числом"); return
        payload = {
            "need_id": int(self.var_need_id.get()) if self.var_need_id.get() else None,
            "category": self.var_category.get(),
            "item_name": name,
            "plan_qty": plan,
            "remaining_qty": float(self.var_remaining_qty.get()) if self.var_remaining_qty.get() else plan,
            "unit": self.var_unit.get(),
            "qualification": self.var_qualification.get() or None,
            "state_register_no": self.var_state_register_no.get() or None,
            "cylinder_volume": self.var_cylinder_volume.get() or None,
            "certified_value": self.var_certified_value.get() or None,
            "purpose": self.var_purpose.get().strip() or None,
        }
        self.on_save(payload); self.destroy()

class UsersWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master); self.title("Пользователи"); self.geometry("640x400")
        from storage import load_users, save_users, hash_password
        from constants import DEPARTMENTS
        self.users = load_users()
        self.tree = ttk.Treeview(self, columns=("username","role","department"), show="headings")
        for k,w in [("username",160),("role",120),("department",320)]:
            self.tree.heading(k, text=k); self.tree.column(k, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, side="top")
        self._reload()
        bar = ttk.Frame(self); bar.pack(fill="x", side="bottom")
        ttk.Button(bar, text="+ Добавить", command=self._add).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="Сбросить пароль", command=self._reset_pw).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="Удалить", command=self._delete).pack(side="left", padx=6, pady=6)
        self.hash_password = hash_password; self.save_users = save_users
        self.departments = DEPARTMENTS
        self.roles = ["user", "admin"]

    def _reload(self):
        for r in self.tree.get_children(): self.tree.delete(r)
        for u in self.users:
            self.tree.insert("", "end", values=(u.get("username"), u.get("role"), u.get("department")))

    def _add(self):
        win = tk.Toplevel(self); win.title("Новый пользователь"); win.resizable(False, False)
        v_user = tk.StringVar(); v_pw = tk.StringVar()
        v_role = tk.StringVar(value="user"); v_dep = tk.StringVar(value=self.departments[0] if self.departments else "")
        ttk.Label(win, text="Логин").grid(row=0, column=0, padx=6, pady=4, sticky="e")
        ttk.Entry(win, textvariable=v_user).grid(row=0, column=1, padx=6, pady=4)
        ttk.Label(win, text="Пароль").grid(row=1, column=0, padx=6, pady=4, sticky="e")
        ttk.Entry(win, textvariable=v_pw, show="*").grid(row=1, column=1, padx=6, pady=4)
        ttk.Label(win, text="Роль").grid(row=2, column=0, padx=6, pady=4, sticky="e")
        cb_role = ttk.Combobox(win, values=self.roles, textvariable=v_role, state="readonly", width=18)
        cb_role.grid(row=2, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(win, text="Отдел").grid(row=3, column=0, padx=6, pady=4, sticky="e")
        cb_dep = ttk.Combobox(win, values=self.departments, textvariable=v_dep, state="readonly", width=28)
        cb_dep.grid(row=3, column=1, padx=6, pady=4, sticky="w")

        def save_new():
            u = v_user.get().strip(); pw = v_pw.get().strip()
            if not u or not pw: messagebox.showerror("Пользователи","Заполните логин и пароль"); return
            role = v_role.get().strip()
            if role not in self.roles: messagebox.showerror("Пользователи","Роль должна быть 'user' или 'admin'"); return
            dep = v_dep.get().strip()
            if dep not in self.departments: messagebox.showerror("Пользователи","Выберите отдел из списка"); return
            self.users.append({"username": u, "password_hash": self.hash_password(pw), "role": role, "department": dep})
            self.save_users(self.users); self._reload(); win.destroy()

        ttk.Button(win, text="Сохранить", command=save_new).grid(row=4, column=0, columnspan=2, pady=8)
        for i in range(2): win.columnconfigure(i, weight=1)

    def _reset_pw(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Пользователи","Выберите пользователя"); return
        username = self.tree.item(sel[0], "values")[0]
        new_pw = simpledialog.askstring("Сброс пароля", f"Новый пароль для {username}:", show="*")
        if not new_pw: return
        for u in self.users:
            if u.get("username")==username:
                u["password_hash"] = self.hash_password(new_pw)
        self.save_users(self.users); messagebox.showinfo("Пользователи","Пароль обновлен")

    def _delete(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("Пользователи","Выберите пользователя"); return
        username = self.tree.item(sel[0], "values")[0]
        if not messagebox.askyesno("Удаление", f"Удалить пользователя {username}?"): return
        self.users = [u for u in self.users if u.get("username")!=username]
        self.save_users(self.users); self._reload()

class QARequestsWindow(tk.Toplevel):
    def __init__(self, master, needs: Dict[str, Any]):
        super().__init__(master); self.title("Входящие запросы ОУК"); self.geometry("900x420")
        self.needs = needs
        self.tree = ttk.Treeview(self, columns=("request_id","department","need_id","category","item_name","requested_qty","excess_qty","unit","status","created"), show="headings")
        headers = [("request_id","ID",70),("department","Отдел",200),("need_id","План ID",80),("category","Категория",120),
                   ("item_name","Наименование",220),("requested_qty","Запрошено",100),("excess_qty","Сверх плана",100),
                   ("unit","Ед.",60),("status","Статус",100),("created","Дата",100)]
        for k,title,w in headers:
            self.tree.heading(k, text=title); self.tree.column(k, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        # Auto-resize columns to fit window width
        self._store_headers = headers
        self._base_widths = {k: w for (k, _title, w) in headers}
        def _autosize_columns_local(event=None):
            try:
                total_width = max(1, self.tree.winfo_width() - 2)
                cols = [k for (k, _t, _w) in headers]
                base_total = sum(self._base_widths.get(k, 80) for k in cols) or 1
                scale = max(0.5, total_width / base_total)
                for k, _t, base in headers:
                    w = int(base * scale)
                    self.tree.column(k, width=w, stretch=True)
            except Exception:
                pass
                
        self.tree.bind("<Configure>", _autosize_columns_local)
        self.after(100, _autosize_columns_local)
        bar = ttk.Frame(self); bar.pack(fill="x")
        ttk.Button(bar, text="Одобрить", command=self.approve).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="Отклонить", command=self.reject).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="История", command=self.show_history).pack(side="right", padx=6, pady=6)
        self._reload()

    def _reload(self):
        for r in self.tree.get_children(): self.tree.delete(r)
        for r in self.needs.get("qa_overflow_requests", []):
            self.tree.insert("", "end", iid=str(r.get("request_id")), values=(
                r.get("request_id"), r.get("department"), r.get("need_id"), r.get("category"),
                r.get("item_name"), r.get("requested_qty"), r.get("excess_qty"), r.get("unit"),
                r.get("status"), r.get("created")
            ))

    def _find_need(self, department, need_id):
        for n in self.needs.get("departments", {}).get(department, []):
            if int(n.get("need_id"))==int(need_id): return n
        return None

    def approve(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("ОУК","Выберите заявку"); return
        rid = int(self.tree.item(sel[0], "values")[0])
        req = next((x for x in self.needs.get("qa_overflow_requests", []) if int(x.get("request_id"))==rid), None)
        if not req: return
        need = self._find_need(req.get("department"), req.get("need_id"))
        if need:
            need["remaining_qty"] = float(need.get("remaining_qty",0)) + float(req.get("excess_qty",0))
        req["status"] = "approved"
        from storage import save_needs
        save_needs(self.needs)
        self._reload(); messagebox.showinfo("ОУК","Заявка одобрена: остаток по плану увеличен")

    def reject(self):
        sel = self.tree.selection()
        if not sel: messagebox.showinfo("ОУК","Выберите заявку"); return
        rid = int(self.tree.item(sel[0], "values")[0])
        req = next((x for x in self.needs.get("qa_overflow_requests", []) if int(x.get("request_id"))==rid), None)
        if not req: return
        req["status"] = "rejected"
        from storage import save_needs
        save_needs(self.needs)
        self._reload(); messagebox.showinfo("ОУК","Заявка отклонена")

class StoreRequestsWindow(tk.Toplevel):
    def __init__(self, master, app: MainApp):
        super().__init__(master); self.title("Входящие запросы склада"); self.geometry("900x420")
        self.app = app
        self.tree = ttk.Treeview(self, columns=("request_id","department","need_id","item_name","dept_remaining","requested_qty","unit","status","created","requested_by"), show="headings")
        headers = [("request_id","ID",70),("department","Отдел",220),("need_id","План ID",80),
                   ("item_name","Наименование",260),("dept_remaining","Остаток по отделу",150),
                   ("requested_qty","Кол-во",100),("unit","Ед.",60),("status","Статус",120),
                   ("created","Дата",100),("requested_by","Запросил",120)]
        for k,title,w in headers:
            self.tree.heading(k, text=title); self.tree.column(k, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        # Auto-resize columns to fit window width
        self._store_headers = headers
        self._base_widths = {k: w for (k, _title, w) in headers}
        def _autosize_columns_local(event=None):
            try:
                total_width = max(1, self.tree.winfo_width() - 2)
                cols = [k for (k, _t, _w) in headers]
                base_total = sum(self._base_widths.get(k, 80) for k in cols) or 1
                scale = max(0.5, total_width / base_total)
                for k, _t, base in headers:
                    w = int(base * scale)
                    self.tree.column(k, width=w, stretch=True)
            except Exception:
                pass
                
        self.tree.bind("<Configure>", _autosize_columns_local)
        self.after(100, _autosize_columns_local)
        bar = ttk.Frame(self); bar.pack(fill="x")
        ttk.Button(bar, text="Выдать", command=self.approve).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="Отклонить", command=self.reject).pack(side="left", padx=6, pady=6)
        ttk.Button(bar, text="История", command=self.show_history).pack(side="right", padx=6, pady=6)
        self._reload()

    def _reload(self):
        for r in self.tree.get_children(): self.tree.delete(r)
        for r in self.app.needs.get("store_requests", []):
            dep = r.get("department"); nid = r.get("need_id")
            need = None
            for n in self.app.needs.get("departments", {}).get(dep, []):
                if int(n.get("need_id")) == int(nid):
                    need = n; break
            item_name = (need.get("item_name") if need else "")
            dept_rem = (need.get("remaining_qty") if need else "")
            self.tree.insert("", "end", iid=str(r.get("request_id")), values=(
                r.get("request_id"), dep, nid,
                item_name, dept_rem,
                r.get("requested_qty"), r.get("unit"), r.get("status"),
                r.get("created"), r.get("requested_by")
            ))

    def _get_selected_request(self):
        sel = self.tree.selection()
        if not sel: return None
        rid = int(self.tree.item(sel[0], "values")[0])
        req = next((x for x in self.app.needs.get("store_requests", []) if int(x.get("request_id"))==rid), None)
        return req

    def approve(self):
        req = self._get_selected_request()
        if not req: messagebox.showinfo("Склад","Выберите заявку"); return
        if req.get("status") != "pending":
            messagebox.showinfo("Склад","Эта заявка уже обработана"); return
        res = self.app._process_issue(req.get("department"), int(req.get("need_id")), float(req.get("requested_qty")))
        req["status"] = "done" if res=="Выдача выполнена" else "redirected"
        from storage import save_needs
        save_needs(self.app.needs)
        self._reload(); messagebox.showinfo("Склад", res)

    def reject(self):
        req = self._get_selected_request()
        if not req: messagebox.showinfo("Склад","Выберите заявку"); return
        req["status"] = "rejected"
        from storage import save_needs
        save_needs(self.app.needs)
        self._reload(); messagebox.showinfo("Склад","Заявка отклонена")

    def show_history(self):
        StoreRequestsHistoryWindow(self, self.app)



def _autosize_columns(self, event=None):
    try:
        total_width = max(1, self.tree.winfo_width() - 2)
        cols = [k for (k, _t, _w) in self._store_headers]
        base_total = sum(self._base_widths.get(k, 80) for k in cols) or 1
        scale = max(0.5, total_width / base_total)
        for k, _t, base in self._store_headers:
            w = int(base * scale)
            self.tree.column(k, width=w, stretch=True)
    except Exception:
        pass


class StoreRequestsHistoryWindow(tk.Toplevel):
    def __init__(self, master, app: MainApp):
        super().__init__(master)
        self.title("История запросов склада")
        self.geometry("1050x500")
        self.app = app

        # Toolbar with filters and search
        bar = ttk.Frame(self); bar.pack(fill="x", padx=8, pady=6)
        ttk.Label(bar, text="Отдел:").pack(side="left")
        self.var_dept = tk.StringVar(value="Все")
        depts = ["Все"] + sorted({r.get("department") for r in self.app.needs.get("store_requests", [])})
        self.cb_dept = ttk.Combobox(bar, values=depts, textvariable=self.var_dept, state="readonly", width=30)
        self.cb_dept.pack(side="left", padx=(4,12))
        ttk.Label(bar, text="Статус:").pack(side="left")
        self.var_status = tk.StringVar(value="Любой")
        self.cb_status = ttk.Combobox(bar, values=["Любой","pending","done","rejected","redirected"], textvariable=self.var_status, state="readonly", width=12)
        self.cb_status.pack(side="left", padx=(4,12))
        ttk.Label(bar, text="Поиск по наименованию:").pack(side="left")
        self.var_query = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.var_query, width=34); ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self._reload())

        ttk.Button(bar, text="Сбросить", command=self._reset_filters).pack(side="right", padx=6)

        # Tree
        cols = ("request_id","department","need_id","item_name","requested_qty","unit","status","created","requested_by")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        headers = [("request_id","ID",70),("department","Отдел",220),("need_id","План ID",80),
                   ("item_name","Наименование",300),("requested_qty","Кол-во",90),("unit","Ед.",60),
                   ("status","Статус",100),("created","Дата",100),("requested_by","Запросил",120)]
        for k,title,w in headers:
            self.tree.heading(k, text=title, command=lambda c=k: self._sort_by(c, False))
            self.tree.column(k, width=w, anchor="w")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=(0,8))

        # Events
        self.cb_dept.bind("<<ComboboxSelected>>", lambda e: self._reload())
        self.cb_status.bind("<<ComboboxSelected>>", lambda e: self._reload())

        self._reload()

    def _reset_filters(self):
        self.var_dept.set("Все"); self.var_status.set("Любой"); self.var_query.set("")
        self._reload()

    def _sort_by(self, col, descending):
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]
        try:
            # Try numeric
            data.sort(key=lambda t: float(t[0]), reverse=descending)
        except ValueError:
            # Fallback lexicographic
            data.sort(key=lambda t: str(t[0]).lower(), reverse=descending)
        for idx, item in enumerate(data):
            self.tree.move(item[1], "", idx)
        self.tree.heading(col, command=lambda: self._sort_by(col, not descending))

    def _reload(self):
        # Clear
        for r in self.tree.get_children(): self.tree.delete(r)

        dept_filter = self.var_dept.get()
        status_filter = self.var_status.get()
        q = (self.var_query.get() or "").strip().lower()

        for r in self.app.needs.get("store_requests", []):
            # Filters
            if dept_filter != "Все" and r.get("department") != dept_filter:
                continue
            if status_filter != "Любой" and r.get("status") != status_filter:
                continue

            # Resolve item name from plan
            dep = r.get("department"); nid = r.get("need_id")
            need = None
            for n in self.app.needs.get("departments", {}).get(dep, []):
                if int(n.get("need_id")) == int(nid):
                    need = n; break
            item_name = (need.get("item_name") if need else "")

            if q and q not in str(item_name).lower():
                continue

            self.tree.insert("", "end", values=(
                r.get("request_id"), r.get("department"), r.get("need_id"),
                item_name, r.get("requested_qty"), r.get("unit"),
                r.get("status"), r.get("created"), r.get("requested_by")
            ))

class ItemDialog(tk.Toplevel):
    def __init__(self, master, title: str, on_save=None, item: Optional[Item] = None, default_responsible: Optional[str] = None):
        super().__init__(master); self.title(title); self.resizable(False, False)
        self.on_save = on_save
        self.item = item
        self.var_category = tk.StringVar(value=item.category if item else "Реактивы")
        self.var_name = tk.StringVar(value=item.name if item else "")
        self.var_qty = tk.StringVar(value=str(item.quantity) if item else "0")
        self.var_unit = tk.StringVar(value=item.unit if item else (UNITS[0] if UNITS else ""))
        self.var_pack = tk.StringVar(value=item.packaging if item else "")
        self.var_storage = tk.StringVar(value=item.storage_place if item else "")
        self.var_expiry = tk.StringVar(value=item.expiry_date or "" if item else "")
        self.var_received = tk.StringVar(value=item.date_received or "" if item else "")
        self.var_batch = tk.StringVar(value=item.batch_number if item else "")
        self.var_resp = tk.StringVar(value=(item.responsible if item else (default_responsible or "")))
        self.var_qual = tk.StringVar(value=item.qualification or "" if item else "")
        self.var_type = tk.StringVar(value=item.reagent_type or "" if item else "")
        self.var_regno = tk.StringVar(value=item.state_register_no or "" if item else "")
        self.var_certv = tk.StringVar(value=item.certified_value or "" if item else "")
        self.var_storage_cond = tk.StringVar(value=item.storage_conditions or "" if item else "")

        frm = ttk.Frame(self, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Категория").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        cb_cat = ttk.Combobox(frm, values=CATEGORIES, textvariable=self.var_category, state="readonly", width=28)
        cb_cat.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        cb_cat.bind("<<ComboboxSelected>>", lambda e: self._update_cat_fields())

        ttk.Label(frm, text="Наименование").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_name, width=45).grid(row=0, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Количество").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_qty, width=12).grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Ед. изм.").grid(row=1, column=2, sticky="e", padx=6, pady=4)
        ttk.Combobox(frm, values=UNITS, textvariable=self.var_unit, width=12, state="readonly").grid(row=1, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Фасовка").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_pack, width=20).grid(row=2, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Место хранения").grid(row=2, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_storage, width=30).grid(row=2, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Срок годности").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self.de_expiry = DateEntry(frm, date_pattern="yyyy-mm-dd", width=18)
        self.de_expiry.grid(row=3, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Дата поступления").grid(row=3, column=2, sticky="e", padx=6, pady=4)
        self.de_received = DateEntry(frm, date_pattern="yyyy-mm-dd", width=18)
        self.de_received.grid(row=3, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Номер партии").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_batch, width=20).grid(row=4, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(frm, text="Ответственный").grid(row=4, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=self.var_resp, width=30, state="readonly").grid(row=4, column=3, sticky="w", padx=6, pady=4)

        self.frm_reagents = ttk.Labelframe(frm, text="Реактивы", padding=8)
        ttk.Label(self.frm_reagents, text="Квалификация").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Combobox(self.frm_reagents, values=REAGENT_QUALIFICATIONS, textvariable=self.var_qual, width=20, state="readonly").grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(self.frm_reagents, text="Тип реактива").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        ttk.Combobox(self.frm_reagents, values=REAGENT_TYPES, textvariable=self.var_type, width=20, state="readonly").grid(row=0, column=3, sticky="w", padx=6, pady=4)

        self.frm_gso = ttk.Labelframe(frm, text="ГСО-ПГС-СО", padding=8)
        ttk.Label(self.frm_gso, text="№ в госреестре СО").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_regno, width=20).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(self.frm_gso, text="Аттестованное значение").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_certv, width=20).grid(row=0, column=3, sticky="w", padx=6, pady=4)

        ttk.Label(self.frm_gso, text="Условия хранения").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.frm_gso, textvariable=self.var_storage_cond, width=40).grid(row=1, column=1, columnspan=3, sticky="w", padx=6, pady=4)

        self.frm_reagents.grid(row=5, column=0, columnspan=4, sticky="ew", padx=6, pady=8)
        self.frm_gso.grid(row=6, column=0, columnspan=4, sticky="ew", padx=6, pady=8)

        btns = ttk.Frame(frm)
        btns.grid(row=7, column=0, columnspan=4, sticky="e", pady=8)
        ttk.Button(btns, text="Сохранить", command=self._save).pack(side="right", padx=6)
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right", padx=6)

        for i in range(4):
            frm.columnconfigure(i, weight=1)

        self._update_cat_fields()

    def _update_cat_fields(self):
        cat = self.var_category.get()
        if cat == "Реактивы":
            self.frm_reagents.grid()
            self.frm_gso.grid_remove()
        elif cat == "ГСО-ПГС-СО":
            self.frm_gso.grid()
            self.frm_reagents.grid_remove()
        else:
            self.frm_reagents.grid_remove()
            self.frm_gso.grid_remove()

    def _save(self):
        try:
            qty = float(self.var_qty.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Сохранение", "Количество должно быть числом")
            return
        payload = {
            "seq_id": self.item.seq_id if self.item else None,
            "name": self.var_name.get().strip(),
            "category": self.var_category.get(),
            "quantity": qty,
            "unit": self.var_unit.get(),
            "packaging": self.var_pack.get().strip(),
            "storage_place": self.var_storage.get().strip(),
            "expiry_date": self.de_expiry.get_date().strftime("%Y-%m-%d"),
            "date_received": self.de_received.get_date().strftime("%Y-%m-%d"),
            "batch_number": self.var_batch.get().strip(),
            "responsible": self.var_resp.get().strip(),
            "qualification": self.var_qual.get().strip() or None,
            "reagent_type": self.var_type.get().strip() or None,
            "state_register_no": self.var_regno.get().strip() or None,
            "certified_value": self.var_certv.get().strip() or None,
            "manufacture_date": None,
            "manufacturer": None,
            "storage_conditions": self.var_storage_cond.get().strip() or None,
        }
        if not payload["name"]:
            messagebox.showerror("Сохранение", "Введите наименование")
            return
        if self.on_save:
            self.on_save(payload)
        self.destroy()
