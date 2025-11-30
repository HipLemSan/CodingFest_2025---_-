import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime

OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

DATA_FILE = "warehouse_data.json"
COLUMNS = [
    "Дата", "Вид материала", "Размер катушки, вес кг.",
    "Сечение", "Цвет", "Условия хранения", "Статус", "Остаток"
]

IDENTIFYING_FIELDS = [
    "Дата", "Вид материала", "Размер катушки, вес кг.",
    "Сечение", "Цвет", "Условия хранения"
]
FILTER_FIELDS = [
    "Вид материала", "Размер катушки, вес кг.", "Сечение",
    "Цвет", "Статус", "Остаток"
]

def ensure_data():
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump([], f)

def load_data():
    ensure_data()
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def parse_weight(weight_str):
    try:
        clean = weight_str.replace("кг", "").replace("г", "").replace(",", ".").strip()
        if " " in clean:
            clean = clean.split()[0]
        return float(clean) if clean else 0.0
    except:
        return 0.0

def format_weight(kg):
    return f"{kg:.1f} кг"

class WarehouseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система учёта расходных материалов на складе")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)

        self.data = load_data()
        self.filtered = self.data.copy()
        self.create_ui()

    def create_ui(self):
        header = tk.Label(
            self.root,
            text="Склад 3D-печати — Учёт расходных материалов",
            font=("Segoe UI", 16, "bold"),
            bg="#2c3e50", fg="white", pady=12
        )
        header.pack(fill="x")

        filter_frame = tk.Frame(self.root, bg="#f5f6fa", pady=10)
        filter_frame.pack(fill="x", padx=15)

        ttk.Label(filter_frame, text="Фильтр по полю:", background="#f5f6fa").pack(side="left")
        self.field_var = tk.StringVar(value=FILTER_FIELDS[0])
        self.field_combo = ttk.Combobox(filter_frame, textvariable=self.field_var, values=FILTER_FIELDS, state="readonly", width=20)
        self.field_combo.pack(side="left", padx=5)

        ttk.Label(filter_frame, text="Значение:", background="#f5f6fa").pack(side="left", padx=(10, 5))
        self.value_entry = ttk.Entry(filter_frame, width=20)
        self.value_entry.pack(side="left", padx=5)

        ttk.Button(filter_frame, text="🔍 Применить", command=self.apply_filter).pack(side="left", padx=5)
        ttk.Button(filter_frame, text="❌ Сбросить", command=self.reset_filter).pack(side="left", padx=5)

        ttk.Label(filter_frame, text="  Быстрый поиск:", background="#f5f6fa").pack(side="left", padx=(20, 5))
        self.search_all = ttk.Entry(filter_frame, width=15)
        self.search_all.pack(side="left")
        self.search_all.bind("<KeyRelease>", lambda e: self.apply_global_search())

        table_frame = tk.Frame(self.root)
        table_frame.pack(pady=5, padx=15, fill="both", expand=True)

        self.tree = ttk.Treeview(table_frame, columns=COLUMNS, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#d4d6d9")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26)

        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor="w")

        btn_frame = tk.Frame(self.root, pady=10)
        btn_frame.pack()

        ttk.Button(btn_frame, text="➕ Добавить", command=self.add_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="✏️ Редактировать", command=self.edit_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🗑 Удалить", command=self.delete_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="➖ Использовать 100 г", command=self.use_100g).pack(side="left", padx=5)
        if OPENPYXL_AVAILABLE:
            ttk.Button(btn_frame, text="📊 Экспорт в Excel", command=self.export_to_excel).pack(side="left", padx=5)
        else:
            ttk.Button(btn_frame, text="📊 Экспорт в Excel (установите openpyxl)", state="disabled").pack(side="left", padx=5)

        self.status_label = tk.Label(self.root, text="", bd=1, relief="sunken", anchor="w", bg="#ecf0f1")
        self.status_label.pack(side="bottom", fill="x")

        self.refresh_table()
        self.update_status()

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.filtered:
            self.tree.insert("", "end", values=[row.get(col, "") for col in COLUMNS])

    def update_status(self):
        total_weight = sum(parse_weight(item.get("Остаток", "0")) for item in self.data)
        self.status_label.config(text=f"Всего записей: {len(self.data)} | Общий остаток: {total_weight:.1f} кг")

    def apply_filter(self):
        field = self.field_var.get()
        value = self.value_entry.get().strip().lower()
        if not value:
            self.filtered = self.data.copy()
        else:
            self.filtered = [
                item for item in self.data
                if value in str(item.get(field, "")).lower()
            ]
        self.refresh_table()

    def apply_global_search(self):
        query = self.search_all.get().strip().lower()
        if not query:
            self.reset_filter()
            return
        self.filtered = [
            item for item in self.data
            if any(query in str(v).lower() for v in item.values())
        ]
        self.refresh_table()

    def reset_filter(self):
        self.value_entry.delete(0, "end")
        self.search_all.delete(0, "end")
        self.filtered = self.data.copy()
        self.refresh_table()

    def get_selected_item(self):
        sel = self.tree.focus()
        if not sel:
            return None, None
        values = self.tree.item(sel, "values")
        return dict(zip(COLUMNS, values)), sel

    def save_item(self, item):
        self.data.append(item)
        save_data(self.data)
        self.reset_filter()
        self.update_status()

    def update_item(self, old_item, new_item):
        for i, item in enumerate(self.data):
            if all(str(item.get(k, "")) == str(old_item.get(k, "")) for k in COLUMNS):
                self.data[i] = new_item
                save_data(self.data)
                self.reset_filter()
                self.update_status()
                return
        self.data.append(new_item)
        save_data(self.data)
        self.reset_filter()
        self.update_status()

    def add_item(self):
        dialog = EditDialog(self.root, "Добавить материал", self.save_item)

    def edit_item(self):
        item, _ = self.get_selected_item()
        if not item:
            messagebox.showinfo("Инфо", "Выберите запись")
            return
        EditDialog(self.root, "Редактировать материал", lambda new: self.update_item(item, new), item)

    def delete_item(self):
        item, _ = self.get_selected_item()
        if not item:
            messagebox.showinfo("Инфо", "Выберите запись")
            return
        for i, d in enumerate(self.data):
            if all(str(d.get(k, "")) == str(item.get(k, "")) for k in COLUMNS):
                self.data.pop(i)
                save_data(self.data)
                self.reset_filter()
                self.update_status()
                return
        messagebox.showerror("Ошибка", "Не удалось найти запись для удаления")

    def use_100g(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showinfo("Инфо", "Выберите запись")
            return

        selected_values = self.tree.item(sel, "values")
        selected_dict = dict(zip(COLUMNS, selected_values))

        original_item = None
        for item in self.data:
            match = True
            for field in IDENTIFYING_FIELDS:
                if str(item.get(field, "")) != str(selected_dict.get(field, "")):
                    match = False
                    break
            if match:
                original_item = item
                break

        if not original_item:
            messagebox.showerror("Ошибка", "Не удалось найти запись для обновления")
            return

        current = parse_weight(original_item["Остаток"])
        if current <= 0.1:
            original_item["Остаток"] = "0 кг"
            original_item["Статус"] = "Израсходован"
        else:
            original_item["Остаток"] = format_weight(current - 0.1)
            original_item["Статус"] = "Используется"

        save_data(self.data)
        self.reset_filter()
        self.update_status()

    def export_to_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Ошибка", "Установите openpyxl: pip install openpyxl")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")]
        )
        if not filepath:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.append(COLUMNS)
            for item in self.data:
                ws.append([item.get(col, "") for col in COLUMNS])
            wb.save(filepath)
            messagebox.showinfo("Успех", f"Данные экспортированы в:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить Excel:\n{str(e)}")


class EditDialog(tk.Toplevel):
    def __init__(self, parent, title, on_save, item=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("520x540")
        self.transient(parent)
        self.grab_set()
        self.on_save = on_save

        self.entries = {}
        for col in COLUMNS:
            frame = tk.Frame(self)
            frame.pack(fill="x", padx=20, pady=4)
            tk.Label(frame, text=col + ":", font=("Segoe UI", 10)).pack(anchor="w")
            if col == "Статус":
                entry = ttk.Combobox(frame, values=["Добавлен", "Используется", "Израсходован"], width=60)
            else:
                entry = ttk.Entry(frame, width=62)
            entry.pack(fill="x", pady=1)
            self.entries[col] = entry

            val = item.get(col, "") if item else ""
            if col == "Дата" and not val:
                val = datetime.now().strftime("%d.%m.%Y")
            entry.insert(0, val)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Сохранить", command=self.save,
                  bg="#27ae60", fg="white", font=("Segoe UI", 10), width=12).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Отмена", command=self.destroy,
                  bg="#95a5a6", fg="white", font=("Segoe UI", 10), width=12).pack(side="left")

    def save(self):
        item = {col: self.entries[col].get().strip() for col in COLUMNS}
        if not all([item.get("Дата"), item.get("Вид материала"), item.get("Остаток")]):
            messagebox.showerror("Ошибка", "Заполните: Дата, Вид материала, Остаток")
            return
        self.on_save(item)
        self.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WarehouseApp(root)
    root.mainloop()
