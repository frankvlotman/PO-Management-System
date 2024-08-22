import tkinter as tk
from tkinter import messagebox, ttk, font
import sqlite3
import os
from datetime import datetime
import win32com.client as win32

# Constants
DB_PATH = r'C:\Users\Frank\Desktop\purchase_orders.db'

# Database Setup
def create_tables():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        email TEXT NOT NULL
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS items (
        id INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        price REAL NOT NULL
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS purchase_orders (
        id INTEGER PRIMARY KEY,
        supplier_id INTEGER NOT NULL,
        date TEXT NOT NULL,
        status TEXT NOT NULL,
        FOREIGN KEY (supplier_id) REFERENCES suppliers (id)
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS purchase_order_items (
        id INTEGER PRIMARY KEY,
        purchase_order_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        quantity INTEGER NOT NULL,
        FOREIGN KEY (purchase_order_id) REFERENCES purchase_orders (id),
        FOREIGN KEY (item_id) REFERENCES items (id)
    )
    ''')

    conn.commit()
    conn.close()

# Email Integration
def send_email(to, subject, body, attachment=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    if attachment:
        mail.Attachments.Add(attachment)
    mail.Send()

# Data Operations
def add_supplier(name, email):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('INSERT INTO suppliers (name, email) VALUES (?, ?)', (name, email))
    conn.commit()
    conn.close()

def update_supplier(supplier_id, name, email):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('UPDATE suppliers SET name=?, email=? WHERE id=?', (name, email, supplier_id))
    conn.commit()
    conn.close()

def delete_supplier(supplier_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM suppliers WHERE id=?', (supplier_id,))
    conn.commit()
    conn.close()

def add_item(name, price):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('INSERT INTO items (name, price) VALUES (?, ?)', (name, price))
    conn.commit()
    conn.close()

def update_item(item_id, name, price):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('UPDATE items SET name=?, price=? WHERE id=?', (name, price, item_id))
    conn.commit()
    conn.close()

def delete_item(item_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM items WHERE id=?', (item_id,))
    conn.commit()
    conn.close()

def create_purchase_order(supplier_id, items):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    date = datetime.now().strftime('%Y-%m-%d')
    status = 'Pending'
    cursor.execute('INSERT INTO purchase_orders (supplier_id, date, status) VALUES (?, ?, ?)', (supplier_id, date, status))
    po_id = cursor.lastrowid

    for item_id, quantity in items:
        cursor.execute('INSERT INTO purchase_order_items (purchase_order_id, item_id, quantity) VALUES (?, ?, ?)', (po_id, item_id, quantity))

    cursor.execute('SELECT email FROM suppliers WHERE id=?', (supplier_id,))
    supplier_email = cursor.fetchone()[0]

    conn.commit()
    conn.close()

    send_email(supplier_email, 'New Purchase Order', f'You have a new purchase order #{po_id}')
    return po_id

def update_purchase_order_status(po_id, status):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('UPDATE purchase_orders SET status=? WHERE id=?', (status, po_id))
    conn.commit()
    conn.close()

def delete_purchase_order(po_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM purchase_order_items WHERE purchase_order_id=?', (po_id,))
    cursor.execute('DELETE FROM purchase_orders WHERE id=?', (po_id,))
    conn.commit()
    conn.close()

def fetch_suppliers():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT id, name FROM suppliers')
    suppliers = cursor.fetchall()
    conn.close()
    return suppliers

def get_table_names():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    conn.close()
    return [table[0] for table in tables]

def fetch_table_data(table_name):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM {table_name}")
    columns = [description[0] for description in cursor.description]
    rows = cursor.fetchall()
    conn.close()
    return columns, rows

# Tkinter GUI
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Order Management System")
        self.root.geometry("800x600")

        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.setup_main_menu()

    def setup_main_menu(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="Purchase Order Management System", font=("Arial", 16)).pack(pady=10)

        tk.Button(self.main_frame, text="Add Supplier", command=self.add_supplier_window, width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Add Item", command=self.add_item_window, width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Create Purchase Order", command=self.create_purchase_order_window, width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Manage Records", command=self.manage_records_window, width=30).pack(pady=5)
        tk.Button(self.main_frame, text="View Tables", command=self.view_tables_window, width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Exit", command=self.root.quit, width=30).pack(pady=5)

    def clear_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def add_supplier_window(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="Add Supplier", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Name").pack(pady=5)
        name_entry = tk.Entry(self.main_frame)
        name_entry.pack(pady=5)

        tk.Label(self.main_frame, text="Email").pack(pady=5)
        email_entry = tk.Entry(self.main_frame)
        email_entry.pack(pady=5)

        def add_supplier_action():
            name = name_entry.get()
            email = email_entry.get()
            add_supplier(name, email)
            messagebox.showinfo("Success", "Supplier added successfully")
            self.setup_main_menu()

        tk.Button(self.main_frame, text="Add Supplier", command=add_supplier_action).pack(pady=20)

    def add_item_window(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="Add Item", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Name").pack(pady=5)
        name_entry = tk.Entry(self.main_frame)
        name_entry.pack(pady=5)

        tk.Label(self.main_frame, text="Price").pack(pady=5)
        price_entry = tk.Entry(self.main_frame)
        price_entry.pack(pady=5)

        def add_item_action():
            name = name_entry.get()
            price = float(price_entry.get())
            add_item(name, price)
            messagebox.showinfo("Success", "Item added successfully")
            self.setup_main_menu()

        tk.Button(self.main_frame, text="Add Item", command=add_item_action).pack(pady=20)

    def create_purchase_order_window(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="Create Purchase Order", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Select Supplier").pack(pady=5)
        suppliers = fetch_suppliers()
        supplier_options = [f"{supplier[1]} (ID: {supplier[0]})" for supplier in suppliers]
        selected_supplier = tk.StringVar(self.main_frame)
        selected_supplier.set(supplier_options[0])
        supplier_menu = tk.OptionMenu(self.main_frame, selected_supplier, *supplier_options)
        supplier_menu.pack(pady=5)

        items = []

        def add_item_to_po():
            try:
                item_id = int(item_id_entry.get())
                quantity = int(quantity_entry.get())
                items.append((item_id, quantity))
                item_listbox.insert(tk.END, f"Item ID: {item_id}, Quantity: {quantity}")
                item_id_entry.delete(0, tk.END)
                quantity_entry.delete(0, tk.END)
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid item ID and quantity.")

        tk.Label(self.main_frame, text="Item ID").pack(pady=5)
        item_id_entry = tk.Entry(self.main_frame)
        item_id_entry.pack(pady=5)

        tk.Label(self.main_frame, text="Quantity").pack(pady=5)
        quantity_entry = tk.Entry(self.main_frame)
        quantity_entry.pack(pady=5)

        tk.Button(self.main_frame, text="Add Item to PO", command=add_item_to_po).pack(pady=10)

        item_listbox = tk.Listbox(self.main_frame)
        item_listbox.pack(pady=5)

        def create_po_action():
            supplier_id = int(selected_supplier.get().split("ID: ")[1].strip(")"))
            if not items:
                messagebox.showwarning("No items", "Please add at least one item to the purchase order.")
                return
            po_id = create_purchase_order(supplier_id, items)
            messagebox.showinfo("Success", f"Purchase order #{po_id} created")
            self.setup_main_menu()

        tk.Button(self.main_frame, text="Create Purchase Order", command=create_po_action).pack(pady=20)

    def manage_records_window(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="Manage Records", font=("Arial", 14)).pack(pady=10)

        tk.Button(self.main_frame, text="Suppliers", command=lambda: self.manage_table("suppliers"), width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Items", command=lambda: self.manage_table("items"), width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Purchase Orders", command=lambda: self.manage_table("purchase_orders"), width=30).pack(pady=5)
        tk.Button(self.main_frame, text="Back to Menu", command=self.setup_main_menu, width=30).pack(pady=5)

    def manage_table(self, table_name):
        self.clear_frame()

        tk.Label(self.main_frame, text=f"Manage {table_name.capitalize()}", font=("Arial", 14)).pack(pady=10)

        columns, rows = fetch_table_data(table_name)

        tree = ttk.Treeview(self.main_frame, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=font.Font().measure(col))

        for row in rows:
            tree.insert("", tk.END, values=row)
            for ix, val in enumerate(row):
                col_w = font.Font().measure(str(val))
                if tree.column(columns[ix], width=None) < col_w:
                    tree.column(columns[ix], width=col_w)

        tree.pack(fill=tk.BOTH, expand=True)

        def delete_record():
            selected_item = tree.selection()[0]
            record_id = tree.item(selected_item)['values'][0]
            if table_name == "suppliers":
                delete_supplier(record_id)
            elif table_name == "items":
                delete_item(record_id)
            elif table_name == "purchase_orders":
                delete_purchase_order(record_id)
            tree.delete(selected_item)
            messagebox.showinfo("Success", "Record deleted successfully")

        def update_record():
            selected_item = tree.selection()[0]
            record_id = tree.item(selected_item)['values'][0]
            if table_name == "suppliers":
                self.update_supplier_window(record_id)
            elif table_name == "items":
                self.update_item_window(record_id)
            elif table_name == "purchase_orders":
                self.update_po_status_window(record_id)

        tk.Button(self.main_frame, text="Delete Record", command=delete_record).pack(pady=10)
        tk.Button(self.main_frame, text="Update Record", command=update_record).pack(pady=10)
        tk.Button(self.main_frame, text="Back to Manage Records", command=self.manage_records_window).pack(pady=10)

    def update_supplier_window(self, supplier_id):
        self.clear_frame()

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT name, email FROM suppliers WHERE id=?', (supplier_id,))
        supplier = cursor.fetchone()
        conn.close()

        tk.Label(self.main_frame, text="Update Supplier", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Name").pack(pady=5)
        name_entry = tk.Entry(self.main_frame)
        name_entry.pack(pady=5)
        name_entry.insert(0, supplier[0])

        tk.Label(self.main_frame, text="Email").pack(pady=5)
        email_entry = tk.Entry(self.main_frame)
        email_entry.pack(pady=5)
        email_entry.insert(0, supplier[1])

        def update_supplier_action():
            name = name_entry.get()
            email = email_entry.get()
            update_supplier(supplier_id, name, email)
            messagebox.showinfo("Success", "Supplier updated successfully")
            self.manage_table("suppliers")

        tk.Button(self.main_frame, text="Update Supplier", command=update_supplier_action).pack(pady=20)

    def update_item_window(self, item_id):
        self.clear_frame()

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT name, price FROM items WHERE id=?', (item_id,))
        item = cursor.fetchone()
        conn.close()

        tk.Label(self.main_frame, text="Update Item", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Name").pack(pady=5)
        name_entry = tk.Entry(self.main_frame)
        name_entry.pack(pady=5)
        name_entry.insert(0, item[0])

        tk.Label(self.main_frame, text="Price").pack(pady=5)
        price_entry = tk.Entry(self.main_frame)
        price_entry.pack(pady=5)
        price_entry.insert(0, item[1])

        def update_item_action():
            name = name_entry.get()
            price = float(price_entry.get())
            update_item(item_id, name, price)
            messagebox.showinfo("Success", "Item updated successfully")
            self.manage_table("items")

        tk.Button(self.main_frame, text="Update Item", command=update_item_action).pack(pady=20)

    def update_po_status_window(self, po_id):
        self.clear_frame()

        tk.Label(self.main_frame, text="Update Purchase Order Status", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.main_frame, text="Status").pack(pady=5)
        status_entry = tk.Entry(self.main_frame)
        status_entry.pack(pady=5)

        def update_po_status_action():
            status = status_entry.get()
            update_purchase_order_status(po_id, status)
            messagebox.showinfo("Success", "Purchase order status updated successfully")
            self.manage_table("purchase_orders")

        tk.Button(self.main_frame, text="Update Status", command=update_po_status_action).pack(pady=20)

    def view_tables_window(self):
        self.clear_frame()

        tk.Label(self.main_frame, text="View Tables", font=("Arial", 14)).pack(pady=10)

        tables = get_table_names()
        table_listbox = tk.Listbox(self.main_frame)
        for table in tables:
            table_listbox.insert(tk.END, table)
        table_listbox.pack(pady=5)

        def view_table_action():
            selected_table = table_listbox.get(table_listbox.curselection())
            self.view_table_content(selected_table)

        tk.Button(self.main_frame, text="View Table", command=view_table_action).pack(pady=20)
        tk.Button(self.main_frame, text="Back to Menu", command=self.setup_main_menu).pack(pady=10)

    def view_table_content(self, table_name):
        self.clear_frame()

        tk.Label(self.main_frame, text=f"Viewing Table: {table_name}", font=("Arial", 14)).pack(pady=10)

        columns, rows = fetch_table_data(table_name)

        tree = ttk.Treeview(self.main_frame, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=font.Font().measure(col))

        for row in rows:
            tree.insert("", tk.END, values=row)
            for ix, val in enumerate(row):
                col_w = font.Font().measure(str(val))
                if tree.column(columns[ix], width=None) < col_w:
                    tree.column(columns[ix], width=col_w)

        tree.pack(fill=tk.BOTH, expand=True)

        tk.Button(self.main_frame, text="Back to Tables", command=self.view_tables_window).pack(pady=10)

if __name__ == '__main__':
    create_tables()
    
    root = tk.Tk()
    app = App(root)
    root.mainloop()
