# ui_components.py
import tkinter as tk
from tkinter import ttk
from config import COLUMN_WIDTHS

class DataTable:
    def __init__(self, parent, on_double_click):
        self.parent = parent
        self.on_double_click = on_double_click
        self.tree = None
        self._create_table()
    
    def _create_table(self):
        columns = ("ID", "Type", "Layer", "Size (WxH/Dia)", "Coordinate (X,Y)", "Area", "Code")
        self.tree = ttk.Treeview(self.parent, columns=columns, show='headings', height=20)

        # Define column headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.CENTER, width=COLUMN_WIDTHS.get(col, 120))

        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5, side=tk.LEFT)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)

        # Double-click event
        self.tree.bind("<Double-1>", self.on_double_click)
    
    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def add_item(self, values):
        self.tree.insert("", tk.END, values=values)
    
    def get_selected(self):
        selection = self.tree.selection()
        if selection:
            return self.tree.item(selection[0], "values")
        return None
    
    def get_all_items(self):
        items = []
        for child in self.tree.get_children():
            items.append(self.tree.item(child)["values"])
        return items

class FilterFrame:
    def __init__(self, parent, on_filter_change):
        self.parent = parent
        self.on_filter_change = on_filter_change
        self.show_rect = tk.BooleanVar(value=True)
        self.show_circle = tk.BooleanVar(value=True)
        self._create_filter_frame()
    
    def _create_filter_frame(self):
        filter_frame = tk.LabelFrame(self.parent, text="Filter", padx=10, pady=5)
        filter_frame.pack(side=tk.RIGHT, padx=10)
        
        tk.Checkbutton(filter_frame, text="Show Rectangles", variable=self.show_rect, 
                      command=self.on_filter_change).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(filter_frame, text="Show Circles", variable=self.show_circle,
                      command=self.on_filter_change).pack(side=tk.LEFT, padx=5)
    
    def get_filters(self):
        return {
            'show_rect': self.show_rect.get(),
            'show_circle': self.show_circle.get()
        }

class ButtonPanel:
    def __init__(self, parent, on_load_dxf, on_connect_cad, on_export, on_remove_duplicates=None):
        self.parent = parent
        self._create_buttons(on_load_dxf, on_connect_cad, on_export, on_remove_duplicates)

    def _create_buttons(self, on_load_dxf, on_connect_cad, on_export, on_remove_duplicates=None):
        btn_frame = tk.Frame(self.parent, pady=10)
        btn_frame.pack(side=tk.TOP, fill=tk.X)

        btn_frame1 = tk.Frame(btn_frame)
        btn_frame1.pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame1, text="📂 1. Load DXF File", command=on_load_dxf, 
                 width=20, bg="#e1e1e1").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="🔗 2. Connect AutoCAD", command=on_connect_cad, 
                 width=20, bg="#d4edda").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="📊 3. Export Excel", command=on_export, 
                 width=15).pack(side=tk.LEFT, padx=5)
        if on_remove_duplicates:
            tk.Button(btn_frame1, text="🧹 Remove Duplicates", command=on_remove_duplicates, 
                     width=18, bg="#ffe082").pack(side=tk.LEFT, padx=5)

class StatusBar:
    def __init__(self, parent):
        self.status_var = tk.StringVar(value="Status: Ready")
        status_bar = tk.Label(parent, textvariable=self.status_var, bd=1, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def set_status(self, message):
        self.status_var.set(message)