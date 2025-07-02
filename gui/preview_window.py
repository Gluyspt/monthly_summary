import customtkinter as ctk
from tkinter import ttk, messagebox
import numpy as np

class PreviewWindow(ctk.CTkToplevel):
    def __init__(self, parent, mode, dataframe, on_combine_callback):
        super().__init__(parent)
        self.title(f"{mode} Preview")
        self.geometry("950x550")
        self.mode = mode
        self.df = dataframe.copy()
        self.on_combine_callback = on_combine_callback

        self.source_name = None
        self.target_name = None
        self.saved = False  # track if user saved

        self.undo_stack = []

        # Treeview for preview data
        self.tree = ttk.Treeview(self, show="tree headings")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        # Setup columns: 12 months
        month_cols = [f"{m:02d}" for m in range(1, 13)]
        self.tree["columns"] = month_cols
        self.tree.column("#0", width=150, anchor="w")
        for col in month_cols:
            self.tree.column(col, width=60, anchor="center")
            self.tree.heading(col, text=col)

        # Label showing total rows count - IMPORTANT: create before populate_treeview()
        self.row_count_label = ctk.CTkLabel(self, text="")
        self.row_count_label.pack(side="bottom", anchor="w", padx=10, pady=5)

        # Bind row click
        self.tree.bind("<Button-1>", self.on_row_click)

        # Controls for combine, clearing selection, undo and save
        control_frame = ctk.CTkFrame(self)
        control_frame.pack(pady=(5, 10))

        ctk.CTkLabel(control_frame, text="Source:").grid(row=0, column=0, padx=5)
        self.source_entry = ctk.CTkEntry(control_frame, width=150, state="readonly")
        self.source_entry.grid(row=0, column=1, padx=5)
        ctk.CTkButton(control_frame, text="Clear", command=self.clear_source).grid(row=0, column=2, padx=5)

        ctk.CTkLabel(control_frame, text="Target:").grid(row=0, column=3, padx=5)
        self.target_entry = ctk.CTkEntry(control_frame, width=150, state="readonly")
        self.target_entry.grid(row=0, column=4, padx=5)
        ctk.CTkButton(control_frame, text="Clear", command=self.clear_target).grid(row=0, column=5, padx=5)

        self.combine_btn = ctk.CTkButton(control_frame, text="Combine", command=self.combine_rows)
        self.combine_btn.grid(row=0, column=6, padx=10)

        self.undo_btn = ctk.CTkButton(control_frame, text="Undo", command=self.undo_combine)
        self.undo_btn.grid(row=0, column=7, padx=10)
        self.undo_btn.configure(state="disabled")  # disable at start

        ctk.CTkButton(control_frame, text="Save", command=self.save_and_close).grid(row=0, column=8, padx=10)

        self.protocol("WM_DELETE_WINDOW", self.on_close)  # intercept close window

        # Populate the treeview with data
        self.populate_treeview()

    def populate_treeview(self):
        # Clear current items
        for item in self.tree.get_children():
            self.tree.delete(item)

        month_cols = [f"{m:02d}" for m in range(1, 13)]
        for name, row in self.df.iterrows():
            values = [row.get(m, 0) for m in month_cols]
            self.tree.insert("", "end", iid=name, text=name, values=values)

        # Update row count label
        self.row_count_label.configure(text=f"Total Rows: {len(self.df)}")

        # Update undo button state
        self.undo_btn.configure(state="normal" if self.undo_stack else "disabled")

    def on_row_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        if self.source_name is None:
            self.source_name = item_id
            self.source_entry.configure(state="normal")
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, self.source_name)
            self.source_entry.configure(state="readonly")
            self.highlight_row(self.source_name, "lightgreen")
        elif self.target_name is None and item_id != self.source_name:
            self.target_name = item_id
            self.target_entry.configure(state="normal")
            self.target_entry.delete(0, "end")
            self.target_entry.insert(0, self.target_name)
            self.target_entry.configure(state="readonly")
            self.highlight_row(self.target_name, "lightblue")

    def highlight_row(self, row_id, color):
        # Remove previous tags
        for iid in self.tree.get_children():
            self.tree.item(iid, tags=())
        # Apply new tags
        if color == "lightgreen":
            self.tree.tag_configure("source", background=color)
            self.tree.item(row_id, tags=("source",))
        elif color == "lightblue":
            self.tree.tag_configure("target", background=color)
            self.tree.item(row_id, tags=("target",))

    def clear_source(self):
        if self.source_name:
            self.tree.item(self.source_name, tags=())
        self.source_name = None
        self.source_entry.configure(state="normal")
        self.source_entry.delete(0, "end")
        self.source_entry.configure(state="readonly")

    def clear_target(self):
        if self.target_name:
            self.tree.item(self.target_name, tags=())
        self.target_name = None
        self.target_entry.configure(state="normal")
        self.target_entry.delete(0, "end")
        self.target_entry.configure(state="readonly")

    def combine_rows(self):
        if not self.source_name or not self.target_name:
            messagebox.showerror("Selection Error", "Please select both source and target rows before combining.")
            return
        if self.source_name == self.target_name:
            messagebox.showerror("Selection Error", "Source and target cannot be the same.")
            return

        # Save current df for undo
        self.undo_stack.append(self.df.copy())

        # Combine rows by adding source row values to target row values
        self.df.loc[self.target_name] = self.df.loc[self.target_name].add(self.df.loc[self.source_name], fill_value=0)

        # Drop source row
        self.df = self.df.drop(index=self.source_name)

        # Round up all values to 2 decimal places
        self.df = np.ceil(self.df * 100) / 100

        # Clear selections and refresh UI
        self.clear_source()
        self.clear_target()
        self.populate_treeview()

    def undo_combine(self):
        if not self.undo_stack:
            return
        # Restore previous dataframe state
        self.df = self.undo_stack.pop()
        # Clear selections
        self.clear_source()
        self.clear_target()
        # Refresh UI
        self.populate_treeview()

    def save_and_close(self):
        self.saved = True
        self.on_combine_callback(self.df)
        self.destroy()

    def on_close(self):
        if not self.saved:
            # Do not call callback on close without saving
            self.destroy()