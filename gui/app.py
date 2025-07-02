from logic.summarizer import read_and_summarize
from logic.utils import generate_suggested_output_filename, extract_sheet_name_from_filename, format_sheet
from gui.preview_window import PreviewWindow
from gui.preview_type_dialog import PreviewTypeDialog
import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Monthly Summary")
        self.geometry("600x750")
        self.resizable(False, False)

        self.selected_modes = {"Importer": False, "Exporter": False}
        # Store dict like: {"Importer": {"Weight": df, "Quantity": df}, "Exporter": {...}}
        self.previewed_data = {
            "Importer": {"Weight": None, "Quantity": None},
            "Exporter": {"Weight": None, "Quantity": None},
        }

        ctk.CTkLabel(
            self,
            text="Summarize monthly import/export weights and quantities.",
            wraplength=560,
            justify="left",
            font=ctk.CTkFont(size=14),
        ).pack(pady=(15, 5), padx=10)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=(0, 10))
        self.importer_btn = ctk.CTkButton(button_frame, text="Importer", command=self.toggle_importer, width=120)
        self.importer_btn.grid(row=0, column=0, padx=10)
        self.exporter_btn = ctk.CTkButton(button_frame, text="Exporter", command=self.toggle_exporter, width=120)
        self.exporter_btn.grid(row=0, column=1, padx=10)

        # Preview Buttons with added dropdown
        ctk.CTkButton(button_frame, text="Preview Importer", command=lambda: self.ask_preview_type("Importer")).grid(row=1, column=0, pady=5)
        ctk.CTkButton(button_frame, text="Preview Exporter", command=lambda: self.ask_preview_type("Exporter")).grid(row=1, column=1, pady=5)

        self.mode_label = ctk.CTkLabel(self, text="Mode: None selected", font=ctk.CTkFont(weight="bold"))
        self.mode_label.pack(pady=(0, 10))

        self.input_path_var = ctk.StringVar()
        ctk.CTkLabel(self, text="Input Excel file:").pack()
        input_frame = ctk.CTkFrame(self)
        input_frame.pack(padx=20, fill="x")
        self.input_entry = ctk.CTkEntry(input_frame, textvariable=self.input_path_var, width=400, state="readonly")
        self.input_entry.pack(side="left", padx=(10, 5), pady=10)
        ctk.CTkButton(input_frame, text="Browse", command=self.browse_input_file).pack(side="left", padx=5)

        self.sheet_name_var = ctk.StringVar(value="")
        ctk.CTkLabel(self, text="Sheet name:").pack(pady=(5, 3))
        self.sheet_entry = ctk.CTkEntry(self, textvariable=self.sheet_name_var, width=150)
        self.sheet_entry.pack()

        self.output_path_var = ctk.StringVar()
        ctk.CTkLabel(self, text="Output Excel file:").pack(pady=(15, 3))
        output_frame = ctk.CTkFrame(self)
        output_frame.pack(padx=20, fill="x")
        self.output_entry = ctk.CTkEntry(output_frame, textvariable=self.output_path_var, width=400, state="readonly")
        self.output_entry.pack(side="left", padx=(10, 5), pady=10)
        ctk.CTkButton(output_frame, text="Save As", command=self.browse_output_file).pack(side="left", padx=5)

        self.run_summary_btn = ctk.CTkButton(self, text="Run Summary", command=self.run_summary)
        self.run_summary_btn.pack(pady=20)
        self.run_summary_btn.configure(state="disabled")  # initially disabled

        self.full_input_path = ""
        self.full_output_path = ""

    def toggle_importer(self):
        self.selected_modes["Importer"] = not self.selected_modes["Importer"]
        self.update_mode_label()
        self.update_buttons()

    def toggle_exporter(self):
        self.selected_modes["Exporter"] = not self.selected_modes["Exporter"]
        self.update_mode_label()
        self.update_buttons()

    def auto_suggest_output_path(self):
        if not self.output_path_var.get() and self.input_path_var.get():
            selected = [k for k, v in self.selected_modes.items() if v]
            if selected:
                suggested = generate_suggested_output_filename(self.input_path_var.get(), selected)
                self.output_path_var.set(suggested)

    def update_output_filename(self):
        if not self.input_path_var.get():
            return  # no input file, can't generate
        selected = [k for k, v in self.selected_modes.items() if v]
        if not selected:
            self.output_path_var.set("")  # clear output file name
            self.run_summary_btn.configure(state="disabled")  # disable run button
            return  # nothing selected
        suggested = generate_suggested_output_filename(self.input_path_var.get(), selected)
        self.output_path_var.set(suggested)
        self.run_summary_btn.configure(state="normal")  # enable run button

    def update_mode_label(self):
        active = [k for k, v in self.selected_modes.items() if v]
        self.mode_label.configure(text="Mode: " + (", ".join(active) if active else "None selected"))

        # Update output path suggestion
        if active and self.input_path_var.get():
            folder = os.path.dirname(self.full_input_path)
            suggested_name = generate_suggested_output_filename(self.input_path_var.get(), active)
            self.full_output_path = os.path.join(folder, suggested_name)
            self.output_path_var.set(suggested_name)  # show only the name
            self.run_summary_btn.configure(state="normal")
        else:
            self.output_path_var.set("")
            self.run_summary_btn.configure(state="disabled")

    def update_buttons(self):
        # Change importer/exporter buttons appearance based on selection
        self.importer_btn.configure(fg_color="#4a4a4a" if self.selected_modes["Importer"] else "#0078D7")
        self.exporter_btn.configure(fg_color="#4a4a4a" if self.selected_modes["Exporter"] else "#0078D7")

    def browse_input_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.full_input_path = path
            self.input_path_var.set(os.path.basename(path))  # show only file name
            self.full_input_path = path  # store full path internally
            # Auto suggest sheet name
            sheet_name = extract_sheet_name_from_filename(path)
            if sheet_name:
                self.sheet_name_var.set(sheet_name)

            # Reset previous preview data since input changed
            self.previewed_data = {
                "Importer": {"Weight": None, "Quantity": None},
                "Exporter": {"Weight": None, "Quantity": None},
            }

            # Generate output path in same folder
            selected_modes = [k for k, v in self.selected_modes.items() if v]
            if selected_modes:
                suggested_name = generate_suggested_output_filename(path, selected_modes)
                folder = os.path.dirname(path)
                self.output_path_var.set(os.path.join(folder, suggested_name))
            else:
                self.output_path_var.set("")  # clear if no mode

    def browse_output_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.full_output_path = path
            self.output_path_var.set(os.path.basename(path))  # show only file name
            self.full_output_path = path  # store full path internally

    def ask_preview_type(self, mode):
        # Open small dialog with Weight and Quantity buttons
        PreviewTypeDialog(self, mode, self.preview_data)

    def preview_data(self, mode, type_):
        if not self.input_path_var.get() or not self.sheet_name_var.get():
            messagebox.showerror("Error", "Please select input file and sheet name first.")
            return

        # ✅ If preview already exists, use it directly
        existing_preview = self.previewed_data[mode][type_]
        if existing_preview is not None:
            df_result = existing_preview
        else:
            try:
                df = pd.read_excel(self.full_input_path, sheet_name=self.sheet_name_var.get())
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel: {e}")
                return

            name_col_index = 8 if mode == "Importer" else 4
            value_col_index = 24 if type_.lower() == "weight" else 23

            try:
                df_result = read_and_summarize(df, name_col_index, value_col_index)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process data: {e}")
                return
            self.previewed_data[mode][type_] = df_result

        def on_combine_done(new_df):
            self.previewed_data[mode][type_] = new_df

        PreviewWindow(self, f"{mode} - {type_}", df_result, on_combine_done)

    def run_summary(self):
        if not any(self.selected_modes.values()):
            messagebox.showerror("Error", "Please select at least one mode (Importer or Exporter).")
            return
        if not self.full_input_path:  # use full path variable
            messagebox.showerror("Error", "Please select an input Excel file.")
            return

        # Suggest output file if empty
        if not self.full_output_path:
            suggested_name = generate_suggested_output_filename(
                self.full_input_path,
                [k for k, v in self.selected_modes.items() if v]
            )
            input_dir = os.path.dirname(self.full_input_path)
            self.full_output_path = os.path.join(input_dir, suggested_name)
            self.output_path_var.set(suggested_name)  # show only filename in UI

        try:
            output_path = self.full_output_path
            sheet_names = []  # Store sheet info for formatting later

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for mode, selected in self.selected_modes.items():
                    if not selected:
                        continue
                    for type_ in ["Weight", "Quantity"]:
                        df_preview = self.previewed_data[mode][type_]

                        # If not previewed yet, read and summarize
                        if df_preview is None:
                            try:
                                df = pd.read_excel(self.full_input_path, sheet_name=self.sheet_name_var.get())
                            except Exception as e:
                                messagebox.showerror("Error", f"Failed to read Excel: {e}")
                                return

                            name_col_index = 8 if mode == "Importer" else 4
                            value_col_index = 24 if type_.lower() == "weight" else 23
                            df_preview = read_and_summarize(df, name_col_index, value_col_index)

                        sheet_name = f"{mode.lower()}_{type_.lower()}"
                        df_preview.to_excel(writer, sheet_name=sheet_name)
                        sheet_names.append((sheet_name, mode, type_))  # Save for formatting

            # After saving, apply formatting
            for sheet_name, mode, type_ in sheet_names:
                format_sheet(output_path, sheet_name, mode, type_)

            messagebox.showinfo("Success", f"Summary exported successfully to {output_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export summary: {e}")