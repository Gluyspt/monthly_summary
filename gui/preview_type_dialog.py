import customtkinter as ctk

# New dialog for selecting Weight or Quantity
class PreviewTypeDialog(ctk.CTkToplevel):
    def __init__(self, parent, mode, callback):
        super().__init__(parent)
        self.title(f"Select Preview Type for {mode}")
        self.geometry("300x120")
        self.resizable(False, False)
        self.callback = callback
        self.mode = mode

        ctk.CTkLabel(self, text=f"Choose preview type for {mode}:", font=ctk.CTkFont(size=14)).pack(pady=(15, 10))

        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10)

        weight_btn = ctk.CTkButton(btn_frame, text="Weight", width=100, command=lambda: self.select_type("Weight"))
        weight_btn.grid(row=0, column=0, padx=10)

        quantity_btn = ctk.CTkButton(btn_frame, text="Quantity", width=100, command=lambda: self.select_type("Quantity"))
        quantity_btn.grid(row=0, column=1, padx=10)

        self.grab_set()  # Make modal
        self.transient(parent)

    def select_type(self, choice):
        self.callback(self.mode, choice)
        self.destroy()