# billing_ui.py
import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from billing_logic import (
    run_single_mode,
    run_batch_mode,
    generate_master_excel
)

# -------------------------------------------------
# UI CONFIG
# -------------------------------------------------
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.geometry("760x560")
app.title("Billing Automation Tool")
app.resizable(False, False)

# -------------------------------------------------
# STATE
# -------------------------------------------------
mode = ctk.StringVar(value="single")
selected_path = ctk.StringVar(value="No folder selected")
status_text = ctk.StringVar(value="Idle")
output_excel = None

# -------------------------------------------------
# FUNCTIONS
# -------------------------------------------------
def select_folder():
    folder = filedialog.askdirectory(title="Select Folder")
    if folder:
        selected_path.set(folder)


def set_running_state(is_running: bool):
    if is_running:
        run_btn.configure(state="disabled")
        select_btn.configure(state="disabled")
        open_btn.configure(state="disabled")
        progress_bar.start()
        status_text.set("Processing… Please wait")
    else:
        progress_bar.stop()
        run_btn.configure(state="normal")
        select_btn.configure(state="normal")
        open_btn.configure(state="normal")


def run_process():
    global output_excel
    try:
        set_running_state(True)

        if mode.get() == "single":
            summary, details = run_single_mode(selected_path.get())
            output_excel = generate_master_excel(
                summary, details, selected_path.get()
            )
        else:
            summary, details = run_batch_mode(selected_path.get())
            output_excel = generate_master_excel(
                summary, details, selected_path.get()
            )

        status_text.set("✅ Completed successfully")
        messagebox.showinfo("Success", "Master Excel generated successfully")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        status_text.set("❌ Error occurred")

    finally:
        set_running_state(False)


def run_thread():
    threading.Thread(target=run_process, daemon=True).start()


def open_excel():
    if output_excel and os.path.exists(output_excel):
        os.startfile(output_excel)

# -------------------------------------------------
# UI LAYOUT
# -------------------------------------------------

# Title
ctk.CTkLabel(
    app,
    text="Billing Automation Tool",
    font=ctk.CTkFont(size=28, weight="bold")
).pack(pady=(20, 5))

ctk.CTkLabel(
    app,
    text="PDF & DOCX | Single Job / Batch Jobs",
    font=ctk.CTkFont(size=14)
).pack(pady=(0, 20))

# Mode Selection
mode_frame = ctk.CTkFrame(app, corner_radius=15)
mode_frame.pack(padx=40, pady=10, fill="x")

ctk.CTkLabel(
    mode_frame,
    text="Select Processing Mode",
    font=ctk.CTkFont(size=16, weight="bold")
).pack(pady=10)

ctk.CTkRadioButton(
    mode_frame,
    text="Single Folder Mode",
    variable=mode,
    value="single"
).pack(anchor="w", padx=20)

ctk.CTkRadioButton(
    mode_frame,
    text="Multi Job (Batch) Mode",
    variable=mode,
    value="batch"
).pack(anchor="w", padx=20, pady=(0, 10))

# Folder Selection
folder_frame = ctk.CTkFrame(app, corner_radius=15)
folder_frame.pack(padx=40, pady=15, fill="x")

select_btn = ctk.CTkButton(
    folder_frame,
    text="📂 Select Folder",
    height=40,
    command=select_folder
)
select_btn.pack(pady=10, padx=20, fill="x")

ctk.CTkLabel(
    folder_frame,
    textvariable=selected_path,
    wraplength=640,
    justify="center"
).pack(pady=(0, 10))

# Progress Section
progress_frame = ctk.CTkFrame(app, corner_radius=15)
progress_frame.pack(padx=40, pady=15, fill="x")

progress_bar = ctk.CTkProgressBar(
    progress_frame,
    mode="indeterminate"
)
progress_bar.pack(padx=20, pady=15, fill="x")

ctk.CTkLabel(
    progress_frame,
    textvariable=status_text,
    font=ctk.CTkFont(size=14, weight="bold")
).pack(pady=(0, 10))

# Action Buttons
run_btn = ctk.CTkButton(
    app,
    text="🚀 Run Billing",
    height=50,
    font=ctk.CTkFont(size=18, weight="bold"),
    command=run_thread
)
run_btn.pack(padx=80, pady=(10, 10), fill="x")

open_btn = ctk.CTkButton(
    app,
    text="📊 Open Excel",
    height=40,
    command=open_excel
)
open_btn.pack(padx=80, pady=(0, 20), fill="x")

# -------------------------------------------------
# START
# -------------------------------------------------
app.mainloop()
