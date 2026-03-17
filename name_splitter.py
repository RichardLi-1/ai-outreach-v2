import threading
import pandas as pd
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def open_name_splitter(root, apply_theme_fn):
    win = tk.Toplevel(root)
    win.title("Name Splitter")
    win.minsize(500, 280)
    apply_theme_fn(win)

    state = {"file_path": None, "cols": []}

    # Column mapping
    col_frame = ttk.LabelFrame(win, text="Column Mapping", padding=(10, 8))
    col_frame.pack(padx=20, pady=(16, 8), fill=tk.X)

    ttk.Label(col_frame, text="Full Name column *").grid(row=0, column=0, sticky=tk.W, pady=2)
    name_combo = ttk.Combobox(col_frame, width=32, state="disabled")
    name_combo.grid(row=0, column=1, padx=(8, 0), pady=2)

    ttk.Label(col_frame, text="First Name output column").grid(row=1, column=0, sticky=tk.W, pady=2)
    first_entry = ttk.Entry(col_frame, width=35)
    first_entry.insert(0, "First Name")
    first_entry.grid(row=1, column=1, padx=(8, 0), pady=2)

    ttk.Label(col_frame, text="Last Name output column").grid(row=2, column=0, sticky=tk.W, pady=2)
    last_entry = ttk.Entry(col_frame, width=35)
    last_entry.insert(0, "Last Name")
    last_entry.grid(row=2, column=1, padx=(8, 0), pady=2)

    ttk.Label(col_frame, text="* splits on first space: 'John Smith Doe' → 'John' / 'Smith Doe'",
              style="Gray.TLabel").grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

    # Status
    status_label = ttk.Label(win, text="Select a file to begin.", style="Gray.TLabel")
    status_label.pack(pady=(4, 0))

    file_label = ttk.Label(win, text="", style="Gray.TLabel")
    file_label.pack()

    # Buttons
    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=(8, 12))

    run_btn = ttk.Button(btn_frame, text="Run", width=10, state="disabled")
    run_btn.pack(side=tk.LEFT, padx=6)

    def _select_file():
        path = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel/CSV Files", "*.xlsx *.csv"), ("CSV Files", "*.csv"),
                       ("XLSX Files", "*.xlsx"), ("All Files", "*.*")],
            parent=win,
        )
        if not path:
            return
        try:
            ext = Path(path).suffix.lower()
            df = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}", parent=win)
            return

        state["file_path"] = path
        cols = [c for c in df.columns if c is not None and str(c).strip()]
        state["cols"] = cols
        file_label.config(text=f"File: {Path(path).name}")

        name_combo.config(state="normal", values=cols)
        # Auto-detect name column
        detected = ""
        for col in df.columns:
            low = str(col).lower().strip()
            if low in ("name", "full name", "fullname", "contact name", "contactname"):
                detected = col
                break
        name_combo.set(detected)
        name_combo.config(state="readonly")
        run_btn.config(state="normal")
        status_label.config(text="Ready. Choose the name column then click Run.")

    def _run():
        src_col = name_combo.get()
        if not src_col:
            messagebox.showwarning("Missing Column", "Please select the Full Name column.", parent=win)
            return
        first_col = first_entry.get().strip() or "First Name"
        last_col = last_entry.get().strip() or "Last Name"
        run_btn.config(state="disabled")
        status_label.config(text="Running…")

        def _worker():
            try:
                path = state["file_path"]
                ext = Path(path).suffix.lower()
                df = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)

                def split(val):
                    parts = str(val).strip().split(" ", 1) if pd.notna(val) and str(val).strip() else ["", ""]
                    return parts[0], parts[1] if len(parts) > 1 else ""

                df[first_col], df[last_col] = zip(*df[src_col].map(split))

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"split_{Path(path).stem}_{timestamp}{ext}"
                out_path = Path(path).parent / out_name
                if ext == ".csv":
                    df.to_csv(out_path, index=False)
                else:
                    df.to_excel(out_path, index=False, engine="openpyxl")

                win.after(0, lambda: status_label.config(text=f"Done — {out_name}"))
            except Exception as e:
                win.after(0, lambda: messagebox.showerror("Error", str(e), parent=win))
            finally:
                win.after(0, lambda: run_btn.config(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    run_btn.config(command=_run)
    ttk.Button(btn_frame, text="Select File", width=14, command=_select_file,
               style="Accent.TButton").pack(side=tk.LEFT, padx=6)
