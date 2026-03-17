import threading
import pandas as pd
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


def open_merge_tool(root, apply_theme_fn):
    win = tk.Toplevel(root)
    win.title("Merge Files")
    win.minsize(560, 500)
    apply_theme_fn(win)

    state = {
        "df1": None, "df2": None,
        "path1": None, "path2": None,
        "output_path": None,
        "run_active": False, "run_id": 0,
    }

    # ── File selectors ────────────────────────────────────────────────────────
    files_frame = ttk.LabelFrame(win, text="Input Files", padding=(10, 8))
    files_frame.pack(padx=20, pady=(14, 6), fill=tk.X)
    files_frame.columnconfigure(1, weight=1)

    lbl1 = ttk.Label(files_frame, text="No file selected.", style="Gray.TLabel")
    lbl2 = ttk.Label(files_frame, text="No file selected.", style="Gray.TLabel")

    def _load_file(n):
        path = filedialog.askopenfilename(
            title=f"Select File {n}",
            filetypes=[("Excel/CSV Files", "*.xlsx *.csv"), ("All Files", "*.*")],
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
        if n == 1:
            state["df1"], state["path1"] = df, path
            lbl1.config(text=Path(path).name)
            if not state["output_path"]:
                _set_output(str(Path(path).parent))
        else:
            state["df2"], state["path2"] = df, path
            lbl2.config(text=Path(path).name)
        _refresh_columns()

    ttk.Button(files_frame, text="File 1…", width=10,
               command=lambda: _load_file(1), style="Accent.TButton").grid(row=0, column=0, pady=3, padx=(0, 8))
    lbl1.grid(row=0, column=1, sticky=tk.W)
    ttk.Button(files_frame, text="File 2…", width=10,
               command=lambda: _load_file(2), style="Accent.TButton").grid(row=1, column=0, pady=3, padx=(0, 8))
    lbl2.grid(row=1, column=1, sticky=tk.W)

    # ── Merge key columns ─────────────────────────────────────────────────────
    key_frame = ttk.LabelFrame(win, text="Merge On (common columns — select one or more)", padding=(10, 8))
    key_frame.pack(padx=20, pady=(0, 6), fill=tk.X)

    key_list = tk.Listbox(key_frame, selectmode=tk.MULTIPLE, height=5, exportselection=False)
    key_scroll = ttk.Scrollbar(key_frame, orient=tk.VERTICAL, command=key_list.yview)
    key_list.config(yscrollcommand=key_scroll.set)
    key_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    key_scroll.pack(side=tk.LEFT, fill=tk.Y)

    no_common_lbl = ttk.Label(key_frame, text="Load both files to see common columns.", style="Gray.TLabel")
    no_common_lbl.pack(side=tk.LEFT, padx=8)

    def _refresh_columns():
        df1, df2 = state["df1"], state["df2"]
        if df1 is None or df2 is None:
            return

        # Case-insensitive, stripped match: map normalised name → original name in df2
        df2_norm = {str(c).strip().lower(): c for c in df2.columns}
        # Build list of (col_in_df1, col_in_df2) pairs
        common_pairs = []
        for c in df1.columns:
            match = df2_norm.get(str(c).strip().lower())
            if match is not None:
                common_pairs.append((c, match))
        state["_col_pairs"] = common_pairs  # store for use in _worker

        key_list.delete(0, tk.END)
        if not common_pairs:
            no_common_lbl.config(text="No common columns found.")
            run_btn.config(state="disabled")
            return
        no_common_lbl.config(text="")
        for col1, col2 in common_pairs:
            label = col1 if str(col1).strip().lower() == str(col2).strip().lower() else f"{col1}  ↔  {col2}"
            key_list.insert(tk.END, label)
        # auto-select columns that look like identifiers
        for i, (col1, _) in enumerate(common_pairs):
            if any(k in str(col1).lower() for k in ("id", "name", "email", "key")):
                key_list.selection_set(i)
        if not key_list.curselection():
            key_list.selection_set(0)
        run_btn.config(state="normal")
        status_label.config(text="Ready. Select merge key(s) then click Run.")

    # ── Output folder ─────────────────────────────────────────────────────────
    out_frame = ttk.LabelFrame(win, text="Output", padding=(10, 8))
    out_frame.pack(padx=20, pady=(0, 8), fill=tk.X)
    ttk.Label(out_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W)
    out_entry = ttk.Entry(out_frame, width=44, state="disabled")
    out_entry.grid(row=0, column=1, padx=(6, 6))

    def _set_output(folder):
        state["output_path"] = folder
        out_entry.config(state="normal")
        out_entry.delete(0, tk.END)
        out_entry.insert(0, folder)
        out_entry.config(state="readonly")

    def _browse_output():
        folder = filedialog.askdirectory(title="Select Output Folder", parent=win)
        if folder:
            _set_output(folder)

    ttk.Button(out_frame, text="Browse…", command=_browse_output).grid(row=0, column=2)

    # ── Status + log ──────────────────────────────────────────────────────────
    status_label = ttk.Label(win, text="Select two files to begin.", style="Gray.TLabel")
    status_label.pack(pady=(2, 0))

    log_box = scrolledtext.ScrolledText(win, width=65, height=7, wrap=tk.WORD, state="disabled")
    log_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 4))

    def _log(msg):
        log_box.config(state="normal")
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)
        log_box.config(state="disabled")

    # ── Buttons ───────────────────────────────────────────────────────────────
    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=(0, 10))

    run_btn = ttk.Button(btn_frame, text="Run", width=10, state="disabled")
    run_btn.pack(side=tk.LEFT, padx=6)
    cancel_btn = ttk.Button(btn_frame, text="Cancel", width=10)
    cancel_btn.pack(side=tk.LEFT, padx=6)
    cancel_btn.pack_forget()

    def _cancel():
        if state["run_active"]:
            state["run_id"] += 1
            state["run_active"] = False
            cancel_btn.pack_forget()
            run_btn.config(state="normal")
            status_label.config(text="Cancelled.")
            _log("Run cancelled.")

    cancel_btn.config(command=_cancel)

    def _run():
        sel = key_list.curselection()
        if not sel:
            messagebox.showwarning("No Key", "Select at least one column to merge on.", parent=win)
            return
        pairs = state.get("_col_pairs", [])
        selected_pairs = [pairs[i] for i in sel]
        how = how_var.get()

        state["run_active"] = True
        state["run_id"] += 1
        run_id = state["run_id"]
        run_btn.config(state="disabled")
        cancel_btn.pack(side=tk.LEFT, padx=6)
        status_label.config(text="Merging…")

        def _worker():
            try:
                if run_id != state["run_id"]:
                    return
                df1 = state["df1"].copy()
                df2 = state["df2"].copy()
                path1 = state["path1"]

                keys1 = [p[0] for p in selected_pairs]
                keys2 = [p[1] for p in selected_pairs]
                # Rename df2 key columns to match df1 so merge works
                rename_map = {k2: k1 for k1, k2 in zip(keys1, keys2) if k1 != k2}
                if rename_map:
                    df2 = df2.rename(columns=rename_map)
                    _log(f"  Renamed in File 2: {rename_map}")
                keys = keys1

                _log(f"Merging on: {', '.join(keys)}  ({how} join)")
                _log(f"File 1: {len(df1)} rows   File 2: {len(df2)} rows")

                dup1 = len(df1) - df1[keys].drop_duplicates().shape[0]
                dup2 = len(df2) - df2[keys].drop_duplicates().shape[0]
                if dup1:
                    _log(f"  Warning: File 1 has {dup1} duplicate key row(s) — keeping first occurrence.")
                    df1 = df1.drop_duplicates(subset=keys, keep="first")
                if dup2:
                    _log(f"  Warning: File 2 has {dup2} duplicate key row(s) — keeping first occurrence.")
                    df2 = df2.drop_duplicates(subset=keys, keep="first")

                merged = pd.merge(df1, df2, on=keys, how=how)

                if run_id != state["run_id"]:
                    _log("Cancelled.")
                    return

                _log(f"Result: {len(merged)} rows")

                ext = Path(path1).suffix.lower()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"merged_{Path(path1).stem}_{timestamp}{ext}"
                out_dir = Path(state["output_path"]) if state["output_path"] else Path(path1).parent
                out_path = out_dir / out_name

                if ext == ".csv":
                    merged.to_csv(out_path, index=False)
                else:
                    merged.to_excel(out_path, index=False, engine="openpyxl")

                _log(f"Saved: {out_name}")
                win.after(0, lambda: status_label.config(text=f"Done — {out_name}", style="Success.TLabel"))
            except Exception as e:
                win.after(0, lambda: (
                    status_label.config(text="Error — see log."),
                    _log(f"Error: {e}"),
                ))
            finally:
                state["run_active"] = False
                win.after(0, lambda: (
                    cancel_btn.pack_forget(),
                    run_btn.config(state="normal" if (state["df1"] is not None and state["df2"] is not None) else "disabled"),
                ))

        threading.Thread(target=_worker, daemon=True).start()

    run_btn.config(command=_run)

    # ── Join type ─────────────────────────────────────────────────────────────
    how_frame = ttk.Frame(win)
    how_frame.pack(before=btn_frame, pady=(0, 4))
    ttk.Label(how_frame, text="Join type:").pack(side=tk.LEFT, padx=(0, 6))
    how_var = tk.StringVar(value="outer")
    for label, val in (("Outer (keep all rows)", "outer"), ("Inner (matching only)", "inner"),
                       ("Left", "left"), ("Right", "right")):
        ttk.Radiobutton(how_frame, text=label, variable=how_var, value=val).pack(side=tk.LEFT, padx=4)
