import threading
import logging
import re
import pandas as pd
import openai_hunter_client
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


def open_hunter_finder(root, formatter, apply_theme_fn, TextHandler):
    win = tk.Toplevel(root)
    win.title("Hunter Email Finder")
    win.minsize(620, 560)
    apply_theme_fn(win)

    state = {
        "file_path": None,
        "output_path": None,
        "run_active": False,
        "run_id": 0,
        "cols": [],
        "column_for": {},
    }

    # Progress
    progress = ttk.Progressbar(win, orient="horizontal", length=460, mode="determinate")
    progress.pack(pady=(14, 2), padx=20, fill=tk.X)
    progress_label = ttk.Label(win, text="Select a file to begin.", style="Gray.TLabel")
    progress_label.pack()

    file_label = ttk.Label(win, text="", style="Gray.TLabel")
    file_label.pack(pady=(2, 0))

    # Column mapping
    col_frame = ttk.LabelFrame(win, text="Column Mapping", padding=(10, 8))
    col_frame.pack(padx=20, pady=(8, 4), fill=tk.X)
    col_fields = ["Name", "Organization Website", "LinkedIn", "Contact Tag"]
    col_entries = {}
    for i, field in enumerate(col_fields):
        req = " *" if field in ("Name", "Organization Website") else ""
        ttk.Label(col_frame, text=field + req).grid(row=i, column=0, sticky=tk.W, pady=2)
        entry = ttk.Combobox(col_frame, width=32, state="disabled")
        entry.grid(row=i, column=1, padx=(8, 0), pady=2)
        col_entries[field] = entry
    ttk.Label(col_frame, text="* required  |  Name is split on first space → First Last",
              style="Gray.TLabel").grid(row=len(col_fields), column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

    # Output folder
    out_frame = ttk.LabelFrame(win, text="Output", padding=(10, 8))
    out_frame.pack(padx=20, pady=(4, 8), fill=tk.X)
    ttk.Label(out_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W)
    out_entry = ttk.Entry(out_frame, width=44, state="disabled")
    out_entry.grid(row=0, column=1, padx=(6, 6))

    def select_output():
        folder = filedialog.askdirectory(title="Select Output Folder", parent=win)
        if folder:
            state["output_path"] = folder
            out_entry.config(state="normal")
            out_entry.delete(0, tk.END)
            out_entry.insert(0, folder)
            out_entry.config(state="readonly")

    ttk.Button(out_frame, text="Browse…", command=select_output).grid(row=0, column=2)

    # Control buttons
    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=(0, 6))
    select_btn = ttk.Button(btn_frame, text="Select File", width=14,
                            command=lambda: _select_file(), style="Accent.TButton")
    select_btn.pack(side=tk.LEFT, padx=6)
    run_btn = ttk.Button(btn_frame, text="Run", width=10,
                         command=lambda: _run_thread(), state="disabled")
    run_btn.pack(side=tk.LEFT, padx=6)
    cancel_btn = ttk.Button(btn_frame, text="Cancel", width=10, command=lambda: _cancel())
    cancel_btn.pack(side=tk.LEFT, padx=6)
    cancel_btn.pack_forget()

    # Log
    log_frame = ttk.Frame(win)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
    log_box = scrolledtext.ScrolledText(log_frame, width=70, height=10, wrap=tk.WORD, state="disabled")
    log_box.pack(fill=tk.BOTH, expand=True)

    win_logger = logging.getLogger(f"hunter_finder_{id(win)}")
    win_logger.setLevel(logging.INFO)
    win_handler = TextHandler(log_box)
    win_handler.setFormatter(formatter)
    win_logger.addHandler(win_handler)
    win.protocol("WM_DELETE_WINDOW",
                 lambda: (_cancel(), win_logger.removeHandler(win_handler), win.destroy()))

    def _detect_cols(df):
        detected = {}
        all_cols = [c for c in df.columns if c is not None and str(c).strip()]
        for col in df.columns:
            s = str(col) if not isinstance(col, str) else col
            low = s.lower()
            norm = "".join(low.split())
            if low in ("name", "full name", "contact name", "fullname", "contactname"):
                detected.setdefault("Name", col)
            elif any(k in norm for k in ("website", "domain", "company", "organization", "org", "govwebsite")):
                detected.setdefault("Organization Website", col)
            elif "linkedin" in norm:
                detected.setdefault("LinkedIn", col)
            elif "contact" in norm and "tag" in norm:
                detected.setdefault("Contact Tag", col)
            elif norm == "tag":
                detected.setdefault("Contact Tag", col)
        return detected, all_cols

    def _split_name(full_name):
        parts = str(full_name).strip().split(" ", 1)
        first = parts[0]
        last = parts[1] if len(parts) > 1 else ""
        return first, last

    def _select_file():
        path = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel/CSV Files", "*.xlsx *.csv"), ("CSV Files", "*.csv"),
                       ("XLSX Files", "*.xlsx"), ("All Files", "*.*")],
            parent=win,
        )
        if not path:
            return
        state["file_path"] = path
        file_label.config(text=f"File: {Path(path).name}")
        if not state["output_path"]:
            default = str(Path(path).parent)
            state["output_path"] = default
            out_entry.config(state="normal")
            out_entry.delete(0, tk.END)
            out_entry.insert(0, default)
            out_entry.config(state="readonly")
        try:
            ext = Path(path).suffix.lower()
            df = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{e}", parent=win)
            return
        detected, all_cols = _detect_cols(df)
        state["cols"] = all_cols
        for field, entry in col_entries.items():
            entry.config(state="normal", values=all_cols)
            entry.set(detected.get(field, ""))
            entry.config(state="readonly")
        run_btn.config(state="normal")
        progress_label.config(text="Ready. Adjust columns then click Run.")

    def _cancel():
        if state["run_active"]:
            state["run_id"] += 1
            state["run_active"] = False
            cancel_btn.pack_forget()
            run_btn.config(state="normal")
            progress_label.config(text="Cancelled.")

    def _run_thread():
        col = {f: e.get() or None for f, e in col_entries.items()}
        for req in ("Name", "Organization Website"):
            if not col.get(req):
                messagebox.showwarning("Missing Column", f'"{req}" column is required.', parent=win)
                return
        state["column_for"] = col
        state["run_active"] = True
        state["run_id"] += 1
        run_id = state["run_id"]
        run_btn.config(state="disabled")
        cancel_btn.pack(side=tk.LEFT, padx=6)
        progress.config(value=0)
        progress_label.config(text="Starting…")
        threading.Thread(target=lambda: _process(run_id), daemon=True).start()

    def _extract_domain(val):
        if not val or (isinstance(val, float) and pd.isna(val)):
            return None
        s = str(val).strip()
        s = re.sub(r"^https?://", "", s)
        s = re.sub(r"^www\.", "", s)
        s = s.split("/")[0].split("?")[0].split("#")[0]
        return s if "." in s else None

    def _process(run_id):
        try:
            path = state["file_path"]
            ext = Path(path).suffix.lower()
            try:
                df = pd.read_csv(path) if ext == ".csv" else pd.read_excel(path)
            except Exception as e:
                win_logger.error(f"Could not read file: {e}")
                win.after(0, lambda: progress_label.config(text="Error reading file."))
                return

            col = state["column_for"]
            for c in ("Email", "Email Confidence", "Email Domain", "Company Website", "Hunter Email Source"):
                if c not in df.columns:
                    df[c] = ""

            total = len(df)
            win.after(0, lambda: progress.config(maximum=max(total, 1), value=0))

            output_dir = Path(state["output_path"]) if state["output_path"] else Path(path).parent
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            def _save(cancelled=False):
                out_name = f"hunter_{Path(path).stem}_{timestamp}{ext}"
                out_path = output_dir / out_name
                if ext == ".csv":
                    df.to_csv(out_path, index=False)
                else:
                    df.to_excel(out_path, index=False, engine="openpyxl")
                label = f"{'Cancelled' if cancelled else 'Done'} — written to {out_name}"
                win_logger.info(label)
                return out_name

            for idx, row in df.iterrows():
                if run_id != state["run_id"]:
                    win_logger.info("Run cancelled.")
                    _save(cancelled=True)
                    return

                win.after(0, lambda i=idx: progress.config(value=i + 1))

                raw_name = str(row.get(col["Name"]) or "").strip()
                first, last = _split_name(raw_name)
                org_raw = str(row.get(col["Organization Website"]) or "").strip()
                domain = _extract_domain(org_raw)

                if not first:
                    win_logger.info(f"Row {idx}: skipping — missing name")
                    continue

                if not domain:
                    if not org_raw or org_raw.lower() == "nan":
                        win_logger.info(f"Row {idx}: skipping — no organization provided")
                        continue
                    win_logger.info(f"Row {idx}: no domain — asking OpenAI for '{org_raw}'")
                    domain = openai_hunter_client.find_domain(org_raw)
                    if not domain:
                        win_logger.info(f"Row {idx}: skipping — could not find domain")
                        continue
                    win_logger.info(f"Row {idx}: found domain '{domain}'")

                win.after(0, lambda f=first, l=last, d=domain: progress_label.config(
                    text=f"Looking up {f} {l} @ {d}…", style="Gray.TLabel"))
                win_logger.info(f"Row {idx}: {first} {last} @ {domain}")

                try:
                    res = openai_hunter_client.find_email(first, last, domain)
                    attempt = 1
                    while str(res[0]) == "202" and attempt <= 5:
                        res = openai_hunter_client.find_email(first, last, domain)
                        attempt += 1

                    if str(res[0]) == "200":
                        data = res[1].get("data") or {}
                        email = data.get("email") or ""
                        score = data.get("score", "")
                        sources = data.get("sources") or []
                        source_uri = sources[0]["uri"] if sources else ""
                        email_domain = data.get("domain") or domain
                        linkedin_url = data.get("linkedin_url") or ""

                        df.at[idx, "Email"] = email
                        df.at[idx, "Email Confidence"] = str(score) if score != "" else ""
                        df.at[idx, "Email Domain"] = email_domain
                        df.at[idx, "Company Website"] = org_raw
                        df.at[idx, "Hunter Email Source"] = source_uri

                        if col.get("LinkedIn") and col["LinkedIn"] in df.columns and linkedin_url:
                            df.at[idx, col["LinkedIn"]] = linkedin_url

                        win_logger.info(f"  → {email or '(no email)'} (score: {score})")
                    else:
                        win_logger.info(f"  → Hunter.io returned {res[0]}")
                except Exception as e:
                    win_logger.warning(f"Row {idx}: error — {e}")

            out_name = _save()
            win.after(0, lambda: (
                progress.config(value=total),
                progress_label.config(text="✓ Done!", style="Success.TLabel"),
                cancel_btn.pack_forget(),
                run_btn.config(state="normal"),
            ))
        except Exception as e:
            win_logger.error(f"Unexpected error: {e}")
            win.after(0, lambda: progress_label.config(text="Error — see log."))
        finally:
            state["run_active"] = False
            win.after(0, lambda: (
                cancel_btn.pack_forget(),
                run_btn.config(state="normal" if state["file_path"] else "disabled"),
            ))
