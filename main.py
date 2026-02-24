import pandas as pd
import openai_hunter_client
import json
from presets import Role, stateCorrectionMap
from datetime import datetime
import winsound
import threading
from settings import settings
import os
import re
import darkdetect
import sys
import openai
import tkinter as tk
import sv_ttk
from tkinter import messagebox, filedialog, scrolledtext, ttk
import logging
from pathlib import Path

try:
    import pywinstyles
    HAS_PYWINSTYLES = True
except ImportError:
    HAS_PYWINSTYLES = False


class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        print(msg)
        self.text_widget.after(0, self.append, msg)

    def append(self, msg):
        self.text_widget.config(state="normal")
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.see(tk.END)
        self.text_widget.config(state="disabled")
        self.text_widget.update_idletasks()
        

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AI Outreach")

        # Apply Sun Valley theme
        sv_ttk.set_theme(darkdetect.theme())

        # Apply theme to Windows title bar
        if HAS_PYWINSTYLES:
            self._apply_theme_to_titlebar()

        # Configure custom styles
        style = ttk.Style()

        # Green button for "Open folder" with higher contrast
        #style.configure("Green.TButton", font=("TkDefaultFont", 10, "bold"), borderwidth=2)
        #style.map("Green.TButton",
        #         foreground=[("!disabled", "#2e7d32")],
        #         background=[("!disabled", "#4CAF50")])

        style.configure("Gray.TLabel", foreground="gray")
        style.configure("Welcome.TLabel", foreground="#555")
        style.configure("Success.TLabel", foreground="#4CAF50", font=("TkDefaultFont", 11, "bold"))
        style.configure("Processing.TLabel", foreground="gray", font=("TkDefaultFont", 9, "bold"))

        self.file_path = None
        self.prompt_gis = settings.initial_prompt
        self.prompt_mayor = settings.initial_prompt_mayor
        self.prompt_assessor = settings.initial_prompt_assessor
        self.output_path = None
        self.cols = []

        self.column_for = {}
        self.dynamic_widgets = []
        self.current_result_event = None
        self.current_run_id = 0
        self.sheet_role_choices = {}
        self._run_active = False

        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=460, mode="determinate")
        self.progress.pack(pady=(14, 2), padx=20, fill=tk.X)
        self.progress_label = ttk.Label(self.root, text="", style="Gray.TLabel")
        self.progress_label.pack()

        # Welcome message — hidden once a file is selected
        self.welcome_label = ttk.Label(
            self.root,
            text="Welcome to AI Outreach.\nGet started by selecting a file.",
            style="Welcome.TLabel", justify=tk.CENTER
        )
        self.welcome_label.pack(pady=(8, 0))

        # Open folder button — shown when processing completes
        self.open_folder_btn = ttk.Button(
            self.root, text="📁 Open Output Folder",
            command=self._open_output_folder,
            style="Accent.TButton", padding=(20, 8)
        )
        # Initially hidden (not packed)

        self.cancel_btn = ttk.Button(
            self.root, text="Cancel",
            command=self._cancel_run,
            padding=(20, 4)
        )
        # Initially hidden (not packed)

        # Current file label
        self.current_file_label = ttk.Label(self.root, text="", style="Gray.TLabel")
        self.current_file_label.pack(pady=(4, 0))

        # Collapsible log section
        log_section = ttk.Frame(self.root)
        log_section.pack(fill=tk.X, padx=10, pady=(6, 0))
        self._log_visible = False
        self._log_toggle_btn = ttk.Button(log_section, text="Logs ▶",
                                         command=self._toggle_logs, width=12)
        self._log_toggle_btn.pack(anchor=tk.W)
        output_box = scrolledtext.ScrolledText(log_section, width=80, height=14, wrap=tk.WORD, state="disabled")
        self._output_box = output_box  # starts hidden (not packed)

        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)

        self.text_handler = TextHandler(output_box)
        self.formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        self.text_handler.setFormatter(self.formatter)

        self.logger.addHandler(self.text_handler)

        # Controls row: Settings + Select File
        controls_frame = ttk.Frame(self.root)
        controls_frame.pack(pady=(14, 6))
        ttk.Button(controls_frame, text="Settings", width=14, command=self.open_settings).pack(side=tk.LEFT, padx=8)
        ttk.Button(controls_frame, text="Select File", width=14, command=self.select_file, style="Accent.TButton").pack(side=tk.LEFT, padx=8)

        # Output path section
        output_frame = ttk.LabelFrame(self.root, text="Output", padding=(10, 8))
        output_frame.pack(padx=20, pady=(4, 14), fill=tk.X)
        ttk.Label(output_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.output_entry = ttk.Entry(output_frame, width=52, state="disabled")
        self.output_entry.grid(row=0, column=1, padx=(6, 6), pady=2)
        self.select_output_btn = ttk.Button(output_frame, text="Browse…",
                                           command=lambda: self.select_output_folder(self.output_entry))
        self.select_output_btn.grid(row=0, column=2, pady=2)




    def _toggle_logs(self):
        if self._log_visible:
            self._output_box.pack_forget()
            self._log_toggle_btn.config(text="Logs ▶")
            self._log_visible = False
        else:
            self._output_box.pack(fill=tk.BOTH, expand=True)
            self._log_toggle_btn.config(text="Logs ▼")
            self._log_visible = True

    def _apply_theme_to_titlebar(self, window=None):
        """Apply dark/light theme to Windows title bar"""
        if window is None:
            window = self.root
        try:
            version = sys.getwindowsversion()
            current_theme = sv_ttk.get_theme()

            if version.major == 10 and version.build >= 22000:
                # Windows 11: Set title bar color to match theme
                pywinstyles.change_header_color(window, "#1c1c1c" if current_theme == "dark" else "#fafafa")
            elif version.major == 10:
                # Windows 10: Apply style
                pywinstyles.apply_style(window, "dark" if current_theme == "dark" else "normal")

                # Hacky way to update title bar color on Windows 10
                window.wm_attributes("-alpha", 0.99)
                window.wm_attributes("-alpha", 1)
        except Exception:
            # Silently fail if pywinstyles doesn't work
            pass

    def _open_output_folder(self):
        folder = self.output_path or (str(Path(self.file_path).parent) if self.file_path else None)
        if folder:
            try:
                os.startfile(folder)
            except Exception:
                self.logger.error(f"Could not open folder: {folder}")

    def _cancel_run(self):
        if not messagebox.askyesno("Cancel Run", "Are you sure you want to cancel the current run?"):
            return
        if not self._run_active:
            return  # run completed while dialog was open
        self._run_active = False
        self.current_run_id += 1
        if self.current_result_event:
            self.current_result_event.set()
            self.current_result_event = None
        self.cancel_btn.pack_forget()
        self.progress.config(value=0)
        self.progress_label.config(text="Cancelling...", style="Gray.TLabel")
        self.select_output_btn.config(state="normal")
        self.logger.info("Run cancelled by user. Exiting...")

    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.grab_set()
        settings_window.minsize(660, 400)

        # Apply theme to settings window title bar
        if HAS_PYWINSTYLES:
            self._apply_theme_to_titlebar(settings_window)

        # Scrollable content area fills window; button bar is fixed at bottom
        settings_window.rowconfigure(0, weight=1)
        settings_window.columnconfigure(0, weight=1)

        canvas = tk.Canvas(settings_window, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(settings_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        scroll_frame = ttk.Frame(canvas)
        scroll_frame.columnconfigure(0, weight=1)
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        settings_window.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        ttk.Label(scroll_frame, text="GIS prompt").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))
        gis_prompt_box = tk.Text(scroll_frame, height=10, wrap=tk.WORD)
        gis_prompt_box.grid(row=1, column=0, sticky="ew", padx=10)
        gis_prompt_box.insert("1.0", self.prompt_gis)

        ttk.Label(scroll_frame, text="Assessor prompt").grid(row=2, column=0, sticky="w", padx=10, pady=(10, 2))
        assessor_prompt_box = tk.Text(scroll_frame, height=10, wrap=tk.WORD)
        assessor_prompt_box.grid(row=3, column=0, sticky="ew", padx=10)
        assessor_prompt_box.insert("1.0", self.prompt_assessor)

        ttk.Label(scroll_frame, text="Mayor prompt").grid(row=4, column=0, sticky="w", padx=10, pady=(10, 2))
        mayor_prompt_box = tk.Text(scroll_frame, height=10, wrap=tk.WORD)
        mayor_prompt_box.grid(row=5, column=0, sticky="ew", padx=10)
        mayor_prompt_box.insert("1.0", self.prompt_mayor)

        # Button bar fixed at bottom, outside scroll area
        btn_frame = ttk.Frame(settings_window)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(4, 10))

        def load():
            config_path = filedialog.askopenfilename(
                title="Select Config File",
                filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
            )

            if not config_path:
                return

            with open(config_path, "r") as file:
                try:
                    config = json.load(file)
                    self.prompt_gis = config.get("prompt_gis", self.prompt_gis)
                    self.prompt_mayor = config.get("prompt_mayor", self.prompt_mayor)
                    self.prompt_assessor = config.get("prompt_assessor", self.prompt_assessor)

                    gis_prompt_box.delete("1.0", tk.END)
                    gis_prompt_box.insert("1.0", self.prompt_gis)
                    gis_prompt_box.see("1.0")
                    assessor_prompt_box.delete("1.0", tk.END)
                    assessor_prompt_box.insert("1.0", self.prompt_assessor)
                    assessor_prompt_box.see("1.0")
                    mayor_prompt_box.delete("1.0", tk.END)
                    mayor_prompt_box.insert("1.0", self.prompt_mayor)
                    mayor_prompt_box.see("1.0")

                    self.logger.info(f"Config loaded from {config_path}")
                except json.JSONDecodeError:
                    self.logger.error("JSON file is corrupted.")
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    messagebox.showinfo("Unable to load file", "JSON file is corrupted.", parent=settings_window)
                except Exception as e:
                    self.logger.error(f"Unexpected error: {e}")
        
        def save():
            if self.prompt_mayor != mayor_prompt_box.get("1.0", tk.END).strip():
                self.prompt_mayor = mayor_prompt_box.get("1.0", tk.END).strip()
                self.logger.info("Mayor prompt temporarily updated to: " + self.prompt_mayor)
            if self.prompt_assessor != assessor_prompt_box.get("1.0", tk.END).strip():
                self.prompt_assessor = assessor_prompt_box.get("1.0", tk.END).strip()
                self.logger.info("Assessor prompt temporarily updated to: " + self.prompt_assessor)
            if self.prompt_gis != gis_prompt_box.get("1.0", tk.END).strip():
                self.prompt_gis = gis_prompt_box.get("1.0", tk.END).strip()
                self.logger.info("GIS prompt temporarily updated to: " + self.prompt_gis)
            settings_window.destroy()
        
        def save_config():
            config_path = filedialog.asksaveasfilename(
                title="Save Config File",
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
                initialfile = "config_" + datetime.now().strftime("%Y%m%d_%H%M%S")
            )

            if not config_path:
                return
            config = {
                "prompt_gis": gis_prompt_box.get("1.0", tk.END).strip(),
                "prompt_mayor": mayor_prompt_box.get("1.0", tk.END).strip(),
                "prompt_assessor": assessor_prompt_box.get("1.0", tk.END).strip(),
            }

            try:
                with open(config_path, "w") as f:
                    json.dump(config, f, indent=2)
                self.logger.info(f"Config saved to {config_path}")
            except PermissionError:
                self.logger.error(f"Failed to save config to {config_path}, permission denied.")
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
                messagebox.showinfo("Unable to save file", f"Failed to save config to {config_path}", parent=settings_window)
        
        def has_unsaved_changes():
            return (
                gis_prompt_box.get("1.0", tk.END).strip() != self.prompt_gis or
                assessor_prompt_box.get("1.0", tk.END).strip() != self.prompt_assessor or
                mayor_prompt_box.get("1.0", tk.END).strip() != self.prompt_mayor
            )

        def on_closing():
            if has_unsaved_changes():
                result = messagebox.askyesnocancel(
                    "Unsaved Changes",
                    "You have unsaved changes. Save before closing?",
                    parent=settings_window
                )
                if result is True:
                    save()
                elif result is False:
                    settings_window.destroy()
                # None = Cancel, do nothing
            else:
                settings_window.destroy()

        settings_window.protocol("WM_DELETE_WINDOW", on_closing)

        ttk.Button(btn_frame, text="Load Config from File", width=20, command=load).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save Config to File", width=20, command=save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save and Close", width=20, command=save).pack(side=tk.LEFT, padx=5)

    def select_role_to_search(self, section_name, sheet_name, run_id, state_val="unknown"):
        result = threading.Event()
        self.current_result_event = result
        role_choice_value = [0]

        def show_dialog():
            winsound.MessageBeep(winsound.MB_ICONASTERISK)

            # Main container with two columns
            container = ttk.Frame(self.root)
            container.pack(pady=5, fill=tk.X)
            self.dynamic_widgets.append(container)

            # Left side: checkboxes
            left_frame = ttk.Frame(container)
            left_frame.pack(side=tk.LEFT, padx=10, anchor=tk.N)

            display_name = re.sub(r"_part(\d+)$", lambda m: f" Part {m.group(1)}", section_name)
            state_display = f" ({state_val})" if state_val and state_val != "unknown" else ""
            label = ttk.Label(left_frame, text=f"Select role(s) to search for in sheet {display_name} section {state_display}")
            label.pack(pady=5)
            check_vars = {}
            role_labels = {1: "GIS Manager", 2: "Mayor/County Manager", 3: "Property Assessor"}
            for text, val in [("GIS Manager", 1), ("Mayor/County Manager", 2), ("Property Assessor", 3), ("Skip", 0)]:
                var = tk.IntVar(value=0)
                check_vars[val] = var
                cb = ttk.Checkbutton(left_frame, text=text, variable=var,
                                    command=lambda v=val: on_check_toggle(v))
                cb.pack(pady=3, anchor=tk.W)

            sheet_action_var = tk.IntVar(value=0)
            sheet_action_label = tk.StringVar()
            last_mode = ["apply"]

            def update_sheet_action_label():
                display_sheet = re.sub(r"_part(\d+)$", lambda m: f" Part {m.group(1)}", sheet_name)
                if check_vars[0].get() == 1:
                    sheet_action_label.set(f"Skip all remaining sections in sheet \"{display_sheet}\"")
                else:
                    selected_roles = [role_labels[val] for val, var in check_vars.items() if var.get() == 1 and val in role_labels]
                    roles_text = ", ".join(selected_roles) if selected_roles else "selected role(s)"
                    sheet_action_label.set(f"Apply {roles_text} to all remaining sections in sheet \"{display_sheet}\"")

            def update_run_button_state():
                has_role = any(entry.get() for entry in check_vars.values())
                btn.config(state="normal" if has_role else "disabled")

            def update_run_button_text():
                btn.config(text="Skip" if check_vars[0].get() == 1 else "Run")

            def update_sheet_action_state():
                has_role = any(entry.get() for entry in check_vars.values())
                if has_role:
                    sheet_action_cb.config(state="normal")
                else:
                    sheet_action_var.set(0)
                    sheet_action_cb.config(state="disabled")

            def on_check_toggle(toggled_val):
                if toggled_val == 0 and check_vars[0].get() == 1:
                    for val, var in check_vars.items():
                        if val != 0:
                            var.set(0)
                elif toggled_val != 0 and check_vars[toggled_val].get() == 1:
                    check_vars[0].set(0)
                mode = "skip" if check_vars[0].get() == 1 else "apply"
                if mode != last_mode[0]:
                    sheet_action_var.set(0)
                    last_mode[0] = mode
                update_sheet_action_label()
                update_run_button_state()
                update_run_button_text()
                update_sheet_action_state()

            update_sheet_action_label()
            sheet_action_cb = ttk.Checkbutton(
                left_frame,
                textvariable=sheet_action_label,
                variable=sheet_action_var
            )
            sheet_action_cb.pack(pady=(14, 3), anchor=tk.W)
            sheet_action_cb.config(state="disabled")

            # Right side: column mapping grid
            right_frame = ttk.Frame(container)
            right_frame.pack(side=tk.LEFT, padx=10, anchor=tk.N)
            rightLabel = ttk.Label(right_frame, text=f"Select columns")
            rightLabel.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)

            mapping_fields = self.column_for

            entries = {}
            for row, (lbl_text, default_val) in enumerate(mapping_fields.items(), start=1):
                ttk.Label(right_frame, text=lbl_text).grid(row=row, column=0, sticky=tk.W, pady=2)
                entry = ttk.Combobox(right_frame, values=self.cols, width=25)
                entry.grid(row=row, column=1, pady=2, padx=(5, 0))
                if default_val:
                    entry.set(default_val)
                entry.config(state="readonly")
                entry.bind("<MouseWheel>", lambda _: "break")
                entries[lbl_text] = entry
            
            def on_run():
                if check_vars[0].get() == 1:
                    selected = [0]
                else:
                    selected = [val for val, var in check_vars.items() if var.get() == 1]
                if not selected:
                    selected = [0]
                role_choice_value.clear()
                role_choice_value.extend(selected)

                if sheet_action_var.get() == 1:
                    self.sheet_role_choices[sheet_name] = selected

                # Update column mappings from combobox selections
                self.column_for = {key: entry.get() or None for key, entry in entries.items()}

                for w in self.dynamic_widgets:
                    try:
                        w.destroy()
                    except Exception:
                        pass
                self.dynamic_widgets.clear()
                self.select_output_btn.config(state="disabled")
                result.set()

            btn = ttk.Button(self.root, text="Run", width=20, command=on_run, state="disabled", style="Accent.TButton")
            btn.pack(pady=5)
            self.dynamic_widgets.append(btn)

            

        if run_id == self.current_run_id:
            self.root.after(0, show_dialog)
        result.wait()
        # Return None to signal cancellation if a new file was selected mid-wait
        if run_id != self.current_run_id:
            return None
        return role_choice_value


    def select_file(self):
        # Warn if a run is in progress
        if self._run_active:
            if not messagebox.askyesno(
                "Run in Progress",
                "A run is currently in progress. Cancel it and select a new file?"
            ):
                return

        # Clean up any widgets from a previous run
        for w in self.dynamic_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self.dynamic_widgets.clear()

        # Invalidate old run BEFORE unblocking the waiting thread, so it
        # sees the new run_id immediately and exits instead of continuing.
        self.current_run_id += 1
        if self.current_result_event:
            self.current_result_event.set()
            self.current_result_event = None
        self.sheet_role_choices = {}

        self.open_folder_btn.pack_forget()
        self.file_path = filedialog.askopenfilename(
            title = "Select a file",
            filetypes = [("Excel/CSV Files", "*.xlsx *.csv"), ("XLSX Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if self.file_path:
            self.welcome_label.pack_forget()
            self.select_output_btn.config(state="normal")
            self.progress.config(value=0)
            self.progress_label.config(text="", style="Gray.TLabel")
            self.current_file_label.config(text=f"Current file: {Path(self.file_path).name}")
            self.logger.info(f"Selected: {self.file_path}")
            # Default output path to the file's parent directory
            if not self.output_path:
                default_dir = str(Path(self.file_path).parent)
                self.output_path = default_dir
                self.output_entry.config(state="normal")
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, default_dir)
                self.output_entry.config(state="readonly")
            run_id = self.current_run_id
            # Run main in a separate thread to keep GUI responsive
            self._run_active = True
            self.cancel_btn.pack(pady=(0, 4))
            thread = threading.Thread(target=lambda: self.main(run_id), daemon=True)
            thread.start()


    def select_output_folder(self, output):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_path = folder
            output.config(state="normal")
            output.delete(0, tk.END)
            output.insert(0, folder)
            output.config(state="readonly")



    def run(self):
        self.root.mainloop()

    def _detect_columns(self, df):
        cols = {}
        self.cols = []
        for column in df.columns:
            # Convert to string in case column name is numeric
            column_str = str(column) if not isinstance(column, str) else column
            lower = column_str.lower()
            normalized = ''.join(lower.split())
            if lower in ["county", "county/city", "city/county"]:
                cols["County/City"] = column
            elif lower in ["email", "contact email"]:
                cols["Email"] = column
            elif lower in ["number", "phone number", "contact phone number", "contact number", "contact phone"]:
                cols["Phone Number"] = column
            elif lower in ["first name", "contact first name", "first"]:
                cols["First Name"] = column
            elif lower in ["last name", "contact last name", "last", "surname"]:
                cols["Last Name"] = column
            elif normalized in ["position", "role", "title", "role/title", "title/role"]:
                cols["Role/Title"] = column
            elif normalized in ["state", "province", "state/province", "province/state", "provinceorstate"]:
                cols["State"] = column
            elif normalized in ["contactlinkedinprofile", "contactlinkedin", "linkedin", "linkedinprofile"]:
                cols["LinkedIn"] = column
            elif normalized in ["tag"]:
                cols["Tag"] = column
            elif normalized in ["contacttag"]:
                cols["Contact Tag"] = column
            if column is not None and str(column).strip() != "":
                self.cols.append(column)
        return cols

    def _find_duplicate_headers(self, df):
        """
        Find rows that look like header rows (indicating stacked datasets).
        Returns a list of row indices where headers are found.
        """
        header_rows = []

        # Keywords commonly found in headers (single words only for whole-word matching)
        header_keywords = ['country', 'county', 'city', 'email', 'state', 'province',
                          'contact', 'population', 'phone', 'role', 'title', 'name',
                          'first', 'last', 'position', 'address']

        for idx in range(len(df)):
            row = df.iloc[idx]
            # Convert all values in the row to lowercase strings
            row_values = [str(val).lower().strip() for val in row if pd.notna(val) and str(val).strip()]

            # Count how many header keywords appear in this row
            # A cell is a header cell if MOST of its words are keywords
            keyword_matches = 0
            for val in row_values:
                # Split into words
                val_words = val.replace('/', ' ').replace('-', ' ').split()
                if len(val_words) == 0:
                    continue

                # Count how many words in this cell are keywords
                keyword_count_in_cell = sum(1 for word in val_words if word in header_keywords)

                # Cell is a header cell if MAJORITY of its words are keywords (>50%)
                # (e.g., "county" = 1/1 ✓, "county/city" = 2/2 ✓, "brevard county" = 1/2 = 50% ✗)
                if keyword_count_in_cell > len(val_words) * 0.5:
                    keyword_matches += 1

            # If 3+ cells contain header keywords, it's likely a header row
            if keyword_matches >= 3:
                header_rows.append(idx)

        return header_rows

    def _split_by_duplicate_headers(self, df, sheet_name):
        """
        Split a dataframe into multiple sections if duplicate headers are found.
        Returns a list of tuples: (section_name, section_dataframe)
        """
        header_indices = self._find_duplicate_headers(df)

        # If no headers detected, fall back to treating row 0 as the header
        if len(header_indices) == 0:
            section_df = df.copy()
            section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
            section_df = section_df.iloc[1:].reset_index(drop=True)
            return [(sheet_name, section_df)]
        # If exactly one header found, normalize into a single section
        if len(header_indices) == 1:
            start_idx = header_indices[0]
            section_df = df.iloc[start_idx:].copy()
            section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
            section_df = section_df.iloc[1:].reset_index(drop=True)
            if len(section_df) < 1:
                return []
            return [(sheet_name, section_df)]

        # Log that we found multiple headers
        self.logger.info(f"**Found {len(header_indices)} sections in sheet '{sheet_name}'**")

        # Split into sections
        sections = []
        for i in range(len(header_indices)):
            start_idx = header_indices[i]
            # End is either the next header or the end of the dataframe
            end_idx = header_indices[i + 1] if i + 1 < len(header_indices) else len(df)

            # Extract this section
            section_df = df.iloc[start_idx:end_idx].copy()

            # Skip if section is too small (less than 2 rows including header)
            if len(section_df) < 2:
                continue

            # Set the first row as column headers and remove it from data
            # Convert all header values to strings to avoid numeric column names
            section_df.columns = [str(col) if pd.notna(col) else "" for col in section_df.iloc[0]]
            section_df = section_df.iloc[1:].reset_index(drop=True)

            # After removing header, check if we still have data
            if len(section_df) < 1:
                continue

            # Create a name for this section
            section_name = f"{sheet_name}_part{i + 1}"
            sections.append((section_name, section_df))

            self.logger.info(f"  - Section {i + 1}: {len(section_df)} rows (data only)")

        return sections
    
    
    

    def main(self, run_id):
        try:
            ext = Path(self.file_path).suffix.lower()
            try:
                if ext == ".csv":
                    sheets = {"Sheet1": pd.read_csv(self.file_path, header=None)}
                else:
                    sheets = pd.read_excel(self.file_path, sheet_name=None, header=None)
            except ValueError as e:
                err = str(e)
                self.logger.error(f"Could not read file: {err}")
                self.root.after(0, lambda err=err: (
                    messagebox.showerror("File Error", f"The file could not be read.\n\nMake sure it is a valid .xlsx, .xls, or .csv file.\n\n{err}"),
                    self.select_output_btn.config(state="normal"),
                    self.cancel_btn.pack_forget(),
                    self.progress.config(value=0),
                    self.progress_label.config(text="", style="Gray.TLabel")
                ))
                return
            

            # Create ExcelWriter to write multiple sheets
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            input_path = Path(self.file_path)
            output_dir = Path(self.output_path) if self.output_path else input_path.parent

            # Write log file to output directory
            log_path = output_dir / f"log_{timestamp}.log"
            file_handler = logging.FileHandler(log_path)
            file_handler.setFormatter(self.formatter)
            self.logger.addHandler(file_handler)

            def sanitize(s):
                result = "".join(c if c.isalnum() else "_" for c in str(s))
                while "__" in result:
                    result = result.replace("__", "_")
                return result.strip("_")
            
            state_val = "unknown"
            tag_str = ""
            stats = {"rows": 0, "files": 0, "written_files": []}

            def _write_file(complete = True):
                #WRITE per role — timestamp at finish time to avoid collisions
                if complete:
                    out_filename = output_dir / f"{sanitize(state_val)}_{sanitize(tag_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                else:
                    out_filename = output_dir / f"{sanitize(state_val)}_{sanitize(tag_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_incomplete{ext}"
                if ext == ".csv":
                    df.to_csv(out_filename, index=False)
                else:
                    df.to_excel(out_filename, index=False, engine="openpyxl")
                self.logger.info(f"Successfully wrote '{out_filename.name}'")
                stats["files"] += 1
                stats["written_files"].append(out_filename.name)
            # Pre-detect all sections to set whole-file progress bar maximum
            all_sheet_sections = {}
            total_rows = 0
            for sheet_name, df in sheets.items():
                secs = self._split_by_duplicate_headers(df, sheet_name)
                all_sheet_sections[sheet_name] = secs
                for _, sec_df in secs:
                    total_rows += len(sec_df)
            self.root.after(0, lambda t=total_rows: self.progress.config(maximum=max(t, 1), value=0))
            total_global_sections = sum(len(secs) for secs in all_sheet_sections.values())
            global_section_idx = 0

            for sheet_name, df in sheets.items():
                if run_id != self.current_run_id:
                    self.logger.info("New file selected — stopping previous run.")
                    return

                # Use pre-detected sections
                sections = all_sheet_sections[sheet_name]
                total_sections = len(sections)

                # Process each section (could be just one if no duplicate headers found)
                for section_idx, (section_name, section_df) in enumerate(sections, start=1):
                    global_section_idx += 1
                    if run_id != self.current_run_id:
                        self.logger.info("New file selected — stopping previous run.")
                        _write_file(False)
                        return

                    userChoice = None
                    # Use the section dataframe instead of the original df
                    df = section_df
                    name = section_name  # Use the section name for logging

                    # Detect columns, retry with skipped header rows if needed
                    cols = self._detect_columns(df)
                    if "County/City" not in cols or "Email" not in cols:
                        for skip in range(1, 5):
                            if ext == ".csv":
                                df = pd.read_csv(self.file_path, header=skip)
                            else:
                                # For sections, we can't re-read from file, skip this optimization
                                break
                            cols = self._detect_columns(df)
                            if "County/City" in cols and "Email" in cols:
                                self.logger.info(f"Found headers at row {skip}")
                                break

                    self.logger.info("Headers found: " + str(cols))

                    self.column_for["County/City"] = cols.get("County/City")
                    self.column_for["Email"] = cols.get("Email")
                    self.column_for["Phone Number"] = cols.get("Phone Number")
                    self.column_for["First Name"] = cols.get("First Name")
                    self.column_for["Last Name"] = cols.get("Last Name")
                    self.column_for["Role/Title"] = cols.get("Role/Title")
                    self.column_for["State"] = cols.get("State")
                    self.column_for["LinkedIn"] = cols.get("LinkedIn")
                    self.column_for["Tag"] = cols.get("Tag")
                    self.column_for["Contact Tag"] = cols.get("Contact Tag")


                    # Read state value now so it can be shown in the role selection dialog
                    state_col = self.column_for.get("State")
                    
                    if state_col and state_col in df.columns:
                        non_null = df[state_col].dropna()
                        non_null = non_null[non_null.astype(str).str.strip() != ""]
                        if not non_null.empty:
                            state_val = str(non_null.iloc[0]).strip()

                    userChoices = self.sheet_role_choices.get(sheet_name)
                    if userChoices is None:
                        if global_section_idx > 1:
                            self.root.after(0, lambda gi=global_section_idx, gt=total_global_sections: (
                                self.progress_label.config(
                                    text=f"Completed {gi - 1}/{gt} sections — waiting for user input...",
                                    style="Processing.TLabel"
                                ),
                            ))
                        try:
                            userChoices = self.select_role_to_search(name, sheet_name, run_id, state_val)
                        except TypeError:
                            self.logger.error("Invalid input, only 1, 2, or 3 accepted")
                        except ValueError:
                            self.logger.error("Not a valid role")

                    if userChoices is None:
                        return  # Cancelled by new file selection or input error

                    # Build CSV filename: stateprovince_tag_datetime.csv
                    tag_map = {1: "NG911", 2: "Mayor", 3: "QQ"}
                    out_filename = None

                    if userChoices == [0]:
                        #tag_str = "original"
                        #out_filename = output_dir / f"{sanitize(state_val)}_{sanitize(tag_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                        #if ext == ".csv":
                        #    df.to_csv(out_filename, index=False)
                        #else:
                        #    df.to_excel(out_filename, index=False, engine="openpyxl")
                        self.logger.info(f"Skipped processing for sheet '{name}'") # - wrote original data to {out_filename.name}")
                        #stats["files"] += 1
                        #stats["written_files"].append(out_filename.name)
                        continue

                    df_original = df.copy()

                    def insert_if_missing(df, idx, col_name, default=""):
                        if col_name.lower() not in [c.lower() for c in df.columns]:
                            df.insert(idx, col_name, default)

                    for userChoice in userChoices:
                        df = df_original.copy()  # fresh copy so roles don't overwrite each other
                        role = Role(userChoice)
                        tag_str = tag_map.get(userChoice, "data")
                        # out_filename is set at write time using finish timestamp to avoid collisions

                        self.root.after(0, lambda n=name, tag=tag_str, si=section_idx, st=total_sections: (
                            self.progress_label.config(text=f"Section {si}/{st} — Processing '{n}' ({tag})...", style="Processing.TLabel"),
                        ))

                        if self.column_for["County/City"]:
                            self.logger.info("Looking for counties or cities under column: " + self.column_for["County/City"])

                        insert_if_missing(df, len(df.columns), "Contact Tag")
                        insert_if_missing(df, len(df.columns), "Source")
                        insert_if_missing(df, len(df.columns), "Email Confidence")
                        insert_if_missing(df, len(df.columns), "Alternative Email")
                        insert_if_missing(df, len(df.columns), "Alternative Email Confidence")
                        insert_if_missing(df, len(df.columns), "Hunter Email Source")

                        #Prevents TypeErrors, clears stale output data
                        for col in [
                            self.column_for["First Name"],
                            self.column_for["Last Name"],
                            self.column_for["Email"],
                            self.column_for["Phone Number"],
                            self.column_for["Role/Title"],
                            self.column_for["LinkedIn"],
                            self.column_for["Contact Tag"],
                            "Source",
                            "Email Confidence",
                            "Alternative Email",
                            "Alternative Email Confidence",
                            "Hunter Email Source",
                        ]:
                            if col and col in df.columns:
                                df[col] = ""
                                #df[col] = df[col].astype("string")

                        section_incomplete_notified = False
                        rows_before = stats["rows"]

                        #Search and verify each column
                        try:
                            for idx, value in df[self.column_for["County/City"]].items():
                                if run_id != self.current_run_id:
                                    _write_file(False)
                                    return

                                # Skip rows that are mostly empty
                                non_empty = df.loc[idx].dropna().replace("", pd.NA).dropna()
                                if len(non_empty) <= 1:
                                    self.logger.info(f"Skipping mostly empty row {idx}")
                                    continue

                                # Skip if county/city value is blank or matches header
                                if pd.isna(value) or str(value).strip() == "" or value == self.column_for["County/City"]:
                                    continue

                                self.logger.info("Currently on: " + str(value))
                                stats["rows"] += 1
                                self.root.after(0, lambda: self.progress.config(
                                    value=min(self.progress['value'] + 1, self.progress['maximum'])
                                ))

                                match role:
                                    case Role.GIS:
                                        system_prompt = self.prompt_gis
                                    case Role.MAYOR:
                                        system_prompt = self.prompt_mayor
                                    case Role.ASSESSOR:
                                        system_prompt = self.prompt_assessor

                                try:
                                    info = openai_hunter_client.search(
                                        value + " " + df.loc[idx, self.column_for["State"]] + " Government",
                                        role,
                                        system_prompt
                                    )
                                except openai.APIConnectionError:
                                    raise
                                except Exception as e:
                                    self.logger.error(f"OpenAI search failed for row {idx} ({value}): {str(e)}")
                                    continue

                                info = (info or "").strip()
                                if not info or info == "None" or info is None:
                                    continue

                                try:
                                    parsedInfo = json.loads(info)
                                except json.JSONDecodeError:
                                    self.logger.info(f"OpenAI returned non-JSON for row {idx} ({value}), skipping.")
                                    continue

                                try:
                                    #Add the new name, email, role, phone number, and info source
                                    df.loc[idx, self.column_for["First Name"]] = parsedInfo.get("firstName", "")
                                    df.loc[idx, self.column_for["Last Name"]] = parsedInfo.get("lastName", "")
                                    df.loc[idx, self.column_for["Email"]] = parsedInfo.get("email", "")
                                    df.loc[idx, self.column_for["Phone Number"]] = parsedInfo.get("phoneNumber", "")
                                    df.loc[idx, self.column_for["Role/Title"]] = parsedInfo.get("role", "")
                                    if role == Role.ASSESSOR:
                                        df.loc[idx, self.column_for["Tag"]] = "QQ"
                                        df.loc[idx, "Contact Tag"] = "QQ"
                                    elif role == Role.GIS:
                                        df.loc[idx, self.column_for["Tag"]] = "NG911"
                                        df.loc[idx, "Contact Tag"] = "NG911"
                                    df.loc[idx, "Source"] = parsedInfo.get("sourceWebsite" , "")

                                    state_mapping = {item: label for lst, label in stateCorrectionMap.items() for item in lst}
                                    if df.loc[idx, self.column_for["State"]].lower() in state_mapping:
                                        df.loc[idx, self.column_for["State"]] = state_mapping.get(df.loc[idx, self.column_for["State"]].lower())

                                    reFind = False
                                    email_val = parsedInfo.get("email")
                                    email_type = parsedInfo.get("emailType")
                                    verifyNeeded = bool(email_val) and email_type == "person" and email_val != "None" and "*" not in email_val and "protected" not in email_val
                                    #Verify personal emails found
                                    if verifyNeeded:
                                        try:
                                            res = openai_hunter_client.verify_email(email_val)
                                            while str(res[0]) == "202" and attempt_count <= 5: #"The email verification is still in progress. To avoid the request running for too long we return HTTP 202 responses."
                                                res = openai_hunter_client.verify_email(email_val)
                                                attempt_count += 1
                                            if str(res[0]) == "200":
                                                verification_data = res[1].get("data")
                                                if verification_data:
                                                    score = verification_data.get("score", -1)
                                                    status = verification_data.get("status")
                                                    sources = verification_data.get("sources", [])
                                                    source = sources[0]["uri"] if sources else ""
                                                    self.logger.info("Data found by hunter.io:")
                                                    self.logger.info("Status: " + status)
                                                    if score >= 0:
                                                        self.logger.info("Score: " + str(score))
                                                        df.loc[idx, "Email Confidence"] = str(score)
                                                        df.loc[idx, "Hunter Email Source"] = source
                                                        self.logger.info("Source: " +  source)
                                                        if score < 80:
                                                            reFind = True
                                                            self.logger.info("Confidence score is too low, attempting to search again")
                                                else:
                                                    self.logger.error("Hunter.io did not return data")
                                                    
                                            elif str(res[0]) == "400":
                                                errors = res[1].get("errors") or []
                                                if errors and errors[0].get("id") == "invalid_email":
                                                    reFind = True
                                            else:
                                                self.logger.warning("Hunter.io Verify API Call Failed")
                                        except Exception as e:
                                            self.logger.warning(f"Hunter.io verify_email API error (skipping): {str(e)}")
                                            # Continue processing without verification

                                    if "protected" in email_val or "*" in email_val:
                                        reFind = True
                                    #Search for personal email if department email or no email was returned
                                    first_name = parsedInfo.get("firstName")
                                    last_name = parsedInfo.get("lastName")
                                    gov_site = parsedInfo.get("govWebsite")
                                    if (not verifyNeeded or reFind) and first_name and last_name and first_name.lower() not in ["none", "gis", "tax", "appraiser"] and last_name.lower() not in ["none", "gis", "tax", "team", "appraiser"] and gov_site:
                                        try:
                                            res = openai_hunter_client.find_email(first_name, last_name, gov_site)
                                            attempt_count = 1
                                            while str(res[0]) == "202" and attempt_count <= 5: #"The email verification is still in progress. To avoid the request running for too long we return HTTP 202 responses."
                                                res = openai_hunter_client.find_email(first_name, last_name, gov_site)
                                                attempt_count += 1
                                            if str(res[0]) == "200":
                                                parsedHunterResponse = res[1].get("data")
                                                if parsedHunterResponse:
                                                    email = parsedHunterResponse.get("email")
                                                    if email is not None:
                                                        score = parsedHunterResponse["score"]
                                                        sources = parsedHunterResponse.get("sources", [])
                                                        number = parsedHunterResponse.get("phone_number")
                                                        linkedin = parsedHunterResponse.get("linkedin_url")
                                                        source = sources[0]["uri"] if sources else ""
                                                        self.logger.info("Data found by hunter.io:")
                                                        self.logger.info("email: " + email)
                                                        self.logger.info("score: " + str(score))
                                                        self.logger.info("source: " + source)
                                                        if (reFind and score >= 90) or (not pd.isna(df.loc[idx, "Email Confidence"]) and str(df.loc[idx, "Email Confidence"]).strip() != "" and score > int(df.loc[idx, "Email Confidence"])) or pd.isna(df.loc[idx, self.column_for["Email"]]) or df.loc[idx, self.column_for["Email"]] == "" or df.loc[idx, self.column_for["Email"]] == 0:
                                                            self.logger.info(f"Saving hunter.io email {email} to email column, moving original email to Alternative Email column")
                                                            df.loc[idx, "Alternative Email"] = df.loc[idx, self.column_for["Email"]]
                                                            df.loc[idx, "Alternative Email Confidence"] = df.loc[idx, "Email Confidence"]   
                                                            df.loc[idx, self.column_for["Email"]] = email
                                                            df.loc[idx, "Email Confidence"] = str(score)
                                                            df.loc[idx, "Hunter Email Source"] = source
                                                            df.loc[idx, self.column_for["LinkedIn"]] = linkedin
                                                            if (pd.isna(df.loc[idx, self.column_for["Phone Number"]]) or df.loc[idx, self.column_for["Phone Number"]] == 0) and number != 0 and number is not None:
                                                                df.loc[idx, self.column_for["Phone Number"]] = number
                                                        elif score>=70:
                                                            self.logger.info(f"Email saving hunter.io email {email} to Alternative Email column")
                                                            df.loc[idx, "Alternative Email"] = email
                                                            df.loc[idx, "Alternative Email Confidence"] = str(score)
                                                            df.loc[idx, "Hunter Email Source"] = source
                                                            df.loc[idx, self.column_for["LinkedIn"]] = linkedin
                                                            if (pd.isna(df.loc[idx, self.column_for["Phone Number"]]) or df.loc[idx, self.column_for["Phone Number"]] == 0) and number != 0 and number is not None:
                                                                df.loc[idx, self.column_for["Phone Number"]] = number
                                                        else:
                                                            self.logger.info(f"Hunter.io email score less than 70, too low to save ({score}): {email}")
                                                    else:
                                                        self.logger.info("Hunter.io did not find an email for " + first_name + " " + last_name)
                                                else:
                                                    self.logger.error("Hunter.io did not return data")
                                            else:
                                                self.logger.info("Hunter.io Find API Call Failed")
                                        except Exception as e:
                                            self.logger.warning(f"Hunter.io find_email API error (skipping): {str(e)}")
                                            # Continue processing without finding alternative email
                                except TypeError as e:
                                    #This should not happen
                                    self.logger.error("TypeError:" + str(e))
                                    self.logger.warning("You may be missing a row of data in the output.")
                                    if not section_incomplete_notified:
                                        section_incomplete_notified = True
                                        self.root.after(0, lambda n=name, tag=tag_str: messagebox.showwarning(
                                            "Incomplete Data",
                                            f"Data is incomplete for section '{n}' ({tag}) and must be rerun."
                                        ))

                        except Exception as e:
                            self.logger.error("Error: " + str(e))
                            fname = out_filename.name if out_filename is not None else name
                            self.logger.info(f"Saving progress to {fname} after error")
                            if not section_incomplete_notified:
                                section_incomplete_notified = True
                                self.root.after(0, lambda n=name, tag=tag_str: messagebox.showwarning(
                                    "Incomplete Data",
                                    f"Data is incomplete for section '{n}' ({tag}) and must be rerun."
                                ))
                            if stats["rows"] == rows_before:
                                self.logger.warning(f"No processable rows found in section '{name}' ({tag_str})")
                            _write_file(False)
                            continue

                        if stats["rows"] == rows_before:
                            self.logger.warning(f"No processable rows found in section '{name}' ({tag_str})")
                        _write_file(True)

            self.logger.info(f"Run complete — {stats['rows']} rows processed, {stats['files']} file(s) written.")
            
            files_text = (
                f"No files written from {Path(self.file_path).name}" if stats["written_files"] == [] else "Files written: " + ", ".join(stats["written_files"]) + "\n" + f"\nfrom {Path(self.file_path).name}"
            )
            self.root.after(0, lambda ft=files_text: (
                self.select_output_btn.config(state="normal"),
                self.cancel_btn.pack_forget(),
                setattr(self, '_run_active', False),
                self.progress.config(value=self.progress['maximum']),
                self.progress_label.config(text="✓ Done!", style="Success.TLabel"),
                self.current_file_label.config(text=ft),
                self.open_folder_btn.pack(pady=(8, 6))
            ))

        except openai.APIConnectionError as e:
            self.logger.error(f"Connection error — run cancelled: {e}")
            self.root.after(0, lambda: (
                messagebox.showerror("Connection Error",
                    "Lost internet connection. The run has been cancelled.\n\nCheck your connection and try again."),
                self.select_output_btn.config(state="normal"),
                self.cancel_btn.pack_forget(),
                setattr(self, '_run_active', False),
                self.progress.config(value=0),
                self.progress_label.config(text="Connection error.", style="Gray.TLabel")
            ))

        except FileNotFoundError:
            self.logger.warning("File not found!")
            self.root.after(0, lambda: (
                self.select_output_btn.config(state="normal"),
                self.cancel_btn.pack_forget(),
                setattr(self, '_run_active', False),
                self.progress.config(value=0),
                self.progress_label.config(text="", style="Gray.TLabel")
            ))
        
        finally:
            if 'file_handler' in locals():
                self.logger.removeHandler(file_handler)
                file_handler.close()
            if run_id != self.current_run_id and 'stats' in locals():
                written = stats.get("written_files", [])
                cancelled_text = "Cancelled."
                if written:
                    cancelled_text += "\nFiles written: " + ", ".join(written) + f"\nfrom {Path(self.file_path).name}"
                self.root.after(0, lambda ct=cancelled_text: (
                    self.progress_label.config(text=""),
                    self.current_file_label.config(text=ct),
                ))

if __name__ == "__main__":
    try:
        App().run()
    except Exception as e:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(e), "Error", 1)
