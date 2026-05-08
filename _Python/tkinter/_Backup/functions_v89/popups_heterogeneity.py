"""
popups_heterogeneity.py — full-fidelity pure-Python port.

Adapted from Source_Py/popups_heterogeneity.py (~1500 LOC).  Replicates
the original 2-step wizard verbatim: aquitard configuration figures,
embedded-layer figures, boring-log table with CSV load/template,
manual-parameter entry, fractured-rock parallel-fracture entry, and
mdflag / volfrac / difflength computation.  Persists to
heterogeneity_inputs.txt — same format as the .exe pipeline.

Called from main.run_script() when:
    HeterogeneityCalculator_Unconsolidated_Media
    HeterogeneityCalculator_Fractured_Rock
"""
from __future__ import annotations
import os
import csv
import platform
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from PIL import Image, ImageTk
    _HAS_PIL = True
except ImportError:
    _HAS_PIL = False

from .state import get_state


MEDIA_UNCONS  = "Unconsolidated Media"
MEDIA_FRAC    = "Fractured Rock"

FT_TO_M       = 0.3048

FONT_TITLE    = ("Arial", 16, "bold")
FONT_STEP     = ("Arial", 13, "bold")
FONT_HEADER   = ("Arial", 12, "bold")
FONT_LABEL    = ("Arial", 11)
FONT_SMALL    = ("Arial", 10, "italic")
FONT_TINY     = ("Arial", 9, "bold")


# ── Canvas-drawn radio (mirrors big_radio in main.py Section 1) ─────────
# Windows locks tk.Radiobutton's circle to the OS-theme size regardless of
# font, so we draw our own.  RADIO_R is the radius in pixels.
def _big_radio(parent, text, variable, value, *,
               text_font=None, bg="#F0F0F0", wraplength=0):
    RADIO_R = 9                      # 18 px diameter — matches main.py
    if text_font is None:
        text_font = FONT_LABEL

    fr  = tk.Frame(parent, bg=bg)
    dia = RADIO_R * 2 + 4
    cv  = tk.Canvas(fr, width=dia, height=dia, bg=bg,
                    highlightthickness=0, bd=0)
    cv.pack(side="left", anchor="n", pady=2)

    lbl = tk.Label(fr, text=text, font=text_font, bg=bg, anchor="w",
                   wraplength=wraplength if wraplength > 0 else 0,
                   justify="left")
    lbl.pack(side="left", padx=(4, 0), fill="x", expand=True)

    def _draw(*_):
        cv.delete("all")
        cv.create_oval(2, 2, RADIO_R * 2 + 2, RADIO_R * 2 + 2,
                       outline="#555555", width=1.5)
        if variable.get() == value:
            inner = max(RADIO_R // 2, 3)
            cx = RADIO_R + 2
            cv.create_oval(cx - inner, cx - inner, cx + inner, cx + inner,
                           fill="#333333", outline="#333333")

    def _select(_event=None):
        variable.set(value)

    cv.bind("<Button-1>", _select)
    lbl.bind("<Button-1>", _select)
    variable.trace_add("write", _draw)
    _draw()
    return fr


FONT_BTN      = ("Arial", 11)


def _figures_dir():
    here = os.path.dirname(os.path.abspath(__file__))
    project = os.path.abspath(os.path.join(here, "..", "..", ".."))
    return os.path.join(project, "Figures")


def _docs_root():
    here = os.path.dirname(os.path.abspath(__file__))
    project = os.path.abspath(os.path.join(here, "..", "..", ".."))
    return os.path.join(project, "docs", "_site")


def _open_help_section(section_id: str = ""):
    f = os.path.join(_docs_root(),
                     "data_chicklets",
                     "Step4_HydrogeologicSettingAndMatrixDiffusion.html")
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found", f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    anchor = f"#{section_id}" if section_id else ""
    url = (f"file:///{abs_p}{anchor}" if os.name == "nt" and abs_p[1] == ":"
           else f"file://{abs_p}{anchor}")
    try:
        if platform.system() == "Windows":
            for exe in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
                if os.path.exists(exe):
                    subprocess.Popen([exe, url]); return
        webbrowser.open(url)
    except Exception:
        webbrowser.open(url)


def _open_help_appendix():
    f = os.path.join(_docs_root(), "appendix", "appendix_4_1.html")
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found", f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    url = (f"file:///{abs_p}" if os.name == "nt" and abs_p[1] == ":"
           else f"file://{abs_p}")
    try:
        webbrowser.open(url)
    except Exception:
        pass


class _HeterogeneityWizard:
    def __init__(self, root, app, media_type=MEDIA_UNCONS, unit_flag=2):
        self.root = root
        self.app = app
        self.media_type = (MEDIA_FRAC if str(media_type).lower().startswith("fractured")
                           else MEDIA_UNCONS)
        self.unit_flag = int(unit_flag) if unit_flag else 2
        self.length_unit = "ft" if self.unit_flag == 1 else "m"
        self.unit_to_m   = FT_TO_M if self.unit_flag == 1 else 1.0
        self.length_unit_bgs = f"{self.length_unit} bgs"
        self.figures_dir = _figures_dir()

        # State vars
        self.step1_var          = tk.StringVar()
        self.step2_choice_var   = tk.StringVar()
        self.embedded_layer_var = tk.StringVar()
        self.thickness_var      = tk.StringVar()
        self.manual_tvf_var     = tk.StringVar()
        self.manual_adl_var     = tk.StringVar()
        self.fracture_a_var     = tk.StringVar()
        self.fracture_b_var     = tk.StringVar()
        self.ss_num_wells_var   = tk.StringVar(value="1")
        self.ss_well_entries    = []
        self.ss_wells_frame     = None

        # Step 1 options (single-line; wraplength on the radio handles wrapping)
        self.step1_options = [
            "Option 1: No Matrix Diffusion in Under- and Overlying Low-k Units",
            "Option 2: Matrix Diffusion in Underlying Low-k Units",
            "Option 3: Matrix Diffusion in Overlying Low-k Units",
            "Option 4: Matrix Diffusion in Under- and Overlying Low-k Units",
        ]

        # Step 2 main options depend on media type
        if self.media_type == MEDIA_FRAC:
            self.step2_main_options = [
                "Option 1: Assume simple parallel fractures",
                "Option 2: Enter heterogeneity parameters manually",
            ]
        else:
            self.step2_main_options = [
                "Option 1: Select from Simple Examples",
                "Option 2: Enter Boring Log Data",
                "Option 3: Enter Heterogeneity Parameters Manually",
            ]

        self.step2_figure_options = [
            "Option 1: ~0% of Plume Thickness is in Low-k Material",
            "Option 2: ~20% of Plume Thickness is in Low-k Material",
            "Option 3: ~40% of Plume Thickness is in Low-k Material",
            "Option 4: ~60% of Plume Thickness is in Low-k Material",
            "Option 5: ~80% of Plume Thickness is in Low-k Material",
        ]
        self.step2_options = self.step2_figure_options + [
            "Use Site-Specific Data",
            "Enter Heterogeneity Parameters Manually",
        ]

        self.mdflag_mapping = self._create_mdflag_mapping()

        # Step navigation
        self.current_step = 1
        self.max_steps = 2
        self.step2_showing_main_options = True
        self.step2_current_option_index = None

        # Defaults
        self.step1_var.set(self.step1_options[0])
        self.step2_choice_var.set(self.step2_main_options[0])

        self._setup_ui()

    # ── mdflag table ────────────────────────────────────────────────────
    def _create_mdflag_mapping(self):
        m = {}
        # Top/Bottom Option 1: 0 if no embedded, else 2
        for j in range(5):
            m[f"1_{j+1}"] = 0 if j == 0 else 2
        # Option 2
        for j in range(5):
            m[f"2_{j+1}"] = 1 if j == 0 else 5
        # Option 3
        for j in range(5):
            m[f"3_{j+1}"] = 3 if j == 0 else 6
        # Option 4
        for j in range(5):
            m[f"4_{j+1}"] = 0 if j == 0 else 7
        return m

    # ── UI scaffolding ──────────────────────────────────────────────────
    def _setup_ui(self):
        title_extra = (" (Fractured Rock)" if self.media_type == MEDIA_FRAC
                       else " (Unconsolidated Media)")
        self.root.title("Heterogeneity Calculator" + title_extra)

        title_frame = tk.Frame(self.root, bg="#F0F0F0")
        title_frame.pack(pady=(12, 6))
        tk.Label(title_frame,
                 text="Heterogeneity Calculator" + title_extra,
                 font=FONT_TITLE, bg="#F0F0F0").pack()

        # Progress
        prog_frame = tk.Frame(self.root, bg="#F0F0F0")
        prog_frame.pack(pady=4, fill="x", padx=20)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(prog_frame,
                                             variable=self.progress_var,
                                             maximum=100)
        self.progress_bar.pack(fill="x", padx=4)

        # Scrollable content area
        canvas_container = tk.Frame(self.root, bg="#F0F0F0")
        canvas_container.pack(expand=True, fill="both", padx=20, pady=10)

        self.content_canvas = tk.Canvas(canvas_container, borderwidth=0,
                                         bg="#F0F0F0", highlightthickness=0)
        sb = tk.Scrollbar(canvas_container, orient="vertical",
                          command=self.content_canvas.yview)
        self.content_frame = tk.Frame(self.content_canvas, bg="#F0F0F0")
        self.canvas_window = self.content_canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw")

        def _scroll(event):
            self.content_canvas.configure(
                scrollregion=self.content_canvas.bbox("all"))
            self.content_canvas.itemconfig(self.canvas_window, width=event.width)
        self.content_frame.bind(
            "<Configure>",
            lambda e: self.content_canvas.configure(
                scrollregion=self.content_canvas.bbox("all")))
        self.content_canvas.bind("<Configure>", _scroll)
        self.content_canvas.configure(yscrollcommand=sb.set)
        self.content_canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Navigation
        self.nav_frame = tk.Frame(self.root, bg="#F0F0F0")
        self.nav_frame.pack(pady=(4, 12))
        self.prev_button = tk.Button(self.nav_frame, text="Previous",
                                      width=10, font=FONT_BTN,
                                      command=self._prev, state="disabled")
        self.prev_button.pack(side="left", padx=4)
        self.next_button = tk.Button(self.nav_frame, text="Next",
                                      width=10, font=FONT_BTN,
                                      command=self._next)
        self.next_button.pack(side="left", padx=4)
        tk.Button(self.nav_frame, text="Cancel", width=10, font=FONT_BTN,
                  command=self._cancel).pack(side="left", padx=4)
        tk.Button(self.nav_frame, text="Help", width=10, font=FONT_BTN,
                  command=_open_help_appendix).pack(side="left", padx=4)

        # PIL image cache (so they don't get GC'd)
        self._photo_refs = []
        self._show_step()

    # ── Image loading ───────────────────────────────────────────────────
    def _load_figure(self, parent, filename, label, max_w=900, max_h=900):
        if not _HAS_PIL:
            label.config(text=f"({os.path.basename(filename)})", fg="gray")
            return
        try:
            if not os.path.exists(filename):
                label.config(text=f"(missing: {os.path.basename(filename)})",
                             fg="gray")
                return
            image = Image.open(filename)
            if image.width > max_w or image.height > max_h:
                s = min(max_w / image.width, max_h / image.height)
                image = image.resize((int(image.width * s),
                                      int(image.height * s)),
                                     Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            label.config(image=photo, text="")
            self._photo_refs.append(photo)
        except Exception as e:
            label.config(text=f"(image error: {e})", fg="gray")

    # ── Step engine ─────────────────────────────────────────────────────
    def _show_step(self):
        for w in self.content_frame.winfo_children():
            w.destroy()
        self._photo_refs.clear()

        progress = (self.current_step / self.max_steps) * 100
        self.progress_var.set(progress)
        self.prev_button.config(
            state="normal" if self.current_step > 1 else "disabled")

        if self.current_step == 1:
            self._show_step1()
        elif self.current_step == 2:
            self._show_step2()

        if self.current_step == 2:
            if self.step2_showing_main_options:
                self.next_button.config(text="Next",
                                         command=self._proceed_from_main_options)
            else:
                self.next_button.config(text="Finish", command=self._finish)
        elif self.current_step == self.max_steps:
            self.next_button.config(text="Finish", command=self._finish)
        else:
            self.next_button.config(text="Next", command=self._next)

    def _show_step1(self):
        head = tk.Frame(self.content_frame, bg="#F0F0F0"); head.pack(pady=10)
        tk.Label(head, text="Step 1: Select Aquitard Configuration",
                 font=FONT_HEADER, bg="#F0F0F0").pack(side="left")
        tk.Button(head, text="?", width=2, height=1, fg="red", font=FONT_TINY,
                  command=lambda: _open_help_section(
                      "step-1-upper-and-lower-aquitards")
                  ).pack(side="left", padx=(6, 0))

        opts = tk.Frame(self.content_frame, bg="#F0F0F0")
        opts.pack(expand=True, fill="both", padx=24, pady=10)
        for i, option in enumerate(self.step1_options):
            cell = tk.Frame(opts, relief="raised", bd=2, bg="#F0F0F0")
            cell.grid(row=0, column=i, padx=10, pady=12, sticky="nsew",
                      ipadx=6, ipady=6)
            _big_radio(cell, option, self.step1_var, option,
                       text_font=FONT_LABEL, bg="#F0F0F0",
                       wraplength=300
                       ).pack(pady=8, padx=10, fill="x")
            fig_lbl = tk.Label(cell, text=f"Figure {i+1}", relief="sunken",
                                bd=1, bg="white")
            fig_lbl.pack(pady=8, padx=8, fill="both", expand=True)
            self._load_figure(cell,
                              os.path.join(self.figures_dir, f"Step1_{i+1}.png"),
                              fig_lbl)
        for i in range(4):
            opts.grid_columnconfigure(i, weight=1)

    def _show_step2(self, force_main=False):
        for w in self.content_frame.winfo_children():
            w.destroy()
        self._photo_refs.clear()

        if force_main:
            self.step2_showing_main_options = True
            self.step2_current_option_index = None
            self._show_step2_main_options(); return
        if self.step2_showing_main_options:
            self._show_step2_main_options(); return

        if (self.step2_current_option_index is not None and
                0 <= self.step2_current_option_index < len(self.step2_main_options)):
            main_choice = self.step2_main_options[self.step2_current_option_index]
        else:
            main_choice = self.step2_choice_var.get()
            if main_choice in self.step2_main_options:
                self.step2_current_option_index = self.step2_main_options.index(main_choice)
            else:
                self.step2_showing_main_options = True
                self._show_step2_main_options(); return

        self.step2_showing_main_options = False
        if self.media_type == MEDIA_FRAC:
            if main_choice == self.step2_main_options[0]:
                self._step2_fractured_option1()
            else:
                self._step2_manual()
        else:
            if main_choice == self.step2_main_options[0]:
                self._step2_unconsolidated_option1()
            elif main_choice == self.step2_main_options[1]:
                self._step2_unconsolidated_option2()
            else:
                self._step2_manual()

    def _show_step2_main_options(self):
        head = tk.Frame(self.content_frame, bg="#F0F0F0")
        head.pack(pady=20)
        if self.media_type == MEDIA_FRAC:
            txt = "Step 2: Select one of 2 options:"
        else:
            txt = "Step 2: Select one of 3 methods to add low-k layers/lenses:"
        tk.Label(head, text=txt, font=FONT_STEP, bg="#F0F0F0").pack()
        opts = tk.Frame(self.content_frame, bg="#F0F0F0")
        opts.pack(expand=True, fill="both", pady=20)
        for option in self.step2_main_options:
            _big_radio(opts, option, self.step2_choice_var, option,
                       text_font=FONT_HEADER, bg="#F0F0F0"
                       ).pack(pady=12, padx=20, anchor="w")

    def _proceed_from_main_options(self):
        choice = self.step2_choice_var.get()
        if not choice or choice not in self.step2_main_options:
            messagebox.showwarning(
                "Warning",
                "Please select one of the options before proceeding.",
                parent=self.root); return
        self.step2_current_option_index = self.step2_main_options.index(choice)
        self.step2_showing_main_options = False
        self._show_step2()
        self._show_step()

    def _step2_unconsolidated_option1(self):
        if (not self.embedded_layer_var.get()
                or self.embedded_layer_var.get() not in self.step2_figure_options):
            self.embedded_layer_var.set(self.step2_figure_options[0])

        head = tk.Frame(self.content_frame, bg="#F0F0F0"); head.pack(pady=10)
        tk.Label(head,
                 text=("Option 1: Pick an Option for the Embedded Layer "
                       "between top and bottom of your plume"),
                 font=FONT_HEADER, bg="#F0F0F0").pack(side="left")
        tk.Button(head, text="?", width=2, height=1, fg="red", font=FONT_TINY,
                  command=lambda: _open_help_section(
                      "step-2-embedded-conditions-simple-method")
                  ).pack(side="left", padx=(6, 0))
        tk.Label(self.content_frame,
                 text=("(Don't count the top and bottom Low-K layers; just "
                       "the in-between embedded layers.)"),
                 font=FONT_SMALL, bg="#F0F0F0").pack(pady=4)

        opts = tk.Frame(self.content_frame, bg="#F0F0F0")
        opts.pack(expand=True, fill="both", padx=24, pady=10)
        for i in range(5):
            cell = tk.Frame(opts, relief="raised", bd=2, bg="#F0F0F0")
            cell.grid(row=0, column=i, padx=10, pady=12, sticky="nsew",
                      ipadx=6, ipady=6)
            _big_radio(cell, self.step2_figure_options[i],
                       self.embedded_layer_var,
                       self.step2_figure_options[i],
                       text_font=FONT_LABEL, bg="#F0F0F0",
                       wraplength=240
                       ).pack(pady=6, padx=8, fill="x")
            fig = tk.Label(cell, text=f"Figure {chr(65+i)}",
                           relief="sunken", bd=1, bg="white")
            fig.pack(pady=8, padx=6, fill="both", expand=True)
            self._load_figure(cell,
                              os.path.join(self.figures_dir, f"Step2_{i+1}.png"),
                              fig, max_w=540, max_h=720)
        for i in range(5):
            opts.grid_columnconfigure(i, weight=1)

        thickness = tk.Frame(self.content_frame, bg="#F0F0F0")
        thickness.pack(pady=15, fill="x")
        row = tk.Frame(thickness, bg="#F0F0F0"); row.pack(pady=5)
        tk.Label(row, text="Enter Typical Thickness of Embedded Layers:",
                 font=FONT_HEADER, bg="#F0F0F0"
                 ).pack(side="left", padx=(0, 10))
        tk.Label(row, text=f"Typical Thickness ({self.length_unit}):",
                 font=FONT_LABEL, bg="#F0F0F0").pack(side="left", padx=(0, 5))
        tk.Entry(row, textvariable=self.thickness_var, font=FONT_LABEL,
                 width=14).pack(side="left", padx=(0, 5))
        tk.Button(row, text="?", width=2, height=1, fg="red", font=FONT_TINY,
                  command=lambda: _open_help_section(
                      "step-3-embedded-thickness-simple-method")
                  ).pack(side="left", padx=(5, 0))
        tk.Label(thickness,
                 text="Half of this value will be used as diffusion length.",
                 font=FONT_SMALL, bg="#F0F0F0").pack(pady=2)

    def _step2_unconsolidated_option2(self):
        head = tk.Frame(self.content_frame, bg="#F0F0F0"); head.pack(pady=10)
        tk.Label(head,
                 text="Option 2: Estimate Site-Specific Heterogeneity Parameters",
                 font=FONT_HEADER, bg="#F0F0F0").pack(side="left")
        tk.Button(head, text="?", width=2, height=1, fg="red", font=FONT_TINY,
                  command=lambda: _open_help_section(
                      "step-3-embedded-thickness-site-specific-boring-logs")
                  ).pack(side="left", padx=(6, 0))

        ctrl = tk.Frame(self.content_frame, bg="#F0F0F0"); ctrl.pack(pady=5)
        tk.Label(ctrl, text="Number of Wells (<=100):", font=FONT_LABEL,
                 bg="#F0F0F0").pack(side="left")
        tk.Entry(ctrl, textvariable=self.ss_num_wells_var, width=6,
                 font=FONT_LABEL).pack(side="left", padx=6)
        tk.Button(ctrl, text="Apply", command=self._build_ss_grid,
                  font=FONT_LABEL).pack(side="left", padx=6)
        tk.Button(ctrl, text="Load CSV", command=self._load_csv_file,
                  font=FONT_LABEL).pack(side="left", padx=6)
        tk.Button(ctrl, text="Download Template",
                  command=self._download_template, font=FONT_LABEL
                  ).pack(side="left", padx=6)

        cont = tk.Frame(self.content_frame, bg="#F0F0F0")
        cont.pack(expand=True, fill="both", pady=10)
        cv = tk.Canvas(cont, borderwidth=0, bg="#F0F0F0", highlightthickness=0)
        sb = tk.Scrollbar(cont, orient="vertical", command=cv.yview)
        self.ss_wells_frame = tk.Frame(cv, bg="#F0F0F0")
        self.ss_wells_frame.bind("<Configure>",
            lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=self.ss_wells_frame, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._build_ss_grid()

    def _step2_fractured_option1(self):
        head = tk.Frame(self.content_frame, bg="#F0F0F0"); head.pack(pady=10)
        tk.Label(head, text="Enter Fracture Data", font=FONT_HEADER,
                 bg="#F0F0F0").pack(side="left")
        body = tk.Frame(self.content_frame, bg="#F0F0F0")
        body.pack(expand=True, fill="both", pady=20)

        fig_lbl = tk.Label(body, text="(Step3_1.png)", relief="sunken",
                            bd=1, bg="white")
        fig_lbl.pack(pady=(0, 12))
        self._load_figure(body,
                          os.path.join(self.figures_dir, "Step3_1.png"),
                          fig_lbl, max_w=560, max_h=420)

        r1 = tk.Frame(body, bg="#F0F0F0"); r1.pack(pady=8)
        tk.Label(r1,
                 text=f'Typical distance between parallel fractures ("a") ({self.length_unit}):',
                 font=FONT_LABEL, bg="#F0F0F0").pack(side="left")
        tk.Entry(r1, textvariable=self.fracture_a_var, font=FONT_LABEL,
                 width=12).pack(side="left", padx=10)

        r2 = tk.Frame(body, bg="#F0F0F0"); r2.pack(pady=8)
        tk.Label(r2,
                 text=f'Typical thickness of aperture/fracture ("b") ({self.length_unit}):',
                 font=FONT_LABEL, bg="#F0F0F0").pack(side="left")
        tk.Entry(r2, textvariable=self.fracture_b_var, font=FONT_LABEL,
                 width=12).pack(side="left", padx=10)

        tk.Label(body,
                 text=(f"Volfrac = b/a x 100 (%).  Average diffusion length "
                       f"= (a-b)/2 ({self.length_unit})."),
                 font=FONT_SMALL, bg="#F0F0F0").pack(pady=10)

    def _step2_manual(self):
        head = tk.Frame(self.content_frame, bg="#F0F0F0"); head.pack(pady=10)
        tk.Label(head,
                 text="Option: Enter Heterogeneity Parameters Manually",
                 font=FONT_HEADER, bg="#F0F0F0").pack(side="left")
        tk.Button(head, text="?", width=2, height=1, fg="red", font=FONT_TINY,
                  command=lambda: _open_help_section(
                      "step-3-embedded-thickness-manual-entry")
                  ).pack(side="left", padx=(6, 0))

        body = tk.Frame(self.content_frame, bg="#F0F0F0")
        body.pack(expand=True, fill="both", pady=20)

        r1 = tk.Frame(body, bg="#F0F0F0"); r1.pack(pady=8)
        tk.Label(r1, text="Transmissive Zone Volume Fraction (%)",
                 font=FONT_LABEL, bg="#F0F0F0").pack(side="left")
        tk.Entry(r1, textvariable=self.manual_tvf_var, font=FONT_LABEL,
                 width=12).pack(side="left", padx=10)

        r2 = tk.Frame(body, bg="#F0F0F0"); r2.pack(pady=8)
        tk.Label(r2, text=f"Average Diffusion Length ({self.length_unit})",
                 font=FONT_LABEL, bg="#F0F0F0").pack(side="left")
        tk.Entry(r2, textvariable=self.manual_adl_var, font=FONT_LABEL,
                 width=12).pack(side="left", padx=10)

        tk.Label(body,
                 text=("If provided, diffusion length overrides the "
                       "automatic half-thickness value."),
                 font=FONT_SMALL, bg="#F0F0F0").pack(pady=4)

    # ── Site-specific grid ──────────────────────────────────────────────
    def _build_ss_grid(self):
        try:
            n = int(self.ss_num_wells_var.get())
            if n < 1: n = 1
            if n > 100:
                n = 100; self.ss_num_wells_var.set(str(n))
        except ValueError:
            n = 1; self.ss_num_wells_var.set("1")

        self.ss_well_entries = []
        if self.ss_wells_frame:
            for w in self.ss_wells_frame.winfo_children():
                w.destroy()

        col_widths = [4, 10, 11] + [9] * 10
        header = tk.Frame(self.ss_wells_frame, bg="lightgray",
                          relief="raised", bd=1)
        header.pack(fill="x", pady=(0, 2))
        tk.Label(header, text="Well #", width=col_widths[0],
                 font=FONT_TINY, bg="lightgray", anchor="center"
                 ).grid(row=0, column=0, padx=1, pady=1, sticky="ew")
        tk.Label(header, text=f"Plume Top\n({self.length_unit_bgs})",
                 width=col_widths[1], font=FONT_TINY, bg="lightgray",
                 anchor="center").grid(row=0, column=1, padx=1, pady=1,
                                        sticky="ew")
        tk.Label(header, text=f"Plume Bottom\n({self.length_unit_bgs})",
                 width=col_widths[2], font=FONT_TINY, bg="lightgray",
                 anchor="center").grid(row=0, column=2, padx=1, pady=1,
                                        sticky="ew")
        for k in range(10):
            tk.Label(header, text=f"Low-K {k+1}\nThick. ({self.length_unit})",
                     width=col_widths[3+k], font=FONT_TINY, bg="lightgray",
                     anchor="center").grid(row=0, column=3+k, padx=1, pady=1,
                                            sticky="ew")
        for i in range(13):
            header.grid_columnconfigure(i, weight=1)

        for wi in range(n):
            row_color = "white" if wi % 2 == 0 else "#E0EFFF"
            row = tk.Frame(self.ss_wells_frame, bg=row_color, relief="flat", bd=1)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=str(wi + 1), width=col_widths[0],
                     font=("Arial", 8), bg=row_color, anchor="center"
                     ).grid(row=0, column=0, padx=1, pady=1, sticky="ew")

            top_var = tk.StringVar()
            tk.Entry(row, textvariable=top_var, width=col_widths[1],
                     font=("Arial", 8), justify="center"
                     ).grid(row=0, column=1, padx=1, pady=1, sticky="ew")
            bot_var = tk.StringVar()
            tk.Entry(row, textvariable=bot_var, width=col_widths[2],
                     font=("Arial", 8), justify="center"
                     ).grid(row=0, column=2, padx=1, pady=1, sticky="ew")
            lowk_vars = []
            for k in range(10):
                v = tk.StringVar()
                tk.Entry(row, textvariable=v, width=col_widths[3+k],
                         font=("Arial", 8), justify="center"
                         ).grid(row=0, column=3+k, padx=1, pady=1, sticky="ew")
                lowk_vars.append(v)
            for i in range(13):
                row.grid_columnconfigure(i, weight=1)
            self.ss_well_entries.append({"top": top_var, "bottom": bot_var,
                                          "lowk": lowk_vars})

    def _load_csv_file(self):
        try:
            path = filedialog.askopenfilename(
                title="Select CSV File",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                parent=self.root)
            if not path:
                return
            with open(path, "r", newline="", encoding="utf-8") as f:
                rows = list(csv.reader(f))
            if not rows:
                messagebox.showerror("Error", "CSV file is empty",
                                     parent=self.root); return
            data_rows = (rows[1:] if len(rows) > 1 and any(c.strip() for c in rows[0])
                         else rows)
            if not data_rows:
                messagebox.showerror("Error", "No data rows found",
                                     parent=self.root); return
            if len(data_rows[0]) < 3:
                messagebox.showerror("Error",
                    "CSV needs at least 3 columns (Well, Top, Bottom).",
                    parent=self.root); return
            n = min(len(data_rows), 100)
            if len(data_rows) > 100:
                messagebox.showwarning("Warning",
                    f"CSV has {len(data_rows)} rows; only first 100 loaded.",
                    parent=self.root)
            self.ss_num_wells_var.set(str(n))
            self._build_ss_grid()
            for i, row_data in enumerate(data_rows[:n]):
                if i >= len(self.ss_well_entries): break
                if len(row_data) > 1 and row_data[1].strip():
                    self.ss_well_entries[i]["top"].set(row_data[1].strip())
                if len(row_data) > 2 and row_data[2].strip():
                    self.ss_well_entries[i]["bottom"].set(row_data[2].strip())
                for k in range(min(10, len(row_data) - 3)):
                    if row_data[3 + k].strip():
                        self.ss_well_entries[i]["lowk"][k].set(row_data[3 + k].strip())
            messagebox.showinfo("Success", f"Loaded {n} wells from CSV.",
                                parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}",
                                 parent=self.root)

    def _download_template(self):
        try:
            path = filedialog.asksaveasfilename(
                title="Save CSV Template", defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile="well_data_template.csv", parent=self.root)
            if not path:
                return
            u  = self.length_unit
            ub = self.length_unit_bgs
            header = (f"Well,Plume Top ({ub}),Plume Bottom ({ub}),"
                      + ",".join([f"Low-K {k+1} Thick. ({u})" for k in range(10)]))
            sample_rows = [
                "1,12.5,18.2,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1",
                "2,10.8,16.5,0.1,0.1,0.1,0.1,0.1,0,0,0,0,0",
                "3,14.2,20.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0,0,0",
            ]
            with open(path, "w", newline="", encoding="utf-8") as f:
                f.write(header + "\n")
                for r in sample_rows:
                    f.write(r + "\n")
            messagebox.showinfo("Success",
                                f"Template saved to:\n{path}",
                                parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save template: {e}",
                                 parent=self.root)

    # ── Computation helpers ─────────────────────────────────────────────
    def _calc_volfrac(self, choice):
        if choice in self.step2_figure_options:
            mapping = [1.0, 0.8, 0.6, 0.4, 0.2]
            return mapping[self.step2_figure_options.index(choice)]
        if choice == "Use Site-Specific Data":
            return self._calc_ss_volfrac()
        if choice == "Enter Heterogeneity Parameters Manually":
            try: return float(self.manual_tvf_var.get()) / 100.0
            except ValueError: return 0.0
        return 0.0

    def _calc_ss_volfrac(self):
        if not self.ss_well_entries:
            return 0.0
        total = 0.0; n = 0
        for w in self.ss_well_entries:
            try:
                t = float(w["top"].get()) if w["top"].get().strip() else None
                b = float(w["bottom"].get()) if w["bottom"].get().strip() else None
            except ValueError:
                continue
            if t is None or b is None:
                continue
            t *= self.unit_to_m; b *= self.unit_to_m
            thick = b - t
            if thick <= 0:
                continue
            total_lk = 0.0
            for v in w["lowk"]:
                try:
                    total_lk += float(v.get()) * self.unit_to_m if v.get().strip() else 0
                except ValueError:
                    continue
            total += (1 - total_lk / thick); n += 1
        return total / n if n else 0.0

    def _calc_ss_difflen(self):
        if not self.ss_well_entries:
            return 0.0
        total = 0.0; n = 0
        for w in self.ss_well_entries:
            tot = 0.0; cnt = 0
            for v in w["lowk"]:
                raw = v.get()
                if not raw or not raw.strip():
                    continue
                try:
                    tot += float(raw) * self.unit_to_m; cnt += 1
                except ValueError:
                    continue
            if cnt == 0:
                continue
            total += (tot / cnt) / 2.0; n += 1
        return total / n if n else 0.0

    # ── Navigation ──────────────────────────────────────────────────────
    def _next(self):
        if self.current_step < self.max_steps:
            self.current_step += 1
            self._show_step()

    def _prev(self):
        if self.current_step == 2:
            if self.step2_showing_main_options:
                self.current_step -= 1
                self._show_step()
            else:
                self.step2_current_option_index = None
                self._show_step2(force_main=True)
                self._show_step()
        elif self.current_step > 1:
            self.current_step -= 1
            self._show_step()

    def _format_difflen(self, m):
        try:
            mv = float(m)
        except (TypeError, ValueError):
            return str(m)
        if self.unit_flag == 1:
            return f"{mv / FT_TO_M:.3f} {self.length_unit} ({mv:.3f} m)"
        return f"{mv:.3f} m"

    def _save(self, mdflag, difflen, volfrac):
        state = get_state()
        work_dir = state.work_dir or os.getcwd()
        path = os.path.join(work_dir, "heterogeneity_inputs.txt")
        if os.path.exists(path):
            try:
                os.chmod(path, 0o666); os.remove(path)
            except Exception:
                pass
        with open(path, "w") as f:
            f.write("Heterogeneity Calculator Results\n")
            f.write(f"mdflag: {mdflag}\n")
            f.write(f"Transmissive Fraction of Model (-): {volfrac}\n")
            f.write(f"Diffusion Length (m): {difflen:.2f}\n")

    def _finish(self):
        try:
            if not self.step1_var.get():
                messagebox.showerror("Error",
                    "Please select a top/bottom condition in Step 1",
                    parent=self.root); return
            if not self.step2_choice_var.get():
                messagebox.showerror("Error",
                    "Please select an option in Step 2", parent=self.root)
                return
            main_choice = self.step2_choice_var.get()

            # Fractured Rock Option 1: parallel fractures
            if (self.media_type == MEDIA_FRAC
                    and main_choice == self.step2_main_options[0]):
                if not self.fracture_a_var.get() or not self.fracture_b_var.get():
                    messagebox.showerror("Error",
                        'Please enter both "a" and "b".',
                        parent=self.root); return
                try:
                    a = float(self.fracture_a_var.get())
                    b = float(self.fracture_b_var.get())
                except ValueError:
                    messagebox.showerror("Error", "a/b must be numeric.",
                                         parent=self.root); return
                if a <= 0 or b < 0:
                    messagebox.showerror("Error",
                        "a > 0 and b >= 0 required.", parent=self.root); return
                if b > a:
                    messagebox.showerror("Error", "b cannot exceed a.",
                                         parent=self.root); return
                if a == b:
                    messagebox.showerror("Error",
                        "a must be greater than b.", parent=self.root); return
                a_m = a * self.unit_to_m; b_m = b * self.unit_to_m
                volfrac = b_m / a_m
                difflen = (a_m - b_m) / 2.0
                step1_idx = self.step1_options.index(self.step1_var.get()) + 1
                if volfrac == 0:
                    md_map = {1: 0, 2: 1, 3: 3, 4: 4}
                else:
                    md_map = {1: 2, 2: 5, 3: 6, 4: 7}
                mdflag = md_map.get(step1_idx, 0)
                self._save(mdflag, difflen, volfrac)
                messagebox.showinfo(
                    "Heterogeneity Calculator Results",
                    f"mdflag: {mdflag}\n"
                    f"volfrac: {volfrac:.3f} ({volfrac*100:.1f}%)\n"
                    f"difflength: {self._format_difflen(difflen)}",
                    parent=self.root)
                self.root.destroy(); return

            # Fractured Rock Option 2 == Manual
            if (self.media_type == MEDIA_FRAC
                    and main_choice == self.step2_main_options[1]):
                if not self.manual_tvf_var.get() or not self.manual_adl_var.get():
                    messagebox.showerror("Error",
                        "Please enter both manual values.",
                        parent=self.root); return
                choice = "Enter Heterogeneity Parameters Manually"
                is_figure = False; is_ss = False; is_manual = True
            else:
                # Unconsolidated branches
                is_o1 = main_choice == self.step2_main_options[0]
                is_o2 = main_choice == self.step2_main_options[1]
                is_o3 = (len(self.step2_main_options) > 2
                         and main_choice == self.step2_main_options[2])
                if is_o1:
                    if not self.embedded_layer_var.get() or not self.thickness_var.get():
                        messagebox.showerror("Error",
                            "Please pick an embedded layer option AND a thickness.",
                            parent=self.root); return
                    choice = self.embedded_layer_var.get()
                    is_figure = True; is_ss = False; is_manual = False
                elif is_o2:
                    choice = "Use Site-Specific Data"
                    is_figure = False; is_ss = True; is_manual = False
                elif is_o3:
                    if not self.manual_tvf_var.get() or not self.manual_adl_var.get():
                        messagebox.showerror("Error",
                            "Please enter both manual values.",
                            parent=self.root); return
                    choice = "Enter Heterogeneity Parameters Manually"
                    is_figure = False; is_ss = False; is_manual = True
                else:
                    messagebox.showerror("Error",
                        "Invalid Step 2 selection.", parent=self.root); return

            volfrac = self._calc_volfrac(choice)
            step1_idx = self.step1_options.index(self.step1_var.get()) + 1
            if is_figure:
                emb_idx = self.step2_figure_options.index(choice) + 1
                mdflag = self.mdflag_mapping.get(f"{step1_idx}_{emb_idx}", 0)
            else:
                if volfrac == 0:
                    md_map = {1: 0, 2: 1, 3: 3, 4: 4}
                else:
                    md_map = {1: 2, 2: 5, 3: 6, 4: 7}
                mdflag = md_map.get(step1_idx, 0)

            try:
                if is_manual:
                    difflen = float(self.manual_adl_var.get()) * self.unit_to_m
                elif is_figure:
                    thick = float(self.thickness_var.get()) * self.unit_to_m
                    difflen = thick / 2.0
                else:
                    difflen = self._calc_ss_difflen()
            except ValueError:
                messagebox.showerror("Error", "Invalid numeric value.",
                                     parent=self.root); return

            self._save(mdflag, difflen, volfrac)
            messagebox.showinfo(
                "Heterogeneity Calculator Results",
                f"mdflag: {mdflag}\n"
                f"volfrac: {volfrac:.3f}\n"
                f"difflength: {self._format_difflen(difflen)}",
                parent=self.root)
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}",
                                 parent=self.root)

    def _cancel(self):
        try: self.root.grab_release()
        except Exception: pass
        self.root.destroy()


def run(app, parent=None, media_type=MEDIA_UNCONS):
    root = tk.Toplevel(parent or app)
    root.configure(bg="#F0F0F0")
    try: root.withdraw()
    except Exception: pass
    # NOTE: NO transient() — that hides the maximize button on Windows.
    # grab_set() alone is enough for modal behaviour.
    try:
        root.grab_set()
    except Exception: pass

    # Determine unit flag from app
    unit_flag = 1 if (getattr(app, "v_units", None)
                      and app.v_units.get() == "feet") else 2

    _HeterogeneityWizard(root, app, media_type=media_type, unit_flag=unit_flag)

    root.protocol("WM_DELETE_WINDOW",
                  lambda: (root.grab_release() if root else None,
                           root.destroy())[-1] if False else root.destroy())
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 1920)
    h = max(root.winfo_reqheight() + 24, 1080)
    try:
        sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
        w = min(w, int(sw * 0.95)); h = min(h, int(sh * 0.92))
        x = max(0, (sw - w) // 2); y = max(0, (sh - h) // 2 - 20)
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        root.geometry(f"{w}x{h}")
    root.minsize(min(w, 1000), min(h, 700))
    root.resizable(True, True)
    try:
        root.deiconify(); root.lift(); root.focus_force()
    except Exception: pass
    root.wait_window()
