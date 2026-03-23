"""
HTML to Office Converter
GitHub: https://github.com/Hamza-op
"""

import os
import sys
import math
import threading
import webbrowser
import time
from pathlib import Path
from tkinter import filedialog, messagebox, Canvas
from typing import Optional

import customtkinter as ctk
from PIL import Image, ImageTk

import converter

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ═══════════════════════════════════════════════════════════
#  DESIGN SYSTEM
# ═══════════════════════════════════════════════════════════

SP = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "2xl": 32, "3xl": 40, "4xl": 48}

# Colors — warm amber/copper palette
PAL = {
    "root":         ("#FAF9F7", "#08090C"),
    "surface_0":    ("#F3F1EE", "#0E1015"),
    "surface_1":    ("#FFFFFF", "#16181F"),
    "surface_2":    ("#F5F3F0", "#1C1F28"),
    "surface_3":    ("#E8E5E0", "#252830"),
    "overlay":      ("#FFFFFF", "#1A1D26"),
    "canvas":       ("#E8E5E0", "#030406"),
    "inset":        ("#FAF8F6", "#0A0C12"),

    "accent":       "#D97706",     # Warm Amber
    "accent_h":     "#F59E0B",
    "accent_dim":   "#B45309",
    "accent_deep":  "#92400E",
    "accent_bg":    ("#FFFBEB", "#1C1308"),
    "accent_soft":  ("#FEF3C7", "#271E05"),
    "accent_ghost": ("#FFFBEB", "#1C1308"),

    "danger":       "#DC2626",
    "danger_h":     "#EF4444",
    "danger_bg":    ("#FEF2F2", "#2A0A0A"),
    "warn":         "#EA580C",
    "info":         "#0284C7",

    "t1":           ("#1C1917", "#FAFAF9"),
    "t2":           ("#44403C", "#D6D3D1"),
    "t3":           ("#57534E", "#A8A29E"),
    "t4":           ("#78716C", "#6E6862"),
    "t5":           ("#A8A29E", "#44403C"),

    "b1":           ("#D6D3D1", "#2E2C2A"),
    "b2":           ("#E7E5E4", "#1F1E1C"),
    "b3":           ("#D1CCC5", "#3B3836"),
    "b_accent":     ("#FDE68A", "#451A03"),
    "b_focus":      ("#D97706", "#D97706"),
}

FONT = "Segoe UI"
MONO = "Cascadia Code"
GITHUB_URL = "https://github.com/Hamza-op"


# ═══════════════════════════════════════════════════════════
#  CUSTOM WIDGETS
# ═══════════════════════════════════════════════════════════


class HoverCard(ctk.CTkFrame):
    """Frame that glows on hover with smooth border transition."""

    def __init__(self, master, hover_border=None, hover_bg=None, **kw):
        super().__init__(master, **kw)
        self._def_border = kw.get("border_color", PAL["b2"])
        self._def_bg = kw.get("fg_color", PAL["surface_1"])
        self._hov_border = hover_border or PAL["b1"]
        self._hov_bg = hover_bg or self._def_bg
        self.bind("<Enter>", lambda e: self.configure(
            border_color=self._hov_border, fg_color=self._hov_bg))
        self.bind("<Leave>", lambda e: self.configure(
            border_color=self._def_border, fg_color=self._def_bg))


class PulseButton(ctk.CTkButton):
    """Convert button with pulsing glow when idle."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self._pulse = True
        self._pulse_idx = 0
        self._colors = ["#D97706", "#B45309", "#92400E", "#B45309"]
        self._do_pulse()

    def _do_pulse(self):
        if self._pulse and str(self.cget("state")) == "normal":
            self.configure(fg_color=self._colors[self._pulse_idx % len(self._colors)])
            self._pulse_idx += 1
        self.after(400, self._do_pulse)

    def set_loading(self, loading: bool):
        self._pulse = not loading


class FormatCard(ctk.CTkFrame):
    """Toggle-style format card with check indicator."""

    def __init__(self, master, icon, title, desc, value, variable, on_change=None, **kw):
        super().__init__(master, corner_radius=16, border_width=2, cursor="hand2", **kw)
        self._val = value
        self._var = variable
        self._cmd = on_change

        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=SP["lg"], pady=SP["md"])

        top = ctk.CTkFrame(inner, fg_color="transparent")
        top.pack(fill="x")

        self._ico = ctk.CTkLabel(top, text=icon, font=ctk.CTkFont(size=32))
        self._ico.pack(side="left")

        col = ctk.CTkFrame(top, fg_color="transparent")
        col.pack(side="left", padx=(SP["md"], 0), fill="x", expand=True)

        self._ttl = ctk.CTkLabel(col, text=title,
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
            text_color=PAL["t1"], anchor="w")
        self._ttl.pack(anchor="w")

        self._dsc = ctk.CTkLabel(col, text=desc,
            font=ctk.CTkFont(family=FONT, size=10),
            text_color=PAL["t4"], anchor="w")
        self._dsc.pack(anchor="w")

        self._chk = ctk.CTkLabel(top, text="",
            font=ctk.CTkFont(size=18, weight="bold"), width=28)
        self._chk.pack(side="right")

        for w in [self, inner, top, self._ico, col, self._ttl, self._dsc]:
            w.bind("<Button-1>", self._click)

    def _click(self, e=None):
        self._var.set(self._val)
        if self._cmd:
            self._cmd()

    def refresh(self):
        on = self._var.get() == self._val
        if on:
            self.configure(fg_color=PAL["accent_bg"], border_color=PAL["accent"])
            self._chk.configure(text="✓", text_color=PAL["accent"])
        else:
            self.configure(fg_color=PAL["surface_1"], border_color=PAL["b2"])
            self._chk.configure(text="")


class StatusDot(ctk.CTkFrame):
    """Tiny pulsing status indicator."""

    def __init__(self, master, color="#8B5CF6", size=8, pulse=False, **kw):
        super().__init__(master, width=size, height=size,
                         corner_radius=size, fg_color=color, **kw)
        self._c = color
        self._pulse = pulse
        self._on = True
        if pulse:
            self._animate()

    def _animate(self):
        if not self._pulse:
            return
        self._on = not self._on
        self.configure(fg_color=self._c if self._on else PAL["surface_1"])
        self.after(600, self._animate)

    def set_color(self, c):
        self._c = c
        self.configure(fg_color=c)


class Toast(ctk.CTkFrame):
    """Floating notification that auto-dismisses."""

    @staticmethod
    def show(master, msg, level="success", duration=3500):
        colors = {
            "success": (PAL["accent"], PAL["accent_bg"]),
            "error":   (PAL["danger"], PAL["danger_bg"]),
            "warn":    (PAL["warn"],   ("#FFFBEB", "#271E05")),
        }
        icons = {"success": "✓", "error": "✗", "warn": "⚠"}
        accent, bg = colors.get(level, colors["success"])
        icon = icons.get(level, "→")

        t = ctk.CTkFrame(master, fg_color=bg, corner_radius=16,
                          border_width=2, border_color=accent)
        inner = ctk.CTkFrame(t, fg_color="transparent")
        inner.pack(padx=SP["xl"], pady=SP["md"])

        ctk.CTkLabel(inner, text=icon,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=accent).pack(side="left")
        ctk.CTkLabel(inner, text=msg,
            font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
            text_color=PAL["t1"]).pack(side="left", padx=(SP["sm"], 0))

        t.place(relx=0.5, rely=0.04, anchor="n")
        t.lift()

        def dismiss():
            try:
                t.place_forget()
                t.destroy()
            except Exception:
                pass

        t.after(duration, dismiss)
        return t


# ═══════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════


class Sidebar(ctk.CTkFrame):
    def __init__(self, parent, app, **kw):
        super().__init__(parent, fg_color=PAL["surface_0"], corner_radius=0, width=360, **kw)
        self.pack_propagate(False)
        self.app = app
        self._pad = SP["lg"]

        scr = ctk.CTkFrame(self, fg_color="transparent")
        scr.pack(fill="both", expand=True)
        self._scr = scr

        # ── INPUT ──
        self._section("INPUT FILES", "📄")

        self.drop_zone = ctk.CTkFrame(scr, fg_color=PAL["surface_1"],
            corner_radius=12, border_width=1, border_color=PAL["b2"])
        self.drop_zone.pack(fill="x", padx=self._pad, pady=(0, SP["sm"]))

        dz = ctk.CTkFrame(self.drop_zone, fg_color="transparent")
        dz.pack(fill="x", padx=SP["md"], pady=SP["md"])

        ctk.CTkLabel(dz, text="📂", font=ctk.CTkFont(size=18), text_color=PAL["accent"]).pack(side="left")
        
        info = ctk.CTkFrame(dz, fg_color="transparent")
        info.pack(side="left", padx=(SP["sm"], 0), fill="x", expand=True)

        self.drop_title = ctk.CTkLabel(info, text="Drop .html or .pdf files here",
            font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
            text_color=PAL["t1"], anchor="w")
        self.drop_title.pack(anchor="w")

        self.drop_sub = ctk.CTkLabel(info, text="or click Browse",
            font=ctk.CTkFont(family=FONT, size=10), text_color=PAL["t4"], anchor="w")
        self.drop_sub.pack(anchor="w")

        btn_row = ctk.CTkFrame(scr, fg_color="transparent")
        btn_row.pack(fill="x", padx=self._pad, pady=(0, 0))

        ctk.CTkButton(btn_row, text="📁 Browse", height=28,
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            fg_color=PAL["accent"], hover_color=PAL["accent_h"],
            text_color="#FFFFFF", corner_radius=8,
            command=self._browse).pack(side="left", fill="x", expand=True)

        self.clear_btn = ctk.CTkButton(btn_row, text="✕", width=28, height=28,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=PAL["surface_2"], hover_color=PAL["danger"],
            text_color=PAL["t4"], corner_radius=8,
            command=lambda: self.app.set_files([]))
        self.clear_btn.pack_forget()

        # File list container
        self.file_list = ctk.CTkFrame(scr, fg_color="transparent")
        self.file_list.pack_forget()
        self.file_widgets: list[ctk.CTkFrame] = []

        # ── FORMAT ──
        self._section("OUTPUT FORMAT", "📦")

        self.format_var = ctk.StringVar(value="DOCX")
        
        self.fmt_seg = ctk.CTkSegmentedButton(scr, values=["📄 DOCX", "📊 PPTX"],
            font=ctk.CTkFont(family=FONT, size=12, weight="bold"),
            variable=self.format_var,
            fg_color=PAL["surface_3"], selected_color=PAL["accent"],
            selected_hover_color=PAL["accent_h"], text_color=PAL["t3"],
            unselected_color=PAL["surface_3"], unselected_hover_color=PAL["surface_2"],
            corner_radius=10, height=32,
            command=self._on_format_change)
        self.fmt_seg.set("📄 DOCX")
        self.fmt_seg.pack(fill="x", padx=self._pad, pady=(0, 0))

        self.settings_container = ctk.CTkFrame(scr, fg_color="transparent")
        self.settings_container.pack(fill="x")

        # ── DOCX SETTINGS ──
        self.docx_section = CollapsibleSection(self.settings_container, "PAGE SETTINGS", "⚙")
        self.docx_section.pack(fill="x", padx=self._pad)
        dc = self.docx_section.content

        self.page_size_var = ctk.StringVar(value="A4")
        self._opt_row(dc, "Size", self.page_size_var, list(converter.PAGE_SIZES.keys()))

        self.orient_var = ctk.StringVar(value="Portrait")
        orow = ctk.CTkFrame(dc, fg_color="transparent")
        orow.pack(fill="x", pady=SP["xs"])
        ctk.CTkLabel(orow, text="Orient.",
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=PAL["t3"], width=64, anchor="w").pack(side="left")
        seg = ctk.CTkSegmentedButton(orow, values=["Portrait", "Landscape"],
            font=ctk.CTkFont(family=FONT, size=11), variable=self.orient_var,
            fg_color=PAL["surface_3"], selected_color=PAL["accent"],
            selected_hover_color=PAL["accent_h"],
            unselected_color=PAL["surface_3"],
            unselected_hover_color=PAL["surface_2"],
            text_color=PAL["t3"], corner_radius=10)
        seg.pack(side="left", padx=(SP["sm"], 0), fill="x", expand=True)

        self.margin_var = ctk.StringVar(value="Normal (0.5 in)")
        self._opt_row(dc, "Margins", self.margin_var,
            ["Narrow (0.3 in)", "Normal (0.5 in)", "Wide (0.8 in)", "Extra Wide (1.0 in)"])

        # ── PPTX SETTINGS ──
        self.pptx_section = CollapsibleSection(self.settings_container, "SLIDE SETTINGS", "⚙")
        # Hidden initially

        pc = self.pptx_section.content
        self.slide_size_var = ctk.StringVar(value="Widescreen (16:9)")
        self._opt_row(pc, "Slide", self.slide_size_var,
            ["Widescreen (16:9)", "Standard (4:3)"])
        self.dpi_var = ctk.StringVar(value="High (200 DPI)")
        self._opt_row(pc, "Quality", self.dpi_var,
            list(converter.DPI_PRESETS.keys()))

        # ── OUTPUT ──
        self._section("OUTPUT", "💾")

        out_card = HoverCard(scr, fg_color=PAL["surface_1"], corner_radius=12,
            border_width=1, border_color=PAL["b2"], hover_border=PAL["b1"])
        out_card.pack(fill="x", padx=self._pad, pady=(0, 0))

        out_inner = ctk.CTkFrame(out_card, fg_color="transparent")
        out_inner.pack(fill="x", padx=SP["sm"], pady=SP["sm"])

        self.output_var = ctk.StringVar(value="Same as input file")

        orow = ctk.CTkFrame(out_inner, fg_color="transparent")
        orow.pack(fill="x")
        ctk.CTkEntry(orow, textvariable=self.output_var, height=36,
            font=ctk.CTkFont(family=FONT, size=11),
            fg_color=PAL["surface_3"], border_color=PAL["b2"],
            text_color=PAL["t4"], corner_radius=12,
            state="disabled").pack(side="left", fill="x", expand=True, padx=(0, SP["sm"]))
        ctk.CTkButton(orow, text="📁", width=38, height=36,
            font=ctk.CTkFont(size=16),
            fg_color=PAL["surface_3"], hover_color=PAL["accent"],
            text_color=PAL["t3"], corner_radius=12,
            command=self._browse_output).pack(side="right")

        # ── CONVERT ──
        spacer = ctk.CTkFrame(scr, fg_color="transparent", height=SP["md"])
        spacer.pack(fill="x")

        self.convert_btn = PulseButton(scr, text="⚡ Convert Now", height=38,
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
            fg_color=PAL["accent"], hover_color=PAL["accent_h"],
            text_color="#FFFFFF", corner_radius=12,
            command=self.app.start_conversion)
        self.convert_btn.pack(fill="x", padx=self._pad, pady=(0, SP["sm"]))

        # Progress area
        prog_card = ctk.CTkFrame(scr, fg_color="transparent", corner_radius=0)
        prog_card.pack(fill="x", padx=self._pad, pady=(0, 0))

        prog_top = ctk.CTkFrame(prog_card, fg_color="transparent")
        prog_top.pack(fill="x", pady=(0, SP["xs"]))

        self.status_dot = StatusDot(prog_top, color=PAL["accent"], size=8)
        self.status_dot.pack(side="left", pady=2)

        self.status_label = ctk.CTkLabel(prog_top, text="Ready to convert",
            font=ctk.CTkFont(family=FONT, size=11),
            text_color=PAL["t4"], anchor="w")
        self.status_label.pack(side="left", padx=(SP["sm"], 0))

        self.pct_label = ctk.CTkLabel(prog_top, text="",
            font=ctk.CTkFont(family=MONO, size=11, weight="bold"),
            text_color=PAL["accent"])
        self.pct_label.pack(side="right")

        self.progress = ctk.CTkProgressBar(prog_card,
            fg_color=PAL["surface_3"], progress_color=PAL["accent"],
            corner_radius=8, height=8)
        self.progress.pack(fill="x")
        self.progress.set(0)

    # ── Helpers ──

    def _section(self, text, icon):
        f = ctk.CTkFrame(self._scr, fg_color="transparent")
        f.pack(fill="x", padx=self._pad, pady=(SP["sm"], SP["xs"]))
        ctk.CTkLabel(f, text=f"{icon}  {text}",
            font=ctk.CTkFont(family=FONT, size=10, weight="bold"),
            text_color=PAL["t4"]).pack(side="left")
        ctk.CTkFrame(f, fg_color=PAL["b2"], height=1
            ).pack(side="left", fill="x", expand=True, padx=(SP["md"], 0), pady=1)

    def _opt_row(self, parent, label, var, vals, w=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=SP["xs"])
        ctk.CTkLabel(row, text=label,
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=PAL["t3"], width=64, anchor="w").pack(side="left")
        ctk.CTkOptionMenu(row, variable=var, values=vals,
            font=ctk.CTkFont(family=FONT, size=11),
            width=w or 160, height=32,
            fg_color=PAL["surface_3"], button_color=PAL["accent_dim"],
            button_hover_color=PAL["accent"],
            dropdown_fg_color=PAL["surface_1"],
            dropdown_hover_color=PAL["surface_2"],
            text_color=PAL["t2"], corner_radius=10
            ).pack(side="left", padx=(SP["sm"], 0), fill="x", expand=True)

    def _browse(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("HTML & PDF Files", "*.html *.htm *.pdf"), ("All", "*.*")])
        if paths:
            existing = set(self.app._files)
            new = [p for p in paths if p not in existing]
            self.app.set_files(self.app._files + new)

    def _browse_output(self):
        p = filedialog.askdirectory()
        if p:
            self.output_var.set(p)

    def _on_format_change(self, val=None):
        if "DOCX" in self.format_var.get():
            self.pptx_section.pack_forget()
            self.docx_section.pack(fill="x", padx=self._pad)
        else:
            self.docx_section.pack_forget()
            self.pptx_section.pack(fill="x", padx=self._pad)

    def update_file_list(self, files: list[str]):
        for w in self.file_widgets:
            w.destroy()
        self.file_widgets.clear()

        if not files:
            self.file_list.pack_forget()
            self.clear_btn.pack_forget()
            self.drop_title.configure(text="Drop files here")
            self.drop_sub.configure(text="or click Browse to select .html or .pdf files")
            self.drop_zone.configure(border_color=PAL["b_accent"])
            return

        self.drop_title.configure(text=f"{len(files)} file{'s' if len(files) != 1 else ''} ready")
        self.drop_sub.configure(text="Click Browse to add more")
        self.drop_zone.configure(border_color=PAL["accent"])
        self.clear_btn.pack(side="left")

        self.file_list.pack(fill="x", padx=self._pad, pady=(SP["sm"], 0),
                            after=self.drop_zone)

        for i, fp in enumerate(files):
            fname = os.path.basename(fp)
            ext = Path(fp).suffix.lower()

            card = HoverCard(self.file_list, fg_color=PAL["surface_1"],
                corner_radius=10, border_width=1, border_color=PAL["b2"],
                hover_border=PAL["b_accent"], height=28)
            card.pack(fill="x", pady=2)
            card.pack_propagate(False)

            # Extension badge
            ext_colors = {".html": ("#E44D26", "#FFF3EE"), ".htm": ("#E44D26", "#FFF3EE"), ".pdf": ("#D9383A", "#FCE8E8")}
            ec, ebg = ext_colors.get(ext, (PAL["accent"], PAL["accent_bg"]))

            badge = ctk.CTkFrame(card, fg_color=ebg if isinstance(ebg, str) else ebg,
                corner_radius=4, width=32, height=18)
            badge.pack(side="left", padx=(SP["sm"], 0))
            badge.pack_propagate(False)
            ctk.CTkLabel(badge, text=ext.replace(".", "").upper(),
                font=ctk.CTkFont(family=MONO, size=8, weight="bold"),
                text_color=ec).place(relx=0.5, rely=0.5, anchor="center")

            # Filename
            ctk.CTkLabel(card, text=fname,
                font=ctk.CTkFont(family=FONT, size=10, weight="bold"),
                text_color=PAL["t2"], anchor="w"
                ).pack(side="left", padx=(SP["sm"], 0), fill="x", expand=True)

            # File size
            try:
                sz = os.path.getsize(fp)
                s = f"{sz/1024:.0f} KB" if sz < 1048576 else f"{sz/1048576:.1f} MB"
                ctk.CTkLabel(card, text=s,
                    font=ctk.CTkFont(family=MONO, size=9),
                    text_color=PAL["t5"], width=54).pack(side="right", padx=(0, SP["xs"]))
            except OSError:
                pass

            idx = i
            ctk.CTkButton(card, text="✕", width=22, height=22,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color="transparent", hover_color=PAL["danger"],
                text_color=PAL["t5"], corner_radius=8,
                command=lambda j=idx: self.app.remove_file(j)
                ).pack(side="right", padx=(0, SP["sm"]))

            self.file_widgets.append(card)

    def get_margin_inches(self):
        return {"Narrow (0.3 in)": 0.3, "Normal (0.5 in)": 0.5,
                "Wide (0.8 in)": 0.8, "Extra Wide (1.0 in)": 1.0
                }.get(self.margin_var.get(), 0.5)

    def get_dpi(self):
        return converter.DPI_PRESETS.get(self.dpi_var.get(), 200)

    def get_output_dir(self):
        v = self.output_var.get()
        return None if v in ("Same as input file", "") else v


class CollapsibleSection(ctk.CTkFrame):
    def __init__(self, master, title, icon="", expanded=True, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self._open = expanded

        self._hdr = ctk.CTkFrame(self, fg_color="transparent", cursor="hand2")
        self._hdr.pack(fill="x", pady=(SP["sm"], SP["xs"]))

        self._arrow = ctk.CTkLabel(self._hdr, text="▾" if expanded else "▸",
            font=ctk.CTkFont(size=10, weight="bold"), text_color=PAL["t4"], width=16)
        self._arrow.pack(side="left")

        ctk.CTkLabel(self._hdr, text=f"{icon}  {title}" if icon else title,
            font=ctk.CTkFont(family=FONT, size=10, weight="bold"),
            text_color=PAL["t4"]).pack(side="left", padx=(SP["xs"], 0))

        ctk.CTkFrame(self._hdr, fg_color=PAL["b2"], height=1
            ).pack(side="left", fill="x", expand=True, padx=(SP["md"], 0), pady=1)

        self._body = ctk.CTkFrame(self, fg_color=PAL["surface_1"],
            corner_radius=12, border_width=1, border_color=PAL["b2"])
        self._body_inner = ctk.CTkFrame(self._body, fg_color="transparent")
        self._body_inner.pack(fill="x", padx=SP["md"], pady=SP["sm"])

        if expanded:
            self._body.pack(fill="x")

        for w in [self._hdr, self._arrow]:
            w.bind("<Button-1>", self._toggle)

    @property
    def content(self):
        return self._body_inner

    def _toggle(self, e=None):
        self._open = not self._open
        if self._open:
            self._body.pack(fill="x")
            self._arrow.configure(text="▾")
        else:
            self._body.pack_forget()
            self._arrow.configure(text="▸")


# ═══════════════════════════════════════════════════════════
#  PREVIEW PANEL
# ═══════════════════════════════════════════════════════════


class PreviewPanel(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=PAL["root"], corner_radius=0, **kw)
        self._pages = []
        self._idx = 0
        self._tk_img = None
        self._zoom = 1.0

        # ── Toolbar ──
        bar = ctk.CTkFrame(self, fg_color=PAL["surface_0"], height=52, corner_radius=0)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        bar_in = ctk.CTkFrame(bar, fg_color="transparent")
        bar_in.pack(fill="x", padx=SP["xl"])

        ctk.CTkLabel(bar_in, text="🖼  Preview",
            font=ctk.CTkFont(family=FONT, size=14, weight="bold"),
            text_color=PAL["t1"]).pack(side="left", pady=SP["md"])

        # Page nav (center area)
        nav = ctk.CTkFrame(bar_in, fg_color="transparent")
        nav.pack(side="left", padx=(SP["3xl"], 0))

        self.prev_btn = ctk.CTkButton(nav, text="‹", width=32, height=32,
            font=ctk.CTkFont(size=20, weight="bold"),
            fg_color=PAL["surface_3"], hover_color=PAL["accent_dim"],
            text_color=PAL["t3"], corner_radius=10, state="disabled",
            command=self._prev)
        self.prev_btn.pack(side="left")

        self.page_pill = ctk.CTkFrame(nav, fg_color=PAL["surface_3"],
            corner_radius=10, border_width=1, border_color=PAL["b2"])
        self.page_pill.pack(side="left", padx=SP["xs"])

        self.page_lbl = ctk.CTkLabel(self.page_pill, text=" — ",
            font=ctk.CTkFont(family=MONO, size=11, weight="bold"),
            text_color=PAL["t3"])
        self.page_lbl.pack(padx=SP["lg"], pady=SP["xs"])

        self.next_btn = ctk.CTkButton(nav, text="›", width=32, height=32,
            font=ctk.CTkFont(size=20, weight="bold"),
            fg_color=PAL["surface_3"], hover_color=PAL["accent_dim"],
            text_color=PAL["t3"], corner_radius=10, state="disabled",
            command=self._next)
        self.next_btn.pack(side="left")

        # Zoom (right)
        zoom = ctk.CTkFrame(bar_in, fg_color="transparent")
        zoom.pack(side="right", pady=SP["sm"])

        zoom_pill = ctk.CTkFrame(zoom, fg_color=PAL["surface_3"],
            corner_radius=10, border_width=1, border_color=PAL["b2"])
        zoom_pill.pack(side="right")

        for txt, cmd in [("−", self._zout), (None, None), ("+", self._zin)]:
            if txt is None:
                self.zoom_lbl = ctk.CTkLabel(zoom_pill, text="100%",
                    font=ctk.CTkFont(family=MONO, size=10, weight="bold"),
                    text_color=PAL["t3"], width=48)
                self.zoom_lbl.pack(side="left")
                continue
            ctk.CTkButton(zoom_pill, text=txt, width=30, height=30,
                font=ctk.CTkFont(size=16, weight="bold"),
                fg_color="transparent", hover_color=PAL["surface_2"],
                text_color=PAL["t3"], corner_radius=8,
                command=cmd).pack(side="left", padx=2)

        ctk.CTkButton(zoom, text="⊡ Fit", width=52, height=30,
            font=ctk.CTkFont(family=FONT, size=10, weight="bold"),
            fg_color=PAL["surface_3"], hover_color=PAL["surface_2"],
            text_color=PAL["t3"], corner_radius=10,
            border_width=1, border_color=PAL["b2"],
            command=self._zfit).pack(side="right", padx=(0, SP["sm"]))

        # Divider
        ctk.CTkFrame(self, fg_color=PAL["b2"], height=1).pack(fill="x")

        # ── Canvas ──
        self.cvs_frame = ctk.CTkFrame(self, fg_color=PAL["inset"], corner_radius=0)
        self.cvs_frame.pack(fill="both", expand=True)

        self.canvas = Canvas(self.cvs_frame, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True, padx=SP["4xl"], pady=SP["3xl"])
        self._update_bg()

        # ── Empty state ──
        self._empty = ctk.CTkFrame(self.cvs_frame, fg_color="transparent")
        self._empty.place(relx=0.5, rely=0.45, anchor="center")

        # Big icon circle
        ico_bg = ctk.CTkFrame(self._empty, fg_color=PAL["surface_1"],
            corner_radius=40, width=80, height=80,
            border_width=2, border_color=PAL["b2"])
        ico_bg.pack()
        ico_bg.pack_propagate(False)
        ctk.CTkLabel(ico_bg, text="📋",
            font=ctk.CTkFont(size=36)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(self._empty, text="No Preview Yet",
            font=ctk.CTkFont(family=FONT, size=22, weight="bold"),
            text_color=PAL["t3"]).pack(pady=(SP["xl"], SP["xs"]))

        ctk.CTkLabel(self._empty, text="Select HTML or PDF files and click Convert\nto see a live preview here",
            font=ctk.CTkFont(family=FONT, size=13),
            text_color=PAL["t5"], justify="center").pack()

        # Keyboard hint
        hint = ctk.CTkFrame(self._empty, fg_color=PAL["surface_1"],
            corner_radius=12, border_width=1, border_color=PAL["b2"])
        hint.pack(pady=(SP["xl"], 0))
        ctk.CTkLabel(hint, text="💡  Tip: Use scroll wheel to navigate pages",
            font=ctk.CTkFont(family=FONT, size=10),
            text_color=PAL["t4"]).pack(padx=SP["lg"], pady=SP["sm"])

        self.canvas.bind("<Configure>", lambda e: self._render())
        self.canvas.bind("<MouseWheel>", self._scroll)

    def _update_bg(self):
        m = ctk.get_appearance_mode()
        self.canvas.configure(bg="#E8E2D8" if m == "Light" else "#030406")

    def set_pages(self, pages):
        self._pages = pages
        self._idx = 0
        self._zoom = 1.0
        self.zoom_lbl.configure(text="100%")
        self._empty.place_forget()
        self._update_nav()
        self._render()

    def clear(self):
        self._pages = []
        self._idx = 0
        self.canvas.delete("all")
        self._empty.place(relx=0.5, rely=0.45, anchor="center")
        self.page_lbl.configure(text=" — ")
        self.prev_btn.configure(state="disabled")
        self.next_btn.configure(state="disabled")

    def _render(self):
        if not self._pages:
            return
        self.canvas.delete("all")
        self._update_bg()

        img = self._pages[self._idx]
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw < 20 or ch < 20:
            return

        pad = 24
        max_w = int((cw - pad * 2) * self._zoom)
        max_h = int((ch - pad * 2) * self._zoom)
        iw, ih = img.size
        r = min(max_w / iw, max_h / ih, self._zoom)
        nw, nh = max(int(iw * r), 1), max(int(ih * r), 1)

        resized = img.resize((nw, nh), Image.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(resized)

        cx, cy = cw // 2, ch // 2
        sx, sy = cx - nw // 2, cy - nh // 2
        ex, ey = cx + nw // 2, cy + nh // 2

        # Multi-layer shadow
        mode = ctk.get_appearance_mode()
        shd = (["#B8AFA4","#C4BCB2","#D0C9C0","#DCD6CE","#E4DED6"]
               if mode == "Light"
               else ["#010203","#020305","#030408","#04060A","#06080E"])
        for i, c in enumerate(shd):
            o = (len(shd) - i) * 3
            self.canvas.create_rectangle(sx+o, sy+o, ex+o, ey+o, fill=c, outline="", width=0)

        bc = "#C4BBB0" if mode == "Light" else "#1E1C19"
        self.canvas.create_rectangle(sx-1, sy-1, ex+1, ey+1, outline=bc, width=1)
        self.canvas.create_image(cx, cy, image=self._tk_img, anchor="center")

    def _update_nav(self):
        n = len(self._pages)
        if n == 0:
            self.page_lbl.configure(text=" — ")
            self.prev_btn.configure(state="disabled")
            self.next_btn.configure(state="disabled")
            return
        self.page_lbl.configure(text=f" {self._idx+1} / {n} ")
        self.prev_btn.configure(state="normal" if self._idx > 0 else "disabled")
        self.next_btn.configure(state="normal" if self._idx < n-1 else "disabled")

    def _prev(self):
        if self._idx > 0:
            self._idx -= 1; self._update_nav(); self._render()

    def _next(self):
        if self._idx < len(self._pages)-1:
            self._idx += 1; self._update_nav(); self._render()

    def _zin(self):
        if self._zoom < 3:
            self._zoom = min(self._zoom + 0.25, 3.0)
            self.zoom_lbl.configure(text=f"{int(self._zoom*100)}%"); self._render()

    def _zout(self):
        if self._zoom > 0.25:
            self._zoom = max(self._zoom - 0.25, 0.25)
            self.zoom_lbl.configure(text=f"{int(self._zoom*100)}%"); self._render()

    def _zfit(self):
        self._zoom = 1.0; self.zoom_lbl.configure(text="100%"); self._render()

    def _scroll(self, e):
        (self._prev if e.delta > 0 else self._next)()


# ═══════════════════════════════════════════════════════════
#  ACTIVITY LOG
# ═══════════════════════════════════════════════════════════


class ActivityLog(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=PAL["surface_0"], width=300, corner_radius=0, **kw)
        self.pack_propagate(False)

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=SP["lg"], pady=(SP["md"], SP["sm"]))

        ctk.CTkLabel(hdr, text="📋  Activity Log",
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            text_color=PAL["t4"]).pack(side="left")

        self._counter = ctk.CTkLabel(hdr, text="",
            font=ctk.CTkFont(family=MONO, size=9),
            text_color=PAL["t5"])
        self._counter.pack(side="left", padx=(SP["sm"], 0))

        ctk.CTkButton(hdr, text="Clear", width=48, height=22,
            font=ctk.CTkFont(family=FONT, size=10, weight="bold"),
            fg_color="transparent", hover_color=PAL["surface_3"],
            text_color=PAL["t4"], corner_radius=8,
            command=self.clear).pack(side="right")

        self._textbox = ctk.CTkTextbox(self,
            font=ctk.CTkFont(family=MONO, size=11),
            fg_color=PAL["surface_1"], text_color=PAL["t3"],
            border_width=1, border_color=PAL["b2"],
            corner_radius=12, wrap="word")
        self._textbox.pack(fill="both", expand=True, padx=SP["lg"], pady=(0, SP["lg"]))
        self._textbox.configure(state="disabled")
        self._count = 0

    def log(self, msg, level="info"):
        icons = {"info": "→", "success": "✓", "error": "✗", "warn": "⚠"}
        colors = {"success": PAL["accent"], "error": PAL["danger"], "warn": PAL["warn"]}
        p = icons.get(level, "→")
        self._textbox.configure(state="normal")
        self._textbox.insert("end", f"  {p}  {msg}\n")
        self._textbox.see("end")
        self._textbox.configure(state="disabled")
        self._count += 1
        self._counter.configure(text=f"({self._count})")

    def clear(self):
        self._textbox.configure(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.configure(state="disabled")
        self._count = 0
        self._counter.configure(text="")


# ═══════════════════════════════════════════════════════════
#  APP
# ═══════════════════════════════════════════════════════════


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HTML/PDF → Office Converter")
        self.geometry("1260x840")
        self.minsize(1000, 660)
        self.configure(fg_color=PAL["root"])

        self._files: list[str] = []
        self._converting = False
        self._temp_pdfs: list[str] = []

        self._build()
        self._check_pw()

    def _build(self):
        # ═══ TOP BAR ═══
        top = ctk.CTkFrame(self, fg_color=PAL["surface_0"], height=58, corner_radius=0)
        top.pack(fill="x")
        top.pack_propagate(False)

        top_in = ctk.CTkFrame(top, fg_color="transparent")
        top_in.pack(fill="x", padx=SP["2xl"])

        # Accent strip
        ctk.CTkFrame(self, fg_color=PAL["accent"], height=3, corner_radius=0).pack(fill="x")

        # Logo
        logo = ctk.CTkFrame(top_in, fg_color="transparent")
        logo.pack(side="left", pady=SP["md"])

        # Logo icon with background
        ico_frame = ctk.CTkFrame(logo, fg_color=PAL["accent_bg"],
            corner_radius=14, width=38, height=38,
            border_width=1, border_color=PAL["b_accent"])
        ico_frame.pack(side="left")
        ico_frame.pack_propagate(False)
        ctk.CTkLabel(ico_frame, text="◈",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=PAL["accent"]).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(logo, text="HTML/PDF → Office",
            font=ctk.CTkFont(family=FONT, size=20, weight="bold"),
            text_color=PAL["t1"]).pack(side="left", padx=(SP["md"], 0))

        ctk.CTkLabel(logo, text="Converter",
            font=ctk.CTkFont(family=FONT, size=20),
            text_color=PAL["t4"]).pack(side="left", padx=(SP["sm"], 0))

        # Version pill
        vpill = ctk.CTkFrame(logo, fg_color=PAL["accent_bg"], corner_radius=8,
            border_width=1, border_color=PAL["b_accent"])
        vpill.pack(side="left", padx=(SP["md"], 0))
        ctk.CTkLabel(vpill, text="v2.1",
            font=ctk.CTkFont(family=MONO, size=9, weight="bold"),
            text_color=PAL["accent"]).pack(padx=SP["sm"], pady=2)

        # Right
        rt = ctk.CTkFrame(top_in, fg_color="transparent")
        rt.pack(side="right", pady=SP["md"])

        ctk.CTkButton(rt, text="⭐  GitHub", width=90, height=34,
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            fg_color=PAL["surface_3"], hover_color=PAL["surface_2"],
            text_color=PAL["t3"], corner_radius=10,
            border_width=1, border_color=PAL["b2"],
            command=lambda: webbrowser.open(GITHUB_URL)).pack(side="left", padx=(0, SP["md"]))

        theme = ctk.CTkSegmentedButton(rt, values=["🌙 Dark", "☀ Light"],
            font=ctk.CTkFont(family=FONT, size=11, weight="bold"),
            fg_color=PAL["surface_3"], selected_color=PAL["accent"],
            selected_hover_color=PAL["accent_h"],
            unselected_color=PAL["surface_3"],
            unselected_hover_color=PAL["surface_2"],
            text_color=PAL["t3"], corner_radius=10,
            command=self._theme)
        theme.set("🌙 Dark")
        theme.pack(side="left")

        # ═══ BODY ═══
        body = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        body.pack(fill="both", expand=True)

        self.sidebar = Sidebar(body, app=self)
        self.sidebar.pack(side="left", fill="y")

        ctk.CTkFrame(body, fg_color=PAL["b2"], width=1).pack(side="left", fill="y")

        self.log = ActivityLog(body)
        self.log.pack(side="right", fill="y")

        ctk.CTkFrame(body, fg_color=PAL["b2"], width=1).pack(side="right", fill="y")

        self.preview = PreviewPanel(body)
        self.preview.pack(side="left", fill="both", expand=True)

        # ═══ BOTTOM ═══

        # Footer
        ft = ctk.CTkFrame(self, fg_color=PAL["surface_0"], height=32, corner_radius=0)
        ft.pack(fill="x", side="bottom")
        ft.pack_propagate(False)

        ft_in = ctk.CTkFrame(ft, fg_color="transparent")
        ft_in.pack(fill="x", padx=SP["2xl"])

        ctk.CTkLabel(ft_in, text="Built with ❤ by Hamza-op",
            font=ctk.CTkFont(family=FONT, size=10),
            text_color=PAL["t5"]).pack(side="left", pady=SP["xs"])

        link = ctk.CTkLabel(ft_in, text="github.com/Hamza-op",
            font=ctk.CTkFont(family=FONT, size=10, underline=True),
            text_color=PAL["accent_dim"], cursor="hand2")
        link.pack(side="right", pady=SP["xs"])
        link.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_URL))
        link.bind("<Enter>", lambda e: link.configure(text_color=PAL["accent"]))
        link.bind("<Leave>", lambda e: link.configure(text_color=PAL["accent_dim"]))

    def _theme(self, v):
        ctk.set_appearance_mode("light" if "Light" in v else "dark")
        self.after(100, self.preview._update_bg)
        self.after(150, self.preview._render)

    def set_files(self, files):
        self._files = files
        self.sidebar.update_file_list(self._files)
        self._log(f"Selected {len(files)} file(s)")

    def remove_file(self, i):
        if 0 <= i < len(self._files):
            self._files.pop(i)
            self.sidebar.update_file_list(self._files)

    def _log(self, msg, level="info"):
        self.after(0, lambda: self.log.log(msg, level))

    def _set_progress(self, v):
        self.after(0, lambda: self.sidebar.progress.set(v))
        self.after(0, lambda: self.sidebar.pct_label.configure(
            text=f"{int(v*100)}%" if v > 0 else ""))

    def _set_status(self, txt, color=None):
        self.after(0, lambda: self.sidebar.status_label.configure(text=txt))
        if color:
            self.after(0, lambda: self.sidebar.status_dot.set_color(color))

    def _set_converting(self, state):
        self._converting = state
        if state:
            self.after(0, lambda: (
                self.sidebar.convert_btn.configure(text="⏳  Converting…",
                    state="disabled", fg_color=PAL["surface_3"], text_color=PAL["t4"]),
                self.sidebar.convert_btn.set_loading(True),
                self.sidebar.status_dot.start_pulse() if hasattr(self.sidebar.status_dot, 'start_pulse') else None,
            ))
        else:
            self.after(0, lambda: (
                self.sidebar.convert_btn.configure(text="⚡  Convert Now",
                    state="normal", fg_color=PAL["accent"], text_color="#FFFFFF"),
                self.sidebar.convert_btn.set_loading(False),
                self.sidebar.status_dot.stop_pulse() if hasattr(self.sidebar.status_dot, 'stop_pulse') else None,
            ))

    def _check_pw(self):
        def check():
            self._log("Checking browser engine…")
            self._set_status("Checking browser…", PAL["warn"])
            name = converter.get_available_browser_name()
            if name:
                self._log(f"Browser ready: {name}.", "success")
                self._set_status(f"Ready ({name})", PAL["accent"])
            else:
                self._log("No browser found — trying to install Chromium…", "warn")
                ok = converter.install_playwright_browser(on_status=self._log)
                if ok:
                    self._log("Browser engine ready.", "success")
                    self._set_status("Ready to convert", PAL["accent"])
                else:
                    self._log("Install Chrome or Edge to use this app.", "error")
                    self._log("Or run: python -m playwright install chromium", "warn")
                    self._set_status("No browser found", PAL["danger"])
        threading.Thread(target=check, daemon=True).start()

    def start_conversion(self):
        if self._converting:
            return
        if not self._files:
            messagebox.showwarning("No Files", "Please select HTML or PDF files first.")
            return

        for p in self._temp_pdfs:
            try: os.unlink(p)
            except OSError: pass
        self._temp_pdfs.clear()

        self._set_converting(True)
        self._set_progress(0)
        self._set_status("Starting…", PAL["warn"])
        self.preview.clear()
        threading.Thread(target=self._convert, daemon=True).start()

    def _convert(self):
        fmt = self.sidebar.format_var.get()
        out = self.sidebar.get_output_dir()
        total = len(self._files)
        ok = 0
        last_pdf = last_shots = None

        try:
            for i, html in enumerate(self._files):
                base = os.path.basename(html)
                self._log(f"Processing: {base}")
                self._set_status(f"Converting {i+1}/{total}…", PAL["warn"])
                self._set_progress(i / total)

                try:
                    od = out or os.path.dirname(html)
                    stem = Path(html).stem
                    ext = Path(html).suffix.lower()
                    op = os.path.join(od, stem + (".docx" if "DOCX" in fmt else ".pptx"))

                    if ext == ".pdf":
                        # Convert from PDF
                        if "DOCX" in fmt:
                            converter.pdf_to_docx(html, op, on_status=self._log)
                        else:
                            converter.pdf_to_editable_pptx(html, op,
                                slide_size=self.sidebar.slide_size_var.get(),
                                on_status=self._log)
                        last_pdf = html
                    else:
                        # Convert from HTML
                        if "DOCX" in fmt:
                            _, pdf = converter.html_to_docx(html, op,
                                page_size=self.sidebar.page_size_var.get(),
                                orientation=self.sidebar.orient_var.get(),
                                margin_inches=self.sidebar.get_margin_inches(),
                                on_status=self._log, keep_pdf=True)
                            if pdf:
                                last_pdf = pdf; self._temp_pdfs.append(pdf)
                        else:
                            shots = converter.render_html_screenshots(html,
                                dpi=self.sidebar.get_dpi(), on_status=self._log)
                            if shots:
                                converter.html_to_editable_pptx(html, op,
                                    slide_size=self.sidebar.slide_size_var.get(),
                                    on_status=self._log)
                                last_shots = shots

                    self._log(f"Saved: {os.path.basename(op)}", "success")
                    ok += 1
                except Exception as e:
                    self._log(f"Error: {base} — {e}", "error")

                self._set_progress((i+1) / total)

            # Preview
            if "DOCX" in fmt and last_pdf:
                try:
                    pages = converter.pdf_to_images(last_pdf, dpi=150, on_status=self._log)
                    self.after(0, lambda p=pages: self.preview.set_pages(p))
                except Exception as e:
                    self._log(f"Preview: {e}", "warn")
            elif "PPTX" in fmt and last_shots:
                try:
                    from io import BytesIO
                    pages = [Image.open(BytesIO(s)) for s in last_shots]
                    self.after(0, lambda p=pages: self.preview.set_pages(p))
                except Exception as e:
                    self._log(f"Preview: {e}", "warn")

            if ok == total:
                self._log(f"All {total} file(s) converted!", "success")
                self._set_status(f"✓ Done — {total} file(s)", PAL["accent"])
                self.after(0, lambda: Toast.show(self, f"All {total} files converted!", "success"))
            elif ok > 0:
                self._log(f"{ok}/{total} converted.", "warn")
                self._set_status(f"⚠ {ok}/{total} done", PAL["warn"])
                self.after(0, lambda: Toast.show(self, f"{ok}/{total} converted", "warn"))
            else:
                self._log("Conversion failed.", "error")
                self._set_status("✗ Failed", PAL["danger"])
                self.after(0, lambda: Toast.show(self, "Conversion failed!", "error"))
        except Exception as e:
            self._log(f"Fatal: {e}", "error")
            self._set_status("✗ Fatal error", PAL["danger"])
        finally:
            self._set_converting(False)
            self._set_progress(1.0 if ok == total else 0)

    def destroy(self):
        for p in self._temp_pdfs:
            try: os.unlink(p)
            except OSError: pass
        super().destroy()


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()   # Required for PyInstaller on Windows
    App().mainloop()
