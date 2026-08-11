"""Einkauf Data Converter - GUI.

Arbeitsfluss in jedem Werkzeug-Tab:
    Quelle wählen -> Vorschau (Dry-Run-Plan) -> Ausführen (Live-Log, Fortschritt)

Reine Standardbibliothek (tkinter/ttk), keine Abhängigkeiten.
"""

import json
import queue
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import os
import subprocess

from . import APP_TITLE, __version__, cad_io, filetools, tabular

# Farbwelt
BG = "#eef1f6"
CARD = "#ffffff"
BORDER = "#d7dde8"
TEXT = "#1c2430"
SUBTLE = "#5b6676"
ACCENT = "#1f6feb"
ACCENT_DARK = "#1858c4"
HEADER_BG = "#16233c"
LOG_BG = "#101623"

FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")


@dataclass
class PlanItem:
    source: str          # Anzeige Quelle
    size: str            # Anzeige Größe
    target: str          # Anzeige Ergebnis/Ziel
    payload: object = None


# ---------------------------------------------------------------------------
# Lauf-Kontext (Worker-Thread -> GUI über Queue)
# ---------------------------------------------------------------------------

class RunContext:
    def __init__(self, out_queue, stop_event, log_path=None):
        self.queue = out_queue
        self.stop_event = stop_event
        self.log_path = log_path
        if log_path is not None:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.write_text(
                f"{APP_TITLE} {__version__}\nStart: {datetime.now():%Y-%m-%d %H:%M:%S}\n\n",
                encoding="utf-8",
            )

    def stopped(self):
        return self.stop_event.is_set()

    def log(self, message, level="info"):
        self.queue.put(("log", message, level))
        if self.log_path is not None:
            with self.log_path.open("a", encoding="utf-8") as handle:
                handle.write(f"{datetime.now():%H:%M:%S}  {message}\n")

    def item(self, index, status):
        self.queue.put(("item", index, status))

    def progress(self, done, total):
        self.queue.put(("progress", done, total))


# ---------------------------------------------------------------------------
# Bausteine
# ---------------------------------------------------------------------------

def make_card(parent, title=None):
    """Weiße Karte mit feinem Rand. Liefert (outer, content)."""
    outer = tk.Frame(parent, bg=BORDER)
    inner = tk.Frame(outer, bg=CARD)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    content = ttk.Frame(inner, style="Card.TFrame", padding=(14, 10, 14, 12))
    content.pack(fill="both", expand=True)
    if title:
        ttk.Label(content, text=title, style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
    return outer, content


class SourcePicker(ttk.Frame):
    """Quelle (Ordner oder mehrere Dateien) plus Filterleiste."""

    def __init__(self, parent, default_exts="", ext_label="Dateitypen"):
        super().__init__(parent, style="Card.TFrame")
        self.sources = []
        self.display_var = tk.StringVar(value="Noch keine Quelle gewählt")
        self.ext_var = tk.StringVar(value=default_exts)
        self.contains_var = tk.StringVar()
        self.max_mb_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)

        row1 = ttk.Frame(self, style="Card.TFrame")
        row1.pack(fill="x")
        entry = ttk.Entry(row1, textvariable=self.display_var, state="readonly")
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="Ordner…", command=self.pick_folder).pack(side="left", padx=(8, 4))
        ttk.Button(row1, text="Dateien…", command=self.pick_files).pack(side="left")

        row2 = ttk.Frame(self, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text=ext_label, style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row2, textvariable=self.ext_var, width=26).pack(side="left", padx=(6, 14))
        ttk.Label(row2, text="Name enthält", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row2, textvariable=self.contains_var, width=16).pack(side="left", padx=(6, 14))
        ttk.Label(row2, text="max. MB", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row2, textvariable=self.max_mb_var, width=7).pack(side="left", padx=(6, 14))
        ttk.Checkbutton(row2, text="Unterordner einbeziehen", variable=self.recursive_var,
                        style="Card.TCheckbutton").pack(side="left")

    def pick_folder(self):
        path = filedialog.askdirectory(title="Quellordner wählen")
        if path:
            self.sources = [Path(path)]
            self.display_var.set(path)

    def pick_files(self):
        paths = filedialog.askopenfilenames(title="Dateien wählen")
        if paths:
            self.sources = [Path(p) for p in paths]
            names = ", ".join(p.name for p in self.sources[:3])
            extra = f" … (+{len(self.sources) - 3})" if len(self.sources) > 3 else ""
            self.display_var.set(f"{len(self.sources)} Datei(en): {names}{extra}")

    def collect(self):
        if not self.sources:
            raise ValueError("Bitte zuerst eine Quelle wählen (Ordner oder Dateien).")
        max_mb = None
        raw = self.max_mb_var.get().strip().replace(",", ".")
        if raw:
            try:
                max_mb = float(raw)
            except ValueError:
                raise ValueError(f"„{raw}“ ist keine gültige MB-Angabe.")
        files = filetools.iter_files(
            self.sources,
            recursive=self.recursive_var.get(),
            exts=filetools.parse_ext_filter(self.ext_var.get()),
            contains=self.contains_var.get().strip(),
            max_mb=max_mb,
        )
        if not files:
            raise ValueError("Keine passenden Dateien gefunden – Quelle und Filter prüfen.")
        return files

    def base_dir(self):
        if not self.sources:
            return Path.cwd()
        first = self.sources[0]
        return first if first.is_dir() else first.parent

    def get_state(self):
        return {
            "sources": [str(p) for p in self.sources],
            "display": self.display_var.get(),
            "exts": self.ext_var.get(),
            "contains": self.contains_var.get(),
            "max_mb": self.max_mb_var.get(),
            "recursive": self.recursive_var.get(),
        }

    def set_state(self, state):
        self.sources = [Path(p) for p in state.get("sources", [])]
        self.display_var.set(state.get("display", "Noch keine Quelle gewählt"))
        self.ext_var.set(state.get("exts", ""))
        self.contains_var.set(state.get("contains", ""))
        self.max_mb_var.set(state.get("max_mb", ""))
        self.recursive_var.set(state.get("recursive", True))


class TargetPicker(ttk.Frame):
    """Zielordner: Unterordner neben der Quelle oder frei gewählt."""

    def __init__(self, parent, subfolder="_ergebnis"):
        super().__init__(parent, style="Card.TFrame")
        self.subfolder = subfolder
        self.mode_var = tk.StringVar(value="subfolder")
        self.custom_var = tk.StringVar()

        ttk.Label(self, text="Ziel", style="CardSub.TLabel").pack(side="left")
        ttk.Radiobutton(self, text=f"Unterordner „{subfolder}“", value="subfolder",
                        variable=self.mode_var, style="Card.TRadiobutton").pack(side="left", padx=(8, 12))
        ttk.Radiobutton(self, text="Eigener Ordner:", value="custom",
                        variable=self.mode_var, style="Card.TRadiobutton").pack(side="left")
        ttk.Entry(self, textvariable=self.custom_var, width=34).pack(side="left", fill="x", expand=True, padx=(6, 4))
        ttk.Button(self, text="…", width=3, command=self.pick).pack(side="left")

    def pick(self):
        path = filedialog.askdirectory(title="Zielordner wählen")
        if path:
            self.custom_var.set(path)
            self.mode_var.set("custom")

    def resolve(self, base_dir):
        if self.mode_var.get() == "custom":
            raw = self.custom_var.get().strip()
            if not raw:
                raise ValueError("Bitte einen Zielordner wählen oder den Unterordner-Modus nutzen.")
            return Path(raw)
        return Path(base_dir) / self.subfolder

    def get_state(self):
        return {"mode": self.mode_var.get(), "custom": self.custom_var.get()}

    def set_state(self, state):
        self.mode_var.set(state.get("mode", "subfolder"))
        self.custom_var.set(state.get("custom", ""))


# ---------------------------------------------------------------------------
# Tab-Basisklasse
# ---------------------------------------------------------------------------

class ToolTab(ttk.Frame):
    key = "tool"
    title = "Werkzeug"
    hint = ""
    default_exts = ""
    target_subfolder = "_ergebnis"
    has_target = True

    def __init__(self, app, parent):
        super().__init__(parent, padding=(12, 12, 12, 0))
        self.app = app
        self.plan = []
        self.run_cfg = {}
        self.log_dir = None
        self.tree_items = []

        src_outer, src_content = make_card(self, "1 · Quelle")
        src_outer.pack(fill="x")
        self.source = SourcePicker(src_content, default_exts=self.default_exts)
        self.source.pack(fill="x")

        opt_outer, opt_content = make_card(self, "2 · Einstellungen")
        opt_outer.pack(fill="x", pady=(10, 0))
        self.build_options(opt_content)
        if self.has_target:
            self.target = TargetPicker(opt_content, subfolder=self.target_subfolder)
            self.target.pack(fill="x", pady=(10, 0))
        else:
            self.target = None

        prev_outer, prev_content = make_card(self, "3 · Vorschau (Plan)")
        prev_outer.pack(fill="both", expand=True, pady=(10, 0))
        bar = ttk.Frame(prev_content, style="Card.TFrame")
        bar.pack(fill="x", pady=(0, 6))
        self.count_var = tk.StringVar(value="Noch kein Plan – „Vorschau“ klicken.")
        ttk.Label(bar, textvariable=self.count_var, style="CardSub.TLabel").pack(side="left")
        if self.hint:
            ttk.Label(bar, text=self.hint, style="CardSub.TLabel").pack(side="right")

        tree_frame = ttk.Frame(prev_content, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True)
        columns = ("quelle", "groesse", "ziel", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=9)
        for col, text, width, stretch in (
            ("quelle", "Quelle", 330, True),
            ("groesse", "Größe", 90, False),
            ("ziel", "→ Ergebnis", 330, True),
            ("status", "Status", 110, False),
        ):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=stretch, anchor="w")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.tag_configure("odd", background="#f6f8fc")
        self.tree.tag_configure("ok", foreground="#0f7b3d")
        self.tree.tag_configure("err", foreground="#c22a2a")
        self.tree.bind("<Double-1>", self._open_item_location)

    def _open_item_location(self, _event):
        selection = self.tree.selection()
        if not selection or selection[0] not in self.tree_items:
            return
        item = self.plan[self.tree_items.index(selection[0])]
        payload = item.payload
        path = payload if isinstance(payload, Path) else \
            payload[0] if isinstance(payload, tuple) and isinstance(payload[0], Path) else None
        if path and path.exists():
            subprocess.Popen(["explorer", "/select,", str(path)])

    # -- Vorschau ----------------------------------------------------------

    def refresh_plan(self):
        """Plan neu berechnen und im Treeview anzeigen. Wirft ValueError."""
        self.plan = self.make_plan()
        self.tree.delete(*self.tree.get_children())
        self.tree_items = []
        for idx, item in enumerate(self.plan):
            tag = "odd" if idx % 2 else "even"
            iid = self.tree.insert("", "end", values=(item.source, item.size, item.target, ""), tags=(tag,))
            self.tree_items.append(iid)
        self.count_var.set(f"{len(self.plan)} Aktion(en) geplant.")
        return self.plan

    def set_item_status(self, index, status):
        if 0 <= index < len(self.tree_items):
            iid = self.tree_items[index]
            self.tree.set(iid, "status", status)
            tag = "err" if status.startswith("FEHLER") else "ok"
            self.tree.item(iid, tags=(self.tree.item(iid, "tags")[0], tag))
            self.tree.see(iid)

    # -- von Unterklassen zu implementieren --------------------------------

    def build_options(self, parent):
        raise NotImplementedError

    def make_plan(self):
        raise NotImplementedError

    def execute(self, ctx):
        raise NotImplementedError

    def after_run(self):
        pass

    # -- Presets -----------------------------------------------------------

    def option_state(self):
        return {}

    def apply_option_state(self, state):
        pass

    def get_state(self):
        state = {"source": self.source.get_state(), "options": self.option_state()}
        if self.target:
            state["target"] = self.target.get_state()
        return state

    def set_state(self, state):
        self.source.set_state(state.get("source", {}))
        if self.target:
            self.target.set_state(state.get("target", {}))
        self.apply_option_state(state.get("options", {}))


def _labeled(parent, text, widget_factory, **pack_kwargs):
    ttk.Label(parent, text=text, style="CardSub.TLabel").pack(side="left", **pack_kwargs)
    widget = widget_factory(parent)
    widget.pack(side="left", padx=(6, 14))
    return widget


# ---------------------------------------------------------------------------
# Tab 1: Tabellen-Konverter
# ---------------------------------------------------------------------------

DECIMAL_LABELS = {"keep": "unverändert", "comma_to_dot": "Komma → Punkt", "dot_to_comma": "Punkt → Komma"}
CASE_LABELS = {"keep": "unverändert", "lower": "kleinbuchstaben", "upper": "GROSSBUCHSTABEN", "title": "Wortanfänge Groß"}


class TabellenTab(ToolTab):
    key = "tabellen"
    title = "Tabellen"
    hint = "CSV · TSV · TXT · XLSX · JSON · XML → CSV · TSV · XLSX · JSON · XML · HTML · MD"
    default_exts = "csv, tsv, txt, xlsx, json, xml"
    target_subfolder = "_konvertiert"

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        ttk.Label(row1, text="Zielformate", style="CardSub.TLabel").pack(side="left")
        self.format_vars = {}
        for ext in tabular.WRITE_EXTS:
            var = tk.BooleanVar(value=(ext == ".xlsx"))
            self.format_vars[ext] = var
            ttk.Checkbutton(row1, text=ext.lstrip(".").upper(), variable=var,
                            style="Card.TCheckbutton").pack(side="left", padx=(8, 0))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        self.delim_combo = _labeled(row2, "CSV-Trennzeichen",
                                    lambda p: ttk.Combobox(p, values=list(tabular.DELIMITER_CHOICES), width=5, state="readonly"))
        self.delim_combo.set(";")
        self.enc_combo = _labeled(row2, "Encoding",
                                  lambda p: ttk.Combobox(p, values=tabular.ENCODING_CHOICES, width=10, state="readonly"))
        self.enc_combo.set("utf-8-sig")
        self.decimal_combo = _labeled(row2, "Dezimalzeichen",
                                      lambda p: ttk.Combobox(p, values=list(DECIMAL_LABELS.values()), width=15, state="readonly"))
        self.decimal_combo.set(DECIMAL_LABELS["keep"])

        row3 = ttk.Frame(parent, style="Card.TFrame")
        row3.pack(fill="x", pady=(8, 0))
        self.trim_var = tk.BooleanVar(value=True)
        self.drop_empty_var = tk.BooleanVar(value=True)
        self.dedupe_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="Werte trimmen", variable=self.trim_var, style="Card.TCheckbutton").pack(side="left")
        ttk.Checkbutton(row3, text="Leerzeilen entfernen", variable=self.drop_empty_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(12, 0))
        ttk.Checkbutton(row3, text="Duplikate entfernen", variable=self.dedupe_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(12, 0))
        ttk.Label(row3, text="Spalten (leer = alle, „Alt>Neu“ benennt um)", style="CardSub.TLabel").pack(side="left", padx=(18, 6))
        self.columns_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.columns_var).pack(side="left", fill="x", expand=True)

        row4 = ttk.Frame(parent, style="Card.TFrame")
        row4.pack(fill="x", pady=(8, 0))
        self.merge_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row4, text="Alle Dateien zu einer Tabelle zusammenführen (Spalte „Quelle“ wird ergänzt)",
                        variable=self.merge_var, style="Card.TCheckbutton").pack(side="left")
        ttk.Label(row4, text="Name", style="CardSub.TLabel").pack(side="left", padx=(14, 6))
        self.merge_name_var = tk.StringVar(value="zusammengefuehrt")
        ttk.Entry(row4, textvariable=self.merge_name_var, width=24).pack(side="left")

    def _selected_formats(self):
        return [ext for ext, var in self.format_vars.items() if var.get()]

    def make_plan(self):
        files = self.source.collect()
        formats = self._selected_formats()
        if not formats:
            raise ValueError("Bitte mindestens ein Zielformat anhaken.")
        target_dir = self.target.resolve(self.source.base_dir())
        decimal = next(k for k, v in DECIMAL_LABELS.items() if v == self.decimal_combo.get())
        self.run_cfg = {
            "files": files,
            "formats": formats,
            "target_dir": target_dir,
            "delimiter": tabular.DELIMITER_CHOICES[self.delim_combo.get()],
            "encoding": self.enc_combo.get(),
            "options": tabular.TransformOptions(
                trim=self.trim_var.get(),
                drop_empty_rows=self.drop_empty_var.get(),
                dedupe=self.dedupe_var.get(),
                decimal=decimal,
                column_spec=self.columns_var.get().strip(),
            ),
            "merge": self.merge_var.get(),
            "merge_name": filetools.safe_filename(self.merge_name_var.get().strip() or "zusammengefuehrt"),
        }
        self.log_dir = target_dir
        plan = []
        if self.run_cfg["merge"]:
            source_label = f"{len(files)} Datei(en) zusammengeführt"
            for ext in formats:
                plan.append(PlanItem(source_label, "", str(target_dir / (self.run_cfg["merge_name"] + ext))))
        else:
            for path in files:
                size = filetools.format_size(path.stat().st_size)
                for ext in formats:
                    plan.append(PlanItem(path.name, size, str(target_dir / (path.stem + ext)), payload=(path, ext)))
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        done = 0
        total = len(self.plan)

        def load(path):
            table = tabular.read_table(path)
            transformed, warnings = tabular.apply_transform(table, cfg["options"])
            for warning in warnings:
                ctx.log(f"{path.name}: {warning}", "info")
            return transformed

        if cfg["merge"]:
            tables = []
            for path in cfg["files"]:
                if ctx.stopped():
                    return
                try:
                    tables.append((path.name, load(path)))
                except Exception as exc:
                    ctx.log(f"FEHLER beim Lesen: {path.name}: {exc}", "err")
            if not tables:
                raise RuntimeError("Keine Datei konnte gelesen werden.")
            merged = tabular.merge_tables(tables)
            ctx.log(f"Zusammengeführt: {len(merged.rows)} Zeilen, {len(merged.headers)} Spalten.")
            for idx, ext in enumerate(cfg["formats"]):
                if ctx.stopped():
                    return
                target = filetools.unique_path(cfg["target_dir"] / (cfg["merge_name"] + ext))
                try:
                    tabular.write_table(merged, target, cfg["delimiter"], cfg["encoding"])
                    ctx.item(idx, "OK")
                    ctx.log(f"OK: {target.name}", "ok")
                except Exception as exc:
                    ctx.item(idx, "FEHLER")
                    ctx.log(f"FEHLER: {target.name}: {exc}", "err")
                done += 1
                ctx.progress(done, total)
            return

        cache = {}
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path, ext = item.payload
            try:
                if path not in cache:
                    cache[path] = load(path)
                target = filetools.unique_path(cfg["target_dir"] / (path.stem + ext))
                tabular.write_table(cache[path], target, cfg["delimiter"], cfg["encoding"])
                ctx.item(idx, "OK")
                ctx.log(f"OK: {path.name} → {target.name}", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name} → {ext}: {exc}", "err")
            done += 1
            ctx.progress(done, total)

    def option_state(self):
        return {
            "formats": {ext: var.get() for ext, var in self.format_vars.items()},
            "delimiter": self.delim_combo.get(),
            "encoding": self.enc_combo.get(),
            "decimal": self.decimal_combo.get(),
            "trim": self.trim_var.get(),
            "drop_empty": self.drop_empty_var.get(),
            "dedupe": self.dedupe_var.get(),
            "columns": self.columns_var.get(),
            "merge": self.merge_var.get(),
            "merge_name": self.merge_name_var.get(),
        }

    def apply_option_state(self, state):
        for ext, value in state.get("formats", {}).items():
            if ext in self.format_vars:
                self.format_vars[ext].set(value)
        self.delim_combo.set(state.get("delimiter", ";"))
        self.enc_combo.set(state.get("encoding", "utf-8-sig"))
        self.decimal_combo.set(state.get("decimal", DECIMAL_LABELS["keep"]))
        self.trim_var.set(state.get("trim", True))
        self.drop_empty_var.set(state.get("drop_empty", True))
        self.dedupe_var.set(state.get("dedupe", False))
        self.columns_var.set(state.get("columns", ""))
        self.merge_var.set(state.get("merge", False))
        self.merge_name_var.set(state.get("merge_name", "zusammengefuehrt"))


# ---------------------------------------------------------------------------
# Tab 2: CAD
# ---------------------------------------------------------------------------

CAD_QUALITY = ["grob", "mittel", "fein"]


class CadTab(ToolTab):
    key = "cad"
    title = "CAD"
    hint = "STL · OBJ · PLY · 3MF · STEP → STL · OBJ · PLY · 3MF · GLB · HTML-3D"
    default_exts = "stl, obj, ply, 3mf, step, stp"
    target_subfolder = "_cad"

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        self.mode_var = tk.StringVar(value="konvertieren")
        ttk.Radiobutton(row1, text="Konvertieren (Mesh & STEP-Tessellierung)", value="konvertieren",
                        variable=self.mode_var, style="Card.TRadiobutton",
                        command=self._mode_changed).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(row1, text="STEP/IGES-Prüfbericht", value="bericht",
                        variable=self.mode_var, style="Card.TRadiobutton",
                        command=self._mode_changed).pack(side="left")

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Zielformate", style="CardSub.TLabel").pack(side="left")
        self.format_vars = {}
        labels = {".stl": "STL", ".obj": "OBJ", ".ply": "PLY", ".3mf": "3MF",
                  ".glb": "GLB", ".html": "HTML-3D-Ansicht"}
        for ext in cad_io.MESH_WRITE_EXTS:
            var = tk.BooleanVar(value=(ext in (".stl", ".html")))
            self.format_vars[ext] = var
            ttk.Checkbutton(row2, text=labels[ext], variable=var,
                            style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        self.quality_combo = _labeled(row2, "STEP-Qualität",
                                      lambda p: ttk.Combobox(p, values=CAD_QUALITY, width=8, state="readonly"),
                                      padx=(18, 0))
        self.quality_combo.set("mittel")

        row3 = ttk.Frame(parent, style="Card.TFrame")
        row3.pack(fill="x", pady=(8, 0))
        ttk.Label(row3, text="Bericht als", style="CardSub.TLabel").pack(side="left")
        self.report_csv_var = tk.BooleanVar(value=True)
        self.report_html_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row3, text="CSV", variable=self.report_csv_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Checkbutton(row3, text="HTML", variable=self.report_html_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Label(row3, text="Name", style="CardSub.TLabel").pack(side="left", padx=(14, 6))
        self.report_name_var = tk.StringVar(value="cad_bericht")
        ttk.Entry(row3, textvariable=self.report_name_var, width=20).pack(side="left")
        ttk.Label(row3, text="Hinweis: SLDPRT/IPT enthalten proprietäre Kernel-Daten – "
                            "beim Lieferanten STEP oder STL anfordern.",
                  style="CardSub.TLabel").pack(side="left", padx=(18, 0))

    def _mode_changed(self):
        self.source.ext_var.set("stl, obj, ply, 3mf, step, stp" if self.mode_var.get() == "konvertieren"
                                else "step, stp, iges, igs")

    def make_plan(self):
        mode = self.mode_var.get()
        files = self.source.collect()
        target_dir = self.target.resolve(self.source.base_dir())
        self.log_dir = target_dir
        if mode == "konvertieren":
            formats = [ext for ext, var in self.format_vars.items() if var.get()]
            if not formats:
                raise ValueError("Bitte mindestens ein Zielformat anhaken.")
            self.run_cfg = {"mode": mode, "files": files, "formats": formats,
                            "target_dir": target_dir, "quality": self.quality_combo.get()}
            plan = []
            for path in files:
                size = filetools.format_size(path.stat().st_size)
                for ext in formats:
                    plan.append(PlanItem(path.name, size, str(target_dir / (path.stem + ext)),
                                         payload=(path, ext)))
            return plan
        formats = [ext for ext, var in ((".csv", self.report_csv_var), (".html", self.report_html_var)) if var.get()]
        if not formats:
            raise ValueError("Bitte mindestens ein Berichtsformat anhaken (CSV oder HTML).")
        name = filetools.safe_filename(self.report_name_var.get().strip() or "cad_bericht")
        self.run_cfg = {"mode": mode, "files": files, "formats": formats,
                        "target_dir": target_dir, "name": name}
        return [PlanItem(path.name, filetools.format_size(path.stat().st_size),
                         "wird geprüft", payload=path) for path in files]

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        if cfg["mode"] == "konvertieren":
            cache = {}
            for idx, item in enumerate(self.plan):
                if ctx.stopped():
                    ctx.log("Gestoppt.")
                    return
                path, ext = item.payload
                try:
                    if path not in cache:
                        cache[path] = cad_io.read_mesh(path, quality=cfg["quality"], log=ctx.log)
                        ctx.log(f"GELESEN: {path.name} ({cache[path].stats()})")
                    target = filetools.unique_path(cfg["target_dir"] / (path.stem + ext))
                    cad_io.write_mesh(cache[path], target)
                    ctx.item(idx, "OK")
                    ctx.log(f"OK: {path.name} → {target.name}", "ok")
                except Exception as exc:
                    ctx.item(idx, "FEHLER")
                    ctx.log(f"FEHLER: {path.name} → {ext}: {exc}", "err")
                ctx.progress(idx + 1, len(self.plan))
            return

        rows = []
        headers = ["Datei"]
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path = item.payload
            try:
                info = cad_io.brep_info(path)
                info = {"Datei": path.name, **info}
                for key in info:
                    if key not in headers:
                        headers.append(key)
                rows.append(info)
                ctx.item(idx, "OK")
                ctx.log(f"GEPRÜFT: {path.name} ({info.get('Format', '?')}, "
                        f"{info.get('Entitäten', '?')} Entitäten)", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))
        if not rows:
            raise RuntimeError("Keine Datei konnte geprüft werden.")
        table = tabular.Table(headers, [[row.get(h, "") for h in headers] for row in rows])
        for ext in cfg["formats"]:
            target = filetools.unique_path(cfg["target_dir"] / (cfg["name"] + ext))
            tabular.write_table(table, target)
            ctx.log(f"BERICHT: {target}", "ok")

    def option_state(self):
        return {"mode": self.mode_var.get(),
                "formats": {ext: var.get() for ext, var in self.format_vars.items()},
                "quality": self.quality_combo.get(),
                "report_csv": self.report_csv_var.get(), "report_html": self.report_html_var.get(),
                "report_name": self.report_name_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "konvertieren"))
        for ext, value in state.get("formats", {}).items():
            if ext in self.format_vars:
                self.format_vars[ext].set(value)
        self.quality_combo.set(state.get("quality", "mittel"))
        self.report_csv_var.set(state.get("report_csv", True))
        self.report_html_var.set(state.get("report_html", True))
        self.report_name_var.set(state.get("report_name", "cad_bericht"))


# ---------------------------------------------------------------------------
# Tab 3: Umbenennen
# ---------------------------------------------------------------------------

class RenameTab(ToolTab):
    key = "umbenennen"
    title = "Umbenennen"
    hint = "Arbeitet direkt an den Quelldateien – mit Journal und Rückgängig."
    default_exts = ""
    has_target = False

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        self.search_var = tk.StringVar()
        self.replace_var = tk.StringVar()
        self.regex_var = tk.BooleanVar(value=False)
        ttk.Label(row1, text="Suchen", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row1, textvariable=self.search_var, width=20).pack(side="left", padx=(6, 14))
        ttk.Label(row1, text="Ersetzen", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row1, textvariable=self.replace_var, width=20).pack(side="left", padx=(6, 14))
        ttk.Checkbutton(row1, text="Regulärer Ausdruck", variable=self.regex_var,
                        style="Card.TCheckbutton").pack(side="left")

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        self.prefix_var = tk.StringVar()
        self.suffix_var = tk.StringVar()
        ttk.Label(row2, text="Präfix", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row2, textvariable=self.prefix_var, width=14).pack(side="left", padx=(6, 14))
        ttk.Label(row2, text="Suffix", style="CardSub.TLabel").pack(side="left")
        ttk.Entry(row2, textvariable=self.suffix_var, width=14).pack(side="left", padx=(6, 14))
        self.case_combo = _labeled(row2, "Schreibweise",
                                   lambda p: ttk.Combobox(p, values=list(CASE_LABELS.values()), width=17, state="readonly"))
        self.case_combo.set(CASE_LABELS["keep"])

        row3 = ttk.Frame(parent, style="Card.TFrame")
        row3.pack(fill="x", pady=(8, 0))
        self.number_var = tk.BooleanVar(value=False)
        self.number_start_var = tk.StringVar(value="1")
        self.number_digits_var = tk.StringVar(value="3")
        self.datestamp_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row3, text="Nummerieren", variable=self.number_var, style="Card.TCheckbutton").pack(side="left")
        ttk.Label(row3, text="ab", style="CardSub.TLabel").pack(side="left", padx=(8, 4))
        ttk.Spinbox(row3, from_=0, to=99999, textvariable=self.number_start_var, width=6).pack(side="left")
        ttk.Label(row3, text="Stellen", style="CardSub.TLabel").pack(side="left", padx=(8, 4))
        ttk.Spinbox(row3, from_=1, to=8, textvariable=self.number_digits_var, width=4).pack(side="left")
        ttk.Checkbutton(row3, text="Änderungsdatum voranstellen (JJJJ-MM-TT_)", variable=self.datestamp_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(16, 0))
        self.undo_button = ttk.Button(row3, text="Rückgängig (letzter Lauf)", command=self.undo, state="disabled")
        self.undo_button.pack(side="right")
        self.last_journal = []

    def _rename_options(self):
        case = next(k for k, v in CASE_LABELS.items() if v == self.case_combo.get())
        try:
            start = int(self.number_start_var.get())
            digits = int(self.number_digits_var.get())
        except ValueError:
            raise ValueError("Nummerierung: Start und Stellen müssen Zahlen sein.")
        return filetools.RenameOptions(
            search=self.search_var.get(),
            replace=self.replace_var.get(),
            use_regex=self.regex_var.get(),
            prefix=self.prefix_var.get(),
            suffix=self.suffix_var.get(),
            case=case,
            numbering=self.number_var.get(),
            number_start=start,
            number_digits=digits,
            date_stamp=self.datestamp_var.get(),
        )

    def make_plan(self):
        files = self.source.collect()
        rename_plan = filetools.plan_rename(files, self._rename_options())
        if not rename_plan:
            raise ValueError("Die Regeln ändern keinen einzigen Dateinamen – bitte anpassen.")
        self.run_cfg = {"plan": rename_plan}
        self.log_dir = self.source.base_dir()
        return [PlanItem(src.name, filetools.format_size(src.stat().st_size), new, payload=(src, new))
                for src, new in rename_plan]

    def execute(self, ctx):
        plan = self.run_cfg["plan"]
        journal = []
        for idx, (source, new_name) in enumerate(plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                break
            try:
                target = filetools.unique_path(source.with_name(new_name))
                source.rename(target)
                journal.append({"von": str(source), "nach": str(target)})
                ctx.item(idx, "OK")
                ctx.log(f"UMBENANNT: {source.name} → {target.name}", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {source.name}: {exc}", "err")
            ctx.progress(idx + 1, len(plan))
        self.last_journal = journal
        if journal:
            journal_path = filetools.unique_path(
                self.log_dir / f"_umbenennen_journal_{datetime.now():%Y%m%d_%H%M%S}.json")
            journal_path.write_text(json.dumps(journal, ensure_ascii=False, indent=2), encoding="utf-8")
            ctx.log(f"Journal gespeichert: {journal_path.name}")

    def after_run(self):
        self.undo_button.config(state="normal" if self.last_journal else "disabled")

    def undo(self):
        if not self.last_journal:
            return
        if not messagebox.askyesno(APP_TITLE, f"{len(self.last_journal)} Umbenennung(en) rückgängig machen?"):
            return
        restored = filetools.undo_rename(self.last_journal, log=self.app.log_line)
        self.app.log_line(f"Rückgängig abgeschlossen: {restored} Datei(en).", "ok")
        self.last_journal = []
        self.undo_button.config(state="disabled")

    def option_state(self):
        return {
            "search": self.search_var.get(), "replace": self.replace_var.get(),
            "regex": self.regex_var.get(), "prefix": self.prefix_var.get(),
            "suffix": self.suffix_var.get(), "case": self.case_combo.get(),
            "numbering": self.number_var.get(), "start": self.number_start_var.get(),
            "digits": self.number_digits_var.get(), "datestamp": self.datestamp_var.get(),
        }

    def apply_option_state(self, state):
        self.search_var.set(state.get("search", ""))
        self.replace_var.set(state.get("replace", ""))
        self.regex_var.set(state.get("regex", False))
        self.prefix_var.set(state.get("prefix", ""))
        self.suffix_var.set(state.get("suffix", ""))
        self.case_combo.set(state.get("case", CASE_LABELS["keep"]))
        self.number_var.set(state.get("numbering", False))
        self.number_start_var.set(state.get("start", "1"))
        self.number_digits_var.set(state.get("digits", "3"))
        self.datestamp_var.set(state.get("datestamp", False))


# ---------------------------------------------------------------------------
# Tab 3: Packen / Entpacken
# ---------------------------------------------------------------------------

PACK_MODES = {
    "einzeln": "Jede Datei einzeln zippen",
    "zip": "Alles in ein ZIP-Paket",
    "targz": "Alles in ein TAR.GZ-Paket",
    "entpacken": "Archive entpacken",
}


class PackTab(ToolTab):
    key = "packen"
    title = "Packen"
    hint = "ZIP/TAR erstellen oder ZIP · TAR · GZ · BZ2 · XZ entpacken"
    default_exts = ""
    target_subfolder = "_archiv"

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        self.mode_var = tk.StringVar(value="zip")
        for value, label in PACK_MODES.items():
            ttk.Radiobutton(row1, text=label, value=value, variable=self.mode_var,
                            style="Card.TRadiobutton", command=self._mode_changed).pack(side="left", padx=(0, 14))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Paketname", style="CardSub.TLabel").pack(side="left")
        self.bundle_var = tk.StringVar(value="Lieferantenpaket")
        ttk.Entry(row2, textvariable=self.bundle_var, width=28).pack(side="left", padx=(6, 16))
        self.flatten_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="Ordnerstruktur glätten (alles auf eine Ebene)",
                        variable=self.flatten_var, style="Card.TCheckbutton").pack(side="left")

    def _mode_changed(self):
        if self.mode_var.get() == "entpacken" and not self.source.ext_var.get().strip():
            self.source.ext_var.set("zip, tar, tgz, gz, bz2, xz")

    def make_plan(self):
        mode = self.mode_var.get()
        if mode == "entpacken" and not self.source.ext_var.get().strip():
            self.source.ext_var.set("zip, tar, tgz, gz, bz2, xz")
        files = self.source.collect()
        base = self.source.base_dir()
        target_dir = self.target.resolve(base)
        bundle_name = filetools.safe_filename(self.bundle_var.get().strip() or "Lieferantenpaket")
        self.run_cfg = {"mode": mode, "files": files, "base": base, "target_dir": target_dir,
                        "bundle": bundle_name, "flatten": self.flatten_var.get()}
        self.log_dir = target_dir
        plan = []
        if mode == "einzeln":
            for path in files:
                plan.append(PlanItem(path.name, filetools.format_size(path.stat().st_size),
                                     str(target_dir / f"{path.stem}.zip"), payload=path))
        elif mode in ("zip", "targz"):
            ext = ".zip" if mode == "zip" else ".tar.gz"
            total = sum(p.stat().st_size for p in files)
            plan.append(PlanItem(f"{len(files)} Datei(en)", filetools.format_size(total),
                                 str(target_dir / (bundle_name + ext))))
        else:
            for path in files:
                plan.append(PlanItem(path.name, filetools.format_size(path.stat().st_size),
                                     f"entpacken nach {target_dir}", payload=path))
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        mode = cfg["mode"]
        if mode in ("zip", "targz"):
            ext = ".zip" if mode == "zip" else ".tar.gz"
            target = cfg["target_dir"] / (cfg["bundle"] + ext)
            filetools.bundle_archive(cfg["files"], target, base_dir=cfg["base"], flatten=cfg["flatten"],
                                     log=ctx.log, stopped=ctx.stopped, progress=ctx.progress)
            ctx.item(0, "OK")
            return
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path = item.payload
            try:
                if mode == "einzeln":
                    filetools.zip_single(path, cfg["target_dir"], log=ctx.log)
                else:
                    filetools.extract_archive(path, cfg["target_dir"], flatten=cfg["flatten"], log=ctx.log)
                ctx.item(idx, "OK")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))

    def option_state(self):
        return {"mode": self.mode_var.get(), "bundle": self.bundle_var.get(), "flatten": self.flatten_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "zip"))
        self.bundle_var.set(state.get("bundle", "Lieferantenpaket"))
        self.flatten_var.set(state.get("flatten", False))


# ---------------------------------------------------------------------------
# Tab: Ordnen
# ---------------------------------------------------------------------------

ORGANIZE_MODES = {
    "typ": "Nach Dateityp sortieren",
    "datum": "Nach Datum (JJJJ-MM) sortieren",
    "flach": "In Zielordner zusammenführen (flach)",
    "leer": "Leere Unterordner entfernen",
}


class OrdnenTab(ToolTab):
    key = "ordnen"
    title = "Ordnen"
    hint = "Downloads & Lieferantenordner aufräumen – mit Vorschau und Rückgängig"
    default_exts = ""
    target_subfolder = "_zusammen"

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        self.mode_var = tk.StringVar(value="typ")
        for value, label in ORGANIZE_MODES.items():
            ttk.Radiobutton(row1, text=label, value=value, variable=self.mode_var,
                            style="Card.TRadiobutton").pack(side="left", padx=(0, 14))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        self.copy_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="Kopieren statt Verschieben (Quelle bleibt unverändert)",
                        variable=self.copy_var, style="Card.TCheckbutton").pack(side="left")
        self.undo_button = ttk.Button(row2, text="Rückgängig (letzter Lauf)",
                                      command=self.undo, state="disabled")
        self.undo_button.pack(side="right")
        self.last_journal = []

    def make_plan(self):
        mode = self.mode_var.get()
        base = self.source.base_dir()
        self.log_dir = base
        if mode == "leer":
            if not self.source.sources:
                raise ValueError("Bitte zuerst einen Quellordner wählen.")
            empty = filetools.find_empty_dirs(base)
            if not empty:
                raise ValueError("Keine leeren Unterordner gefunden.")
            self.run_cfg = {"mode": mode, "dirs": empty}
            return [PlanItem(str(folder.relative_to(base)), "", "Ordner wird entfernt", payload=folder)
                    for folder in empty]
        files = self.source.collect()
        if mode == "flach":
            target_dir = self.target.resolve(base)
            move_plan = [(path, target_dir / path.name) for path in files]
            self.log_dir = target_dir
        else:
            move_plan = filetools.plan_organize(files, base, mode)
        if not move_plan:
            raise ValueError("Alles liegt schon am richtigen Platz – nichts zu tun.")
        self.run_cfg = {"mode": mode, "plan": move_plan, "copy": self.copy_var.get()}
        verb = "kopieren nach" if self.copy_var.get() else "verschieben nach"
        return [PlanItem(src.name, filetools.format_size(src.stat().st_size),
                         f"{verb} {dst.parent.name}/", payload=(src, dst))
                for src, dst in move_plan]

    def execute(self, ctx):
        cfg = self.run_cfg
        if cfg["mode"] == "leer":
            removed = 0
            for idx, item in enumerate(self.plan):
                if ctx.stopped():
                    break
                try:
                    item.payload.rmdir()
                    removed += 1
                    ctx.item(idx, "OK")
                except OSError as exc:
                    ctx.item(idx, "FEHLER")
                    ctx.log(f"FEHLER: {item.payload}: {exc}", "err")
                ctx.progress(idx + 1, len(self.plan))
            ctx.log(f"{removed} leere(r) Ordner entfernt.", "ok")
            return
        journal = []
        for idx, (source, target) in enumerate(cfg["plan"]):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                break
            try:
                entries = filetools.apply_moves([(source, target)], copy=cfg["copy"], log=ctx.log)
                journal.extend(entries)
                ctx.item(idx, "OK")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {source.name}: {exc}", "err")
            ctx.progress(idx + 1, len(cfg["plan"]))
        self.last_journal = journal
        if journal:
            journal_path = filetools.unique_path(
                Path(self.log_dir) / f"_ordnen_journal_{datetime.now():%Y%m%d_%H%M%S}.json")
            journal_path.write_text(json.dumps(journal, ensure_ascii=False, indent=2), encoding="utf-8")
            ctx.log(f"Journal gespeichert: {journal_path.name}")

    def after_run(self):
        self.undo_button.config(state="normal" if self.last_journal else "disabled")

    def undo(self):
        if not self.last_journal:
            return
        if not messagebox.askyesno(APP_TITLE, f"{len(self.last_journal)} Verschiebung(en) rückgängig machen?"):
            return
        restored = filetools.undo_moves(self.last_journal, log=self.app.log_line)
        self.app.log_line(f"Rückgängig abgeschlossen: {restored} Datei(en).", "ok")
        self.last_journal = []
        self.undo_button.config(state="disabled")

    def option_state(self):
        return {"mode": self.mode_var.get(), "copy": self.copy_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "typ"))
        self.copy_var.set(state.get("copy", False))


# ---------------------------------------------------------------------------
# Tab 4: Inventar / Prüfbericht
# ---------------------------------------------------------------------------

class InventarTab(ToolTab):
    key = "inventar"
    title = "Inventar"
    hint = "Lieferantenpakete dokumentieren: Manifest mit Hash und Duplikat-Prüfung"
    default_exts = ""
    target_subfolder = "_bericht"

    def build_options(self, parent):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x")
        self.hash_var = tk.BooleanVar(value=True)
        self.csv_var = tk.BooleanVar(value=True)
        self.html_var = tk.BooleanVar(value=True)
        self.xlsx_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="SHA-256 berechnen (findet Duplikate, dauert bei großen Dateien)",
                        variable=self.hash_var, style="Card.TCheckbutton").pack(side="left")
        ttk.Label(row, text="Bericht als", style="CardSub.TLabel").pack(side="left", padx=(18, 6))
        ttk.Checkbutton(row, text="CSV", variable=self.csv_var, style="Card.TCheckbutton").pack(side="left")
        ttk.Checkbutton(row, text="HTML", variable=self.html_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Checkbutton(row, text="XLSX", variable=self.xlsx_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Label(row, text="Name", style="CardSub.TLabel").pack(side="left", padx=(18, 6))
        self.report_var = tk.StringVar(value="inventar")
        ttk.Entry(row, textvariable=self.report_var, width=20).pack(side="left")
        self.sums_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="SHA256SUMS.txt erzeugen", variable=self.sums_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(14, 0))

    def make_plan(self):
        files = self.source.collect()
        formats = [ext for ext, var in ((".csv", self.csv_var), (".html", self.html_var), (".xlsx", self.xlsx_var)) if var.get()]
        if not formats:
            raise ValueError("Bitte mindestens ein Berichtsformat anhaken (CSV, HTML oder XLSX).")
        target_dir = self.target.resolve(self.source.base_dir())
        name = filetools.safe_filename(self.report_var.get().strip() or "inventar")
        self.run_cfg = {"files": files, "formats": formats, "target_dir": target_dir,
                        "name": name, "hash": self.hash_var.get(), "base": self.source.base_dir(),
                        "sums": self.sums_var.get()}
        self.log_dir = target_dir
        plan = [PlanItem(path.name, filetools.format_size(path.stat().st_size), "wird erfasst", payload=path)
                for path in files]
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg

        def progress(done, total):
            ctx.item(done - 1, "OK")
            ctx.progress(done, total)

        manifest = filetools.build_manifest(cfg["files"], base_dir=cfg["base"], with_hash=cfg["hash"],
                                            stopped=ctx.stopped, progress=progress, log=ctx.log)
        if ctx.stopped():
            ctx.log("Gestoppt.")
            return
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        for ext in cfg["formats"]:
            target = filetools.unique_path(cfg["target_dir"] / (cfg["name"] + ext))
            tabular.write_table(manifest, target)
            ctx.log(f"BERICHT: {target}", "ok")
        if cfg["sums"]:
            sums_path = filetools.unique_path(cfg["target_dir"] / "SHA256SUMS.txt")
            count = filetools.write_checksum_file(cfg["files"], cfg["base"], sums_path, stopped=ctx.stopped)
            ctx.log(f"PRÜFSUMMEN: {sums_path.name} ({count} Einträge)", "ok")

    def option_state(self):
        return {"hash": self.hash_var.get(), "csv": self.csv_var.get(), "html": self.html_var.get(),
                "xlsx": self.xlsx_var.get(), "name": self.report_var.get(), "sums": self.sums_var.get()}

    def apply_option_state(self, state):
        self.hash_var.set(state.get("hash", True))
        self.csv_var.set(state.get("csv", True))
        self.html_var.set(state.get("html", True))
        self.xlsx_var.set(state.get("xlsx", False))
        self.report_var.set(state.get("name", "inventar"))


# ---------------------------------------------------------------------------
# Tab 5: E-Mail-Anhänge
# ---------------------------------------------------------------------------

class EmlTab(ToolTab):
    key = "email"
    title = "E-Mail-Anhänge"
    hint = "Anhänge einsammeln oder Mails als HTML/Text archivieren (.eml)"
    default_exts = "eml"
    target_subfolder = "_anhaenge"

    def build_options(self, parent):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x")
        self.mode_var = tk.StringVar(value="anhaenge")
        ttk.Radiobutton(row, text="Anhänge extrahieren", value="anhaenge", variable=self.mode_var,
                        style="Card.TRadiobutton").pack(side="left", padx=(0, 14))
        ttk.Radiobutton(row, text="Als HTML konvertieren (mit eingebetteten Bildern)", value="html",
                        variable=self.mode_var, style="Card.TRadiobutton").pack(side="left", padx=(0, 14))
        ttk.Radiobutton(row, text="Als Text konvertieren", value="txt", variable=self.mode_var,
                        style="Card.TRadiobutton").pack(side="left", padx=(0, 18))
        self.subfolder_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row, text="Pro E-Mail einen Unterordner (nur Anhänge)",
                        variable=self.subfolder_var, style="Card.TCheckbutton").pack(side="left")

    def make_plan(self):
        files = self.source.collect()
        mode = self.mode_var.get()
        target_dir = self.target.resolve(self.source.base_dir())
        self.run_cfg = {"files": files, "target_dir": target_dir, "mode": mode,
                        "subfolder": self.subfolder_var.get()}
        self.log_dir = target_dir
        targets = {"anhaenge": "Anhänge extrahieren",
                   "html": lambda p: str(target_dir / (p.stem + ".html")),
                   "txt": lambda p: str(target_dir / (p.stem + ".txt"))}
        plan = []
        for path in files:
            label = targets[mode] if mode == "anhaenge" else targets[mode](path)
            plan.append(PlanItem(path.name, filetools.format_size(path.stat().st_size), label, payload=path))
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        total_attachments = 0
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path = item.payload
            try:
                if cfg["mode"] == "anhaenge":
                    count = filetools.extract_eml_attachments(path, cfg["target_dir"],
                                                              per_mail_subfolder=cfg["subfolder"], log=ctx.log)
                    total_attachments += count
                    ctx.item(idx, f"OK ({count})")
                else:
                    ext = ".html" if cfg["mode"] == "html" else ".txt"
                    target = filetools.unique_path(cfg["target_dir"] / (path.stem + ext))
                    if cfg["mode"] == "html":
                        filetools.eml_to_html(path, target)
                    else:
                        filetools.eml_to_text(path, target)
                    ctx.item(idx, "OK")
                    ctx.log(f"OK: {path.name} → {target.name}", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))
        if cfg["mode"] == "anhaenge":
            ctx.log(f"Gesamt: {total_attachments} Anhang/Anhänge gespeichert.", "ok")

    def option_state(self):
        return {"mode": self.mode_var.get(), "subfolder": self.subfolder_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "anhaenge"))
        self.subfolder_var.set(state.get("subfolder", True))


# ---------------------------------------------------------------------------
# Tab 6: Text-Werkzeuge
# ---------------------------------------------------------------------------

NEWLINE_LABELS = {"keep": "unverändert", "crlf": "Windows (CRLF)", "lf": "Unix (LF)"}


class TextTab(ToolTab):
    key = "text"
    title = "Text/Encoding"
    hint = "Encoding und Zeilenenden für ERP-Importe vereinheitlichen"
    default_exts = "txt, csv, tsv, md, json, xml, html"
    target_subfolder = "_text"

    def build_options(self, parent):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x")
        self.enc_combo = _labeled(row, "Ziel-Encoding",
                                  lambda p: ttk.Combobox(p, values=["utf-8", "utf-8-sig", "cp1252", "latin-1"],
                                                         width=10, state="readonly"))
        self.enc_combo.set("utf-8")
        self.newline_combo = _labeled(row, "Zeilenenden",
                                      lambda p: ttk.Combobox(p, values=list(NEWLINE_LABELS.values()),
                                                             width=16, state="readonly"))
        self.newline_combo.set(NEWLINE_LABELS["keep"])
        ttk.Label(row, text="Quell-Encoding wird automatisch erkannt (UTF-8/cp1252).",
                  style="CardSub.TLabel").pack(side="left", padx=(12, 0))

    def make_plan(self):
        files = self.source.collect()
        target_dir = self.target.resolve(self.source.base_dir())
        newline = next(k for k, v in NEWLINE_LABELS.items() if v == self.newline_combo.get())
        self.run_cfg = {"files": files, "target_dir": target_dir,
                        "encoding": self.enc_combo.get(), "newline": newline}
        self.log_dir = target_dir
        return [PlanItem(path.name, filetools.format_size(path.stat().st_size),
                         str(target_dir / path.name), payload=path) for path in files]

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path = item.payload
            try:
                target = filetools.unique_path(cfg["target_dir"] / path.name)
                detected = filetools.convert_text_file(path, target, cfg["encoding"], cfg["newline"])
                ctx.item(idx, "OK")
                ctx.log(f"OK: {path.name} ({detected} → {cfg['encoding']})", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))

    def option_state(self):
        return {"encoding": self.enc_combo.get(), "newline": self.newline_combo.get()}

    def apply_option_state(self, state):
        self.enc_combo.set(state.get("encoding", "utf-8"))
        self.newline_combo.set(state.get("newline", NEWLINE_LABELS["keep"]))


TAB_CLASSES = [TabellenTab, CadTab, RenameTab, PackTab, OrdnenTab, InventarTab, EmlTab, TextTab]


# ---------------------------------------------------------------------------
# Hauptfenster
# ---------------------------------------------------------------------------

class DataConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} {__version__}")
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.status_var = tk.StringVar(value="Bereit")

        self._configure_style()
        self._build_menu()
        self._build_ui()
        self.root.after(80, self._drain_queue)

    # -- Styling -----------------------------------------------------------

    def _configure_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", background=BG, foreground=TEXT, font=FONT)
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD)
        style.configure("CardTitle.TLabel", background=CARD, font=FONT_BOLD, foreground=TEXT)
        style.configure("CardSub.TLabel", background=CARD, foreground=SUBTLE, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=BG, foreground=SUBTLE)
        style.configure("Card.TCheckbutton", background=CARD, font=("Segoe UI", 9))
        style.map("Card.TCheckbutton", background=[("active", CARD)])
        style.configure("Card.TRadiobutton", background=CARD, font=("Segoe UI", 9))
        style.map("Card.TRadiobutton", background=[("active", CARD)])
        style.configure("TButton", padding=(12, 6))
        style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", borderwidth=0,
                        focusthickness=0, padding=(18, 7), font=FONT_BOLD)
        style.map("Accent.TButton",
                  background=[("active", ACCENT_DARK), ("disabled", "#9db9e8")],
                  foreground=[("disabled", "#f0f4fb")])
        style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 4, 0, 0))
        style.configure("TNotebook.Tab", padding=(16, 8), font=FONT)
        style.map("TNotebook.Tab",
                  background=[("selected", CARD), ("!selected", "#dde3ee")],
                  foreground=[("selected", ACCENT_DARK)])
        style.configure("Treeview", background=CARD, fieldbackground=CARD, rowheight=26,
                        font=("Segoe UI", 9), borderwidth=0)
        style.configure("Treeview.Heading", background="#e8edf5", font=("Segoe UI", 9, "bold"),
                        padding=(6, 5), relief="flat")
        style.map("Treeview.Heading", background=[("active", "#dfe6f1")])
        style.configure("TProgressbar", background=ACCENT, troughcolor="#dde3ee",
                        borderwidth=0, thickness=8)
        style.configure("TEntry", padding=3)
        self.root.configure(background=BG)

    # -- Menü --------------------------------------------------------------

    def _build_menu(self):
        self.root.option_add("*tearOff", False)
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar)
        file_menu.add_command(label="Preset speichern…", command=self.save_preset)
        file_menu.add_command(label="Preset laden…", command=self.load_preset)
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self.root.destroy)
        menubar.add_cascade(label="Datei", menu=file_menu)
        help_menu = tk.Menu(menubar)
        help_menu.add_command(label="Funktionsumfang", command=self.show_capabilities)
        help_menu.add_command(label="Über", command=self.show_about)
        menubar.add_cascade(label="Hilfe", menu=help_menu)
        self.root.config(menu=menubar)

    # -- Aufbau ------------------------------------------------------------

    def _build_ui(self):
        header = tk.Frame(self.root, bg=HEADER_BG)
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=HEADER_BG)
        title_box.pack(side="left", padx=18, pady=10)
        tk.Label(title_box, text=APP_TITLE, bg=HEADER_BG, fg="#ffffff", font=FONT_TITLE).pack(anchor="w")
        tk.Label(title_box, text="Werkzeugkasten für Lieferantendaten · 100 % Python-Bordmittel",
                 bg=HEADER_BG, fg="#a9b8d4", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(header, textvariable=self.status_var, bg=HEADER_BG, fg="#7ee2a8",
                 font=FONT_BOLD).pack(side="right", padx=18)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)
        self.tabs = []
        for tab_class in TAB_CLASSES:
            tab = tab_class(self, self.notebook)
            self.notebook.add(tab, text=f"  {tab.title}  ")
            self.tabs.append(tab)

        footer = ttk.Frame(self.root, padding=(14, 8, 14, 6))
        footer.pack(fill="x")
        self.progress = ttk.Progressbar(footer, mode="determinate", length=260)
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 14))
        self.stop_button = ttk.Button(footer, text="Stop", command=self.request_stop, state="disabled")
        self.stop_button.pack(side="right")
        self.run_button = ttk.Button(footer, text="Ausführen  ▶", style="Accent.TButton", command=self.start_run)
        self.run_button.pack(side="right", padx=(10, 10))
        ttk.Button(footer, text="Vorschau aktualisieren", command=self.preview).pack(side="right")

        log_frame = tk.Frame(self.root, bg=LOG_BG)
        log_frame.pack(fill="x", side="bottom")
        self.log_text = tk.Text(log_frame, height=7, bg=LOG_BG, fg="#d6e2f5", relief="flat",
                                font=("Consolas", 9), state="disabled", padx=10, pady=6, wrap="none")
        self.log_text.pack(fill="x", padx=2, pady=2)
        self.log_text.tag_configure("err", foreground="#ff8f8f")
        self.log_text.tag_configure("ok", foreground="#7ee2a8")
        self.log_line(f"{APP_TITLE} {__version__} gestartet – Werkzeug wählen, Quelle setzen, Vorschau, Ausführen.")

    # -- Aktionen ----------------------------------------------------------

    def current_tab(self):
        return self.tabs[self.notebook.index(self.notebook.select())]

    def preview(self):
        tab = self.current_tab()
        try:
            plan = tab.refresh_plan()
            self.log_line(f"[{tab.title}] Vorschau: {len(plan)} Aktion(en) geplant.")
            self.status_var.set("Plan bereit")
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))

    def start_run(self):
        if self.worker and self.worker.is_alive():
            return
        tab = self.current_tab()
        try:
            tab.refresh_plan()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return

        self.stop_event.clear()
        self.progress.configure(value=0, maximum=max(1, len(tab.plan)))
        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_var.set("Läuft …")
        log_path = None
        if tab.log_dir is not None:
            log_path = filetools.unique_path(Path(tab.log_dir) / f"lauf_log_{datetime.now():%Y%m%d_%H%M%S}.txt")
        ctx = RunContext(self.queue, self.stop_event, log_path)
        self.active_tab = tab

        def worker():
            started = datetime.now()
            try:
                tab.execute(ctx)
                seconds = (datetime.now() - started).total_seconds()
                if ctx.stopped():
                    self.queue.put(("done", f"[{tab.title}] Gestoppt nach {seconds:.1f} s."))
                else:
                    self.queue.put(("done", f"[{tab.title}] Fertig in {seconds:.1f} s."))
            except Exception as exc:
                ctx.log(f"ABBRUCH: {exc}", "err")
                ctx.log(traceback.format_exc(limit=3), "err")
                self.queue.put(("error", str(exc)))
            finally:
                self.queue.put(("state", "idle"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def request_stop(self):
        self.stop_event.set()
        self.status_var.set("Stop angefordert …")

    # -- Queue/Log ---------------------------------------------------------

    def log_line(self, message, level="info"):
        self.log_text.configure(state="normal")
        tag = () if level == "info" else (level,)
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _drain_queue(self):
        while True:
            try:
                message = self.queue.get_nowait()
            except queue.Empty:
                break
            kind = message[0]
            if kind == "log":
                self.log_line(message[1], message[2])
            elif kind == "item":
                self.active_tab.set_item_status(message[1], message[2])
            elif kind == "progress":
                _kind, done, total = message
                self.progress.configure(maximum=max(1, total), value=done)
            elif kind == "done":
                self.status_var.set("Fertig")
                self.log_line(message[1], "ok")
                self.progress.configure(value=self.progress["maximum"])
            elif kind == "error":
                self.status_var.set("Fehler")
                messagebox.showerror(APP_TITLE, message[1])
            elif kind == "state":
                self.run_button.config(state="normal")
                self.stop_button.config(state="disabled")
                if hasattr(self, "active_tab"):
                    self.active_tab.after_run()
        self.root.after(80, self._drain_queue)

    # -- Presets -----------------------------------------------------------

    def save_preset(self):
        path = filedialog.asksaveasfilename(title="Preset speichern", defaultextension=".json",
                                            filetypes=[("Preset (JSON)", "*.json")])
        if not path:
            return
        state = {
            "app": APP_TITLE, "version": __version__,
            "active_tab": self.notebook.index(self.notebook.select()),
            "tabs": {tab.key: tab.get_state() for tab in self.tabs},
        }
        Path(path).write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log_line(f"Preset gespeichert: {path}", "ok")

    def load_preset(self):
        path = filedialog.askopenfilename(title="Preset laden", filetypes=[("Preset (JSON)", "*.json")])
        if not path:
            return
        try:
            state = json.loads(Path(path).read_text(encoding="utf-8"))
            for tab in self.tabs:
                if tab.key in state.get("tabs", {}):
                    tab.set_state(state["tabs"][tab.key])
            index = state.get("active_tab", 0)
            if 0 <= index < len(self.tabs):
                self.notebook.select(index)
            self.log_line(f"Preset geladen: {path}", "ok")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Preset konnte nicht geladen werden:\n{exc}")

    # -- Hilfe -------------------------------------------------------------

    def show_capabilities(self):
        messagebox.showinfo(
            "Funktionsumfang",
            "Alles läuft mit reiner Python-Standardbibliothek – ohne Installationen:\n\n"
            "• Tabellen: CSV/TSV/TXT/XLSX/JSON/XML lesen → CSV/TSV/XLSX/JSON/XML/HTML/MD schreiben,\n"
            "  inkl. Dezimalzeichen, Spaltenauswahl, Duplikate, Zusammenführen.\n"
            "• CAD: STL/OBJ/PLY/3MF konvertieren, STEP mit eigenem B-Rep-Kern tessellieren\n"
            "  (→ STL/OBJ/PLY/3MF/GLB/HTML-3D-Ansicht), STEP/IGES-Prüfbericht.\n"
            "• Umbenennen: Suchen/Ersetzen, Präfix/Suffix, Nummern, Datum – mit Vorschau und Rückgängig.\n"
            "• Packen: ZIP/TAR.GZ erstellen, ZIP/TAR/GZ/BZ2/XZ entpacken.\n"
            "• Ordnen: nach Typ/Datum sortieren, Ordner glätten, leere Ordner entfernen – mit Rückgängig.\n"
            "• Inventar: Manifest mit Größe, Datum, SHA-256, Duplikat-Prüfung, SHA256SUMS.txt.\n"
            "• E-Mail: Anhänge extrahieren, Mails als HTML/Text archivieren (.eml).\n"
            "• Text: Encoding und Zeilenenden konvertieren.\n\n"
            "Nicht möglich ohne Hersteller-Kernel: SLDPRT/IPT/eDrawings lesen (proprietär) –\n"
            "beim Lieferanten STEP oder STL anfordern. Ebenso: Office→PDF, RAR/7z, Medien.",
        )

    def show_about(self):
        messagebox.showinfo("Über", f"{APP_TITLE} {__version__}\n\nWerkzeugkasten für die tägliche "
                                    "Einkaufsarbeit.\nReine Python-Standardbibliothek, keine Abhängigkeiten.")


def _enable_windows_dpi():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def main(argv=None):
    argv = argv or []
    if "--self-test" in argv:
        from .selftest import run_self_test
        run_self_test()
        return
    _enable_windows_dpi()
    root = tk.Tk()
    root.geometry("1180x760")
    root.minsize(980, 640)
    DataConverterApp(root)
    root.mainloop()
