import csv
import json
import queue
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, Toplevel, filedialog, messagebox
from tkinter import ttk


APP_TITLE = "Einkauf Data Converter"


FORMAT_GROUPS = {
    "CAD 3D Native": {
        "inputs": [
            ".asm", ".catpart", ".catproduct", ".cgr", ".ipt", ".iam", ".model",
            ".neu", ".par", ".prt", ".psm", ".sldasm", ".sldprt", ".xpr", ".xas",
        ],
        "outputs": [".step", ".stp", ".jt", ".iges", ".igs", ".x_t", ".stl", ".obj", ".3mf", ".pdf"],
    },
    "CAD 3D Neutral": {
        "inputs": [
            ".3dxml", ".3mf", ".3dm", ".brep", ".fbx", ".glb", ".gltf", ".ifc",
            ".iges", ".igs", ".jt", ".obj", ".sat", ".sab", ".stl", ".step",
            ".stp", ".usd", ".usdz", ".wrl", ".x_b", ".x_t",
        ],
        "outputs": [".step", ".stp", ".jt", ".iges", ".igs", ".stl", ".obj", ".3mf", ".glb", ".ifc", ".pdf"],
    },
    "CAD 2D": {
        "inputs": [".dwg", ".dxf", ".dwf", ".dwfx", ".dgn", ".hpgl", ".plt"],
        "outputs": [".dxf", ".dwg", ".pdf", ".svg", ".png"],
    },
    "CAM/NC": {
        "inputs": [".cnc", ".drl", ".g", ".gcode", ".nc", ".ncc", ".tap"],
        "outputs": [".txt", ".zip"],
    },
    "Elektronik": {
        "inputs": [".brd", ".dsn", ".gbr", ".ger", ".kicad_pcb", ".pcb", ".sch"],
        "outputs": [".pdf", ".zip", ".dxf", ".svg"],
    },
    "Tabellen": {
        "inputs": [".accdb", ".csv", ".db", ".json", ".mdb", ".ods", ".parquet", ".sqlite", ".tsv", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xml", ".yaml", ".yml"],
        "outputs": [".csv", ".json", ".tsv", ".xlsx", ".xml"],
    },
    "MS Office": {
        "inputs": [".doc", ".docm", ".docx", ".dot", ".dotx", ".odp", ".ods", ".odt", ".ott", ".ppt", ".pptm", ".pptx", ".rtf", ".vsd", ".vsdx", ".xls", ".xlsb", ".xlsm", ".xlsx"],
        "outputs": [".docx", ".pdf", ".pptx", ".xlsx", ".odt", ".ods", ".odp"],
    },
    "Dokumente/Text": {
        "inputs": [".eml", ".html", ".htm", ".md", ".msg", ".tex", ".txt", ".xhtml", ".xml"],
        "outputs": [".pdf", ".docx", ".html", ".txt", ".md"],
    },
    "Bilder": {
        "inputs": [".avif", ".bmp", ".cr2", ".dng", ".gif", ".heic", ".heif", ".jfif", ".jp2", ".jpeg", ".jpg", ".jxl", ".png", ".psd", ".raw", ".tga", ".tif", ".tiff", ".webp"],
        "outputs": [".jpg", ".png", ".pdf", ".tiff", ".webp", ".bmp"],
    },
    "Vektor/Print": {
        "inputs": [".ai", ".cdr", ".eps", ".indd", ".ps", ".svg"],
        "outputs": [".pdf", ".svg", ".png", ".jpg"],
    },
    "PDF": {
        "inputs": [".pdf"],
        "outputs": [".pdf", ".png", ".jpg", ".txt", ".zip"],
    },
    "Archive": {
        "inputs": [".7z", ".bz2", ".gz", ".rar", ".tar", ".tgz", ".xz", ".zip"],
        "outputs": [".zip", ".tar"],
    },
    "Medien": {
        "inputs": [".avi", ".mov", ".mp3", ".mp4", ".wav", ".webm", ".wmv"],
        "outputs": [".mp4", ".webm", ".mp3", ".wav", ".zip"],
    },
}


PURCHASING_RECOMMENDATIONS = [
    "STEP/STP fuer neutrale 3D-CAD-Daten im Einkauf",
    "PDF fuer Zeichnungen, Spezifikationen und Freigaben",
    "XLSX/CSV fuer Preislisten, Stuecklisten und Lieferantendaten",
    "PNG/JPG/WebP fuer Bilder, Etiketten und Screenshots",
    "ZIP fuer gebuendelte Lieferantenpakete",
]


def all_input_formats():
    values = set()
    for group in FORMAT_GROUPS.values():
        values.update(group["inputs"])
    return sorted(values)


def suggested_outputs(input_exts):
    outputs = set()
    selected = set(input_exts)
    for group in FORMAT_GROUPS.values():
        if selected.intersection(group["inputs"]):
            outputs.update(group["outputs"])
    return sorted(outputs)


def summarize(values, empty="Auswaehlen"):
    if not values:
        return empty
    if len(values) <= 3:
        return ", ".join(values)
    return f"{', '.join(values[:3])} +{len(values) - 3}"


@dataclass
class SpecialOptions:
    zip_outputs: bool = False
    compress_images: bool = False
    redact_placeholder: bool = False
    search_text: str = ""
    replace_text: str = ""
    prefix: str = ""
    suffix: str = ""
    rename_inputs: bool = False


@dataclass
class JobConfig:
    source: Path
    source_kind: str
    input_exts: list[str]
    output_exts: list[str]
    contains: str = ""
    max_mb: str = ""
    special: SpecialOptions = field(default_factory=SpecialOptions)


class FileLogger:
    def __init__(self, path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(f"{APP_TITLE} Log\nStart: {datetime.now():%Y-%m-%d %H:%M:%S}\n\n", encoding="utf-8")

    def __call__(self, message):
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(f"{datetime.now():%H:%M:%S}  {message}\n")


class ConverterBackend:
    def __init__(self, log, should_stop):
        self.log = log
        self.should_stop = should_stop

    def collect_files(self, job, recursive):
        if job.source_kind == "file":
            candidates = [job.source]
        elif recursive:
            candidates = [p for p in job.source.rglob("*") if p.is_file()]
        else:
            candidates = [p for p in job.source.iterdir() if p.is_file()]

        selected = []
        max_bytes = None
        if job.max_mb.strip():
            try:
                max_bytes = float(job.max_mb.replace(",", ".")) * 1024 * 1024
            except ValueError:
                self.log(f"Filter ignoriert: '{job.max_mb}' ist keine gueltige MB-Zahl.")

        for path in candidates:
            if job.input_exts and path.suffix.lower() not in job.input_exts:
                continue
            if job.contains and job.contains.lower() not in path.name.lower():
                continue
            if max_bytes is not None and path.stat().st_size > max_bytes:
                continue
            selected.append(path)
        return selected

    def run_job(self, job, output_root, same_folder, recursive):
        files = self.collect_files(job, recursive)
        self.log(f"{len(files)} Datei(en) gefunden fuer {job.source}")
        for path in files:
            if self.should_stop():
                self.log("Gestoppt.")
                return

            active_source = self.rename_source_if_needed(path, job.special)
            target_dir = active_source.parent if same_folder else output_root
            target_dir.mkdir(parents=True, exist_ok=True)
            generated = []

            for out_ext in job.output_exts:
                if self.should_stop():
                    return
                target_name = self.build_target_stem(active_source.stem, job.special) + out_ext
                target = self.unique_path(target_dir / target_name)
                try:
                    self.convert(active_source, target, job.special)
                    generated.append(target)
                    self.log(f"OK: {active_source.name} -> {target.name}")
                except Exception as exc:
                    self.log(f"FEHLER: {active_source.name} -> {out_ext}: {exc}")

            if job.special.zip_outputs and generated:
                zip_path = self.unique_path(target_dir / f"{self.build_target_stem(active_source.stem, job.special)}_converted.zip")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
                    for item in generated:
                        if item.exists():
                            archive.write(item, arcname=item.name)
                self.log(f"ZIP: {zip_path.name}")

    def rename_source_if_needed(self, source, special):
        if not special.rename_inputs:
            return source
        new_stem = self.build_target_stem(source.stem, special)
        target = self.unique_path(source.with_name(new_stem + source.suffix))
        if target == source:
            return source
        source.rename(target)
        self.log(f"UMBENANNT: {source.name} -> {target.name}")
        return target

    def build_target_stem(self, stem, special):
        if special.search_text:
            stem = stem.replace(special.search_text, special.replace_text)
        return f"{special.prefix}{stem}{special.suffix}"

    def unique_path(self, path):
        if not path.exists():
            return path
        counter = 2
        while True:
            candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
            if not candidate.exists():
                return candidate
            counter += 1

    def convert(self, source, target, special):
        src_ext = source.suffix.lower()
        dst_ext = target.suffix.lower()

        if src_ext == dst_ext:
            shutil.copy2(source, target)
            return

        if src_ext in {".csv", ".tsv", ".json"} and dst_ext in {".csv", ".tsv", ".json"}:
            self.convert_structured_text(source, target)
            return

        if src_ext in FORMAT_GROUPS["Bilder"]["inputs"] and dst_ext in FORMAT_GROUPS["Bilder"]["outputs"]:
            self.convert_image(source, target, special)
            return

        if dst_ext == ".zip":
            with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.write(source, arcname=source.name)
            return

        if src_ext in FORMAT_GROUPS["MS Office"]["inputs"] and dst_ext in {".pdf", ".docx", ".pptx", ".xlsx", ".odt", ".ods", ".odp"}:
            self.convert_with_libreoffice(source, target)
            return

        cad_3d_inputs = set(FORMAT_GROUPS["CAD 3D Native"]["inputs"]) | set(FORMAT_GROUPS["CAD 3D Neutral"]["inputs"])
        if src_ext in cad_3d_inputs and dst_ext in {".step", ".stp", ".stl", ".obj", ".3mf", ".ifc", ".iges", ".igs"}:
            self.convert_with_freecad(source, target)
            return

        if src_ext in FORMAT_GROUPS["CAD 2D"]["inputs"] and dst_ext in {".dxf", ".pdf"}:
            raise RuntimeError("DWG/DXF-Konvertierung benoetigt ODA File Converter oder ein CAD-Backend.")

        raise RuntimeError("Diese Kombination braucht ein externes Backend oder ist nicht sinnvoll.")

    def convert_structured_text(self, source, target):
        if source.suffix.lower() == ".json":
            data = json.loads(source.read_text(encoding="utf-8"))
            rows = data if isinstance(data, list) else [data]
        else:
            delimiter = "\t" if source.suffix.lower() == ".tsv" else ","
            with source.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle, delimiter=delimiter))

        if target.suffix.lower() == ".json":
            target.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
            return

        delimiter = "\t" if target.suffix.lower() == ".tsv" else ","
        keys = sorted({key for row in rows if isinstance(row, dict) for key in row.keys()})
        with target.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=keys, delimiter=delimiter)
            writer.writeheader()
            for row in rows:
                writer.writerow(row if isinstance(row, dict) else {"value": row})

    def convert_image(self, source, target, special):
        try:
            from PIL import Image, ImageDraw
        except ImportError as exc:
            raise RuntimeError("Pillow fehlt. Installieren mit: pip install pillow") from exc

        image = Image.open(source)
        if special.redact_placeholder:
            image = image.convert("RGB")
            draw = ImageDraw.Draw(image)
            w, h = image.size
            draw.rectangle((int(w * 0.05), int(h * 0.05), int(w * 0.35), int(h * 0.18)), fill="white")
        if target.suffix.lower() in {".jpg", ".jpeg", ".pdf"}:
            image = image.convert("RGB")
        save_kwargs = {}
        if special.compress_images and target.suffix.lower() in {".jpg", ".jpeg", ".webp"}:
            save_kwargs["quality"] = 75
            save_kwargs["optimize"] = True
        image.save(target, **save_kwargs)

    def convert_with_libreoffice(self, source, target):
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise RuntimeError("LibreOffice ist nicht im PATH. Fuer Office/PDF-Konvertierung LibreOffice installieren.")
        subprocess.run(
            [soffice, "--headless", "--convert-to", target.suffix.lower().lstrip("."), "--outdir", str(target.parent), str(source)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        produced = target.parent / f"{source.stem}{target.suffix.lower()}"
        if produced.exists() and produced != target:
            produced.replace(target)

    def convert_with_freecad(self, source, target):
        freecad = shutil.which("FreeCADCmd") or shutil.which("freecadcmd")
        if not freecad:
            raise RuntimeError("FreeCADCmd ist nicht im PATH. Native CAD-Dateien brauchen FreeCAD oder Hersteller-Exporter.")
        script = (
            "import Import, Mesh, sys\n"
            f"src = r'''{source}'''\n"
            f"dst = r'''{target}'''\n"
            "doc = App.newDocument()\n"
            "Import.insert(src, doc.Name)\n"
            "objs = doc.Objects\n"
            "if dst.lower().endswith(('.stl', '.obj')):\n"
            "    Mesh.export(objs, dst)\n"
            "else:\n"
            "    Import.export(objs, dst)\n"
        )
        script_path = target.parent / "_freecad_convert_tmp.py"
        script_path.write_text(script, encoding="utf-8")
        try:
            subprocess.run([freecad, str(script_path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        finally:
            script_path.unlink(missing_ok=True)


class FormatPicker(Toplevel):
    def __init__(self, parent, title, values_by_group, selected, on_save):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, True)
        self.on_save = on_save
        self.vars = {}
        self.configure(padx=14, pady=14)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)
        for group, values in values_by_group.items():
            frame = ttk.Frame(notebook, padding=10)
            notebook.add(frame, text=group)
            for idx, value in enumerate(values):
                var = BooleanVar(value=value in selected)
                self.vars[value] = var
                ttk.Checkbutton(frame, text=value, variable=var).grid(row=idx // 4, column=idx % 4, sticky="w", padx=8, pady=4)

        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Alle abwaehlen", command=self.clear).pack(side="left")
        ttk.Button(actions, text="OK", command=self.save).pack(side="right")

    def clear(self):
        for var in self.vars.values():
            var.set(False)

    def save(self):
        self.on_save(sorted(value for value, var in self.vars.items() if var.get()))
        self.destroy()


class JobRow:
    def __init__(self, app, parent, index):
        self.app = app
        self.source_kind = StringVar(value="folder")
        self.source_var = StringVar()
        self.contains_var = StringVar()
        self.max_mb_var = StringVar()
        self.input_exts = []
        self.output_exts = []
        self.special = SpecialOptions()

        self.frame = ttk.Frame(parent, padding=(8, 6))
        self.number = ttk.Label(self.frame, text=str(index), width=3, anchor="center")
        self.number.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Entry(self.frame, textvariable=self.source_var, width=38).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(self.frame, text="Ordner", command=self.pick_folder).grid(row=0, column=2, padx=2)
        ttk.Button(self.frame, text="Datei", command=self.pick_file).grid(row=0, column=3, padx=2)

        self.input_button = ttk.Button(self.frame, text="Auswaehlen", command=self.pick_inputs)
        self.input_button.grid(row=0, column=4, sticky="ew", padx=4)
        ttk.Entry(self.frame, textvariable=self.contains_var, width=18).grid(row=0, column=5, sticky="ew", padx=4)
        ttk.Entry(self.frame, textvariable=self.max_mb_var, width=10).grid(row=0, column=6, sticky="ew", padx=4)
        self.output_button = ttk.Button(self.frame, text="Nach Input", command=self.pick_outputs)
        self.output_button.grid(row=0, column=7, sticky="ew", padx=4)
        self.special_button = ttk.Button(self.frame, text="Zahnrad", command=self.open_specials)
        self.special_button.grid(row=0, column=8, padx=4)
        ttk.Button(self.frame, text="-", width=3, command=lambda: app.remove_row(self)).grid(row=0, column=9, padx=(6, 0))

        self.frame.columnconfigure(1, weight=3)
        self.frame.columnconfigure(4, weight=1)
        self.frame.columnconfigure(5, weight=1)
        self.frame.columnconfigure(7, weight=1)

    def set_index(self, index):
        self.number.config(text=str(index))

    def pick_folder(self):
        path = filedialog.askdirectory(title="Ordner waehlen")
        if path:
            self.source_kind.set("folder")
            self.source_var.set(path)

    def pick_file(self):
        path = filedialog.askopenfilename(title="Datei waehlen")
        if path:
            self.source_kind.set("file")
            self.source_var.set(path)
            ext = Path(path).suffix.lower()
            if ext in all_input_formats() and ext not in self.input_exts:
                self.input_exts = [ext]
                self.output_exts = []
                self.refresh_buttons()

    def pick_inputs(self):
        groups = {name: group["inputs"] for name, group in FORMAT_GROUPS.items()}
        FormatPicker(self.app.root, "Inputformate", groups, self.input_exts, self.save_inputs)

    def save_inputs(self, values):
        self.input_exts = values
        allowed_outputs = set(suggested_outputs(self.input_exts))
        self.output_exts = [ext for ext in self.output_exts if ext in allowed_outputs]
        self.refresh_buttons()

    def pick_outputs(self):
        outputs = suggested_outputs(self.input_exts)
        groups = {"Empfohlen": outputs or [".pdf", ".zip"]}
        FormatPicker(self.app.root, "Outputformate", groups, self.output_exts, self.save_outputs)

    def save_outputs(self, values):
        self.output_exts = values
        self.refresh_buttons()

    def refresh_buttons(self):
        self.input_button.config(text=summarize(self.input_exts))
        self.output_button.config(text=summarize(self.output_exts, "Nach Input"))
        active_specials = sum([
            self.special.zip_outputs,
            self.special.compress_images,
            self.special.redact_placeholder,
            bool(self.special.search_text or self.special.prefix or self.special.suffix),
            self.special.rename_inputs,
        ])
        self.special_button.config(text=f"Zahnrad ({active_specials})" if active_specials else "Zahnrad")

    def open_specials(self):
        window = Toplevel(self.app.root)
        window.title("Spezialoperationen")
        window.resizable(False, False)
        window.configure(padx=16, pady=16)

        zip_var = BooleanVar(value=self.special.zip_outputs)
        compress_var = BooleanVar(value=self.special.compress_images)
        redact_var = BooleanVar(value=self.special.redact_placeholder)
        rename_inputs_var = BooleanVar(value=self.special.rename_inputs)
        search_var = StringVar(value=self.special.search_text)
        replace_var = StringVar(value=self.special.replace_text)
        prefix_var = StringVar(value=self.special.prefix)
        suffix_var = StringVar(value=self.special.suffix)

        ops = ttk.LabelFrame(window, text="Operationen", padding=10)
        ops.pack(fill="x")
        ttk.Checkbutton(ops, text="Ergebnisse zusaetzlich zippen", variable=zip_var).grid(row=0, column=0, sticky="w", pady=3)
        ttk.Checkbutton(ops, text="Bilder komprimieren", variable=compress_var).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Checkbutton(ops, text="Demo-Zone in Bildern/PDF weiss ueberdecken", variable=redact_var).grid(row=2, column=0, sticky="w", pady=3)

        naming = ttk.LabelFrame(window, text="Namenskonventionen", padding=10)
        naming.pack(fill="x", pady=(12, 0))
        ttk.Label(naming, text="Suchen").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(naming, textvariable=search_var, width=24).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(naming, text="Ersetzen").grid(row=0, column=2, sticky="w", padx=(12, 6), pady=4)
        ttk.Entry(naming, textvariable=replace_var, width=24).grid(row=0, column=3, sticky="ew", pady=4)
        ttk.Label(naming, text="Praefix").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(naming, textvariable=prefix_var, width=24).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(naming, text="Suffix").grid(row=1, column=2, sticky="w", padx=(12, 6), pady=4)
        ttk.Entry(naming, textvariable=suffix_var, width=24).grid(row=1, column=3, sticky="ew", pady=4)
        ttk.Checkbutton(naming, text="Auch Inputdateien umbenennen", variable=rename_inputs_var).grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        def save():
            self.special.zip_outputs = zip_var.get()
            self.special.compress_images = compress_var.get()
            self.special.redact_placeholder = redact_var.get()
            self.special.rename_inputs = rename_inputs_var.get()
            self.special.search_text = search_var.get()
            self.special.replace_text = replace_var.get()
            self.special.prefix = prefix_var.get()
            self.special.suffix = suffix_var.get()
            self.refresh_buttons()
            window.destroy()

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="OK", command=save).pack(side="right")

    def config(self):
        source_text = self.source_var.get().strip()
        if not source_text:
            raise ValueError("Eine Zeile hat keinen Pfad.")
        if not self.output_exts:
            raise ValueError(f"Keine Outputformate gewaehlt fuer {source_text}.")
        return JobConfig(
            source=Path(source_text),
            source_kind=self.source_kind.get(),
            input_exts=self.input_exts,
            output_exts=self.output_exts,
            contains=self.contains_var.get().strip(),
            max_mb=self.max_mb_var.get().strip(),
            special=self.special,
        )


class DataConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.rows = []
        self.queue = queue.Queue()
        self.worker = None
        self.stop_event = threading.Event()
        self.log_file = None

        self.same_folder_var = BooleanVar(value=True)
        self.recursive_var = BooleanVar(value=True)
        self.output_path_var = StringVar()
        self.status_var = StringVar(value="Bereit")

        self.configure_style()
        self.build_menu()
        self.build_ui()
        self.add_row()
        self.root.after(100, self.drain_queue)

    def configure_style(self):
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("Header.TLabel", background="#f6f7f9", font=("Segoe UI", 13, "bold"))
        style.configure("Subtle.TLabel", background="#f6f7f9", foreground="#586174")
        style.configure("TableHead.TLabel", background="#e9edf3", foreground="#222", font=("Segoe UI", 9, "bold"), padding=(6, 7))
        style.configure("TButton", padding=(10, 5))
        style.configure("TCheckbutton", background="#f6f7f9")
        style.configure("TLabelframe", background="#f6f7f9")
        style.configure("TLabelframe.Label", background="#f6f7f9")

    def build_menu(self):
        menu = ttk.Frame(self.root)
        self.root.option_add("*tearOff", False)
        native_menu = __import__("tkinter").Menu(self.root)
        help_menu = __import__("tkinter").Menu(native_menu, tearoff=False)
        help_menu.add_command(label="Formate und Empfehlung", command=self.show_formats)
        native_menu.add_cascade(label="Hilfe", menu=help_menu)
        self.root.config(menu=native_menu)
        return menu

    def build_ui(self):
        self.root.configure(background="#f6f7f9")
        page = ttk.Frame(self.root, padding=14)
        page.pack(fill="both", expand=True)

        header = ttk.Frame(page)
        header.pack(fill="x")
        ttk.Label(header, text="Einkauf Data Converter", style="Header.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var, style="Subtle.TLabel").pack(side="right")

        table = ttk.Frame(page, padding=(0, 14, 0, 0))
        table.pack(fill="both", expand=True)
        self.header_frame = ttk.Frame(table)
        self.header_frame.pack(fill="x")
        headers = ["#", "Pfad", "", "", "Input", "Filter Name", "Max MB", "Output", "Optionen", ""]
        widths = [3, 38, 7, 6, 18, 18, 9, 18, 11, 3]
        for idx, (text, width) in enumerate(zip(headers, widths)):
            label = ttk.Label(self.header_frame, text=text, width=width, style="TableHead.TLabel")
            label.grid(row=0, column=idx, sticky="ew", padx=1)
        self.header_frame.columnconfigure(1, weight=3)
        self.header_frame.columnconfigure(4, weight=1)
        self.header_frame.columnconfigure(5, weight=1)
        self.header_frame.columnconfigure(7, weight=1)

        self.rows_frame = ttk.Frame(table)
        self.rows_frame.pack(fill="x")

        footer = ttk.Frame(page)
        footer.pack(fill="x", pady=(12, 0))
        ttk.Button(footer, text="+ Zeile", command=self.add_row).pack(side="left")
        ttk.Checkbutton(footer, text="Output im selben Ordner erstellen", variable=self.same_folder_var, command=self.toggle_output_path).pack(side="left", padx=(18, 0))
        ttk.Checkbutton(footer, text="Rekursiv", variable=self.recursive_var).pack(side="left", padx=(14, 0))
        self.output_entry = ttk.Entry(footer, textvariable=self.output_path_var, state="disabled")
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(14, 4))
        self.output_button = ttk.Button(footer, text="Outputpfad", command=self.pick_output, state="disabled")
        self.output_button.pack(side="left")
        self.go_button = ttk.Button(footer, text="Go", command=self.start)
        self.go_button.pack(side="left", padx=(14, 4))
        self.stop_button = ttk.Button(footer, text="Stop", command=self.stop, state="disabled")
        self.stop_button.pack(side="left")

    def add_row(self):
        row = JobRow(self, self.rows_frame, len(self.rows) + 1)
        row.frame.pack(fill="x", pady=(1, 0))
        self.rows.append(row)

    def remove_row(self, row):
        if len(self.rows) == 1:
            messagebox.showinfo(APP_TITLE, "Mindestens eine Zeile bleibt erhalten.")
            return
        self.rows.remove(row)
        row.frame.destroy()
        for idx, item in enumerate(self.rows, start=1):
            item.set_index(idx)

    def toggle_output_path(self):
        state = "disabled" if self.same_folder_var.get() else "normal"
        self.output_entry.config(state=state)
        self.output_button.config(state=state)
        if not self.same_folder_var.get() and not self.output_path_var.get():
            self.pick_output()

    def pick_output(self):
        path = filedialog.askdirectory(title="Output-Ordner waehlen")
        if path:
            self.output_path_var.set(path)

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        try:
            jobs = [row.config() for row in self.rows]
            for job in jobs:
                if not job.source.exists():
                    raise ValueError(f"Pfad existiert nicht: {job.source}")
            output_root = None
            if not self.same_folder_var.get():
                if not self.output_path_var.get().strip():
                    raise ValueError("Bitte Outputpfad waehlen.")
                output_root = Path(self.output_path_var.get().strip())
            self.log_file = self.create_log_file(jobs, output_root)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self.stop_event.clear()
        self.go_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_var.set("Laeuft...")
        self.worker = threading.Thread(
            target=self.run_worker,
            args=(jobs, output_root, self.same_folder_var.get(), self.recursive_var.get(), self.log_file),
            daemon=True,
        )
        self.worker.start()

    def create_log_file(self, jobs, output_root):
        if output_root:
            log_dir = output_root
        else:
            first = jobs[0].source
            log_dir = first if first.is_dir() else first.parent
        return log_dir / f"conversion_log_{datetime.now():%Y%m%d_%H%M%S}.txt"

    def stop(self):
        self.stop_event.set()
        self.status_var.set("Stop angefordert...")

    def run_worker(self, jobs, output_root, same_folder, recursive, log_file):
        logger = FileLogger(log_file)
        backend = ConverterBackend(logger, self.stop_event.is_set)
        started = time.time()
        try:
            for job in jobs:
                if self.stop_event.is_set():
                    break
                backend.run_job(job, output_root or job.source.parent, same_folder, recursive)
            logger(f"Fertig nach {time.time() - started:.1f}s.")
            self.queue.put(("DONE", str(log_file)))
        except Exception as exc:
            logger(f"ABBRUCH: {exc}")
            self.queue.put(("ERROR", str(exc)))
        finally:
            self.queue.put(("STATE", "idle"))

    def drain_queue(self):
        while True:
            try:
                kind, message = self.queue.get_nowait()
            except queue.Empty:
                break
            if kind == "DONE":
                self.status_var.set(f"Fertig. Log: {Path(message).name}")
                messagebox.showinfo(APP_TITLE, f"Fertig. Log gespeichert:\n{message}")
            elif kind == "ERROR":
                self.status_var.set("Fehler")
                messagebox.showerror(APP_TITLE, message)
            elif kind == "STATE":
                self.go_button.config(state="normal")
                self.stop_button.config(state="disabled")
        self.root.after(100, self.drain_queue)

    def show_formats(self):
        lines = ["Empfohlen fuer Einkauf:"]
        lines.extend(f"- {item}" for item in PURCHASING_RECOMMENDATIONS)
        lines.append("")
        lines.append("Lieferbare Formatfamilien:")
        for name, group in FORMAT_GROUPS.items():
            lines.append(f"- {name}: Input {', '.join(group['inputs'])} -> Output {', '.join(group['outputs'])}")
        lines.append("")
        lines.append("Hinweis: CAD STEP aus nativen Formaten braucht FreeCAD, Hersteller-Exporter oder kommerzielle Konverter.")
        messagebox.showinfo("Formate", "\n".join(lines))


def self_test():
    assert ".step" in suggested_outputs([".sldprt"])
    assert ".xlsx" in suggested_outputs([".csv"])
    assert ".pdf" in suggested_outputs([".docx"])
    assert ".pdf" in suggested_outputs([".dwg"])
    assert ".zip" in suggested_outputs([".gbr"])
    assert ".webp" in suggested_outputs([".heic"])
    special = SpecialOptions(search_text="ALT", replace_text="NEU", prefix="P_", suffix="_S")
    backend = ConverterBackend(lambda _message: None, lambda: False)
    assert backend.build_target_stem("ALT_Name", special) == "P_NEU_Name_S"
    print("Self-test OK")


def main():
    if "--self-test" in sys.argv:
        self_test()
        return
    root = Tk()
    root.geometry("1220x420")
    root.minsize(1040, 320)
    DataConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
