"""Einkauf Data Converter - Werkzeugkasten fuer Lieferantendaten (All-in-One).

Eine Datei, keine Abhaengigkeiten: reine Python-Standardbibliothek.
Start:            python data_converter_gui.py
Selbsttest:       python data_converter_gui.py --self-test
Headless:         python data_converter_gui.py --preset lauf.json --run
Watch-Ordner:     python data_converter_gui.py --preset lauf.json --watch 30

Werkzeuge: Tabellen (CSV/TSV/XLSX/JSON/XML), CAD (STL/OBJ/PLY/3MF, STEP- und
IGES-Tessellierungs-Kern, DXF->SVG, HTML-3D-Viewer), PDF (mergen/splitten/
drehen/umsortieren/abdecken), Bilder (PNG/BMP), Umbenennen, Packen, Ordnen,
Inventar, E-Mail (.eml), Text-Encoding - jeweils mit Vorschau/Dry-Run,
Live-Log, Undo und Presets.
"""

__version__ = "2.3.0"
APP_TITLE = "Einkauf Data Converter"

from collections import Counter
from dataclasses import dataclass
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.parser import BytesParser
from html import unescape as html_unescape
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import base64
import bz2
import csv
import gzip
import hashlib
import json
import lzma
import math
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import tarfile
import tempfile
import threading
import tkinter as tk
import traceback
import zipfile
import zlib


# ===========================================================================
# XLSX-Reader/-Writer (zipfile + ElementTree)
# ===========================================================================

"""Minimaler XLSX-Reader und -Writer auf Basis der Standardbibliothek.

XLSX ist ein ZIP-Container mit XML-Dateien. Fuer einfache Datentabellen
(eine Arbeitsmappe, ein Blatt, Text und Zahlen) reicht zipfile + ElementTree.
"""


_NS_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# Zahl im Excel-Sinn: optionales Minus, Ziffern, optionaler Dezimalpunkt.
# Fuehrende Nullen (Artikelnummern wie "007") bleiben bewusst Text.
_NUMBER_RE = re.compile(r"^-?(0|[1-9]\d*)(\.\d+)?$")


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _col_index(cell_ref):
    """'BC12' -> Spaltenindex 54 (0-basiert)."""
    idx = 0
    for ch in cell_ref:
        if ch.isalpha():
            idx = idx * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return idx - 1


def _col_letter(index):
    """0-basierter Index -> 'A', 'B', ... 'AA'."""
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _element_text(element):
    """Alle <t>-Texte unterhalb eines Elements zusammensetzen (rich text)."""
    parts = []
    for node in element.iter():
        if _local(node.tag) == "t" and node.text:
            parts.append(node.text)
    return "".join(parts)


def _shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [_element_text(si) for si in root if _local(si.tag) == "si"]


def _first_sheet_path(archive):
    names = set(archive.namelist())
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet = next(el for el in workbook.iter() if _local(el.tag) == "sheet")
        rid = sheet.get(_NS_REL_ID)
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target", "").lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                if target in names:
                    return target
    except (KeyError, StopIteration, ET.ParseError):
        pass
    if "xl/worksheets/sheet1.xml" in names:
        return "xl/worksheets/sheet1.xml"
    for name in sorted(names):
        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
            return name
    raise ValueError("XLSX enthaelt kein Tabellenblatt.")


def _format_number(raw):
    """Excel speichert Zahlen als '12.4' oder '3' - ganze Zahlen ohne '.0' zeigen."""
    try:
        value = float(raw)
    except ValueError:
        return raw
    if value.is_integer() and "e" not in raw.lower() and abs(value) < 1e15:
        return str(int(value))
    return raw


def read_xlsx(path):
    """Erstes Tabellenblatt lesen. Liefert Zeilen als Listen von Strings."""
    with zipfile.ZipFile(path) as archive:
        shared = _shared_strings(archive)
        sheet_root = ET.fromstring(archive.read(_first_sheet_path(archive)))

    rows = []
    for row_el in sheet_root.iter():
        if _local(row_el.tag) != "row":
            continue
        cells = {}
        auto_col = 0
        for cell in row_el:
            if _local(cell.tag) != "c":
                continue
            ref = cell.get("r")
            col = _col_index(ref) if ref else auto_col
            auto_col = col + 1
            cells[col] = _cell_value(cell, shared)
        if cells:
            width = max(cells) + 1
            rows.append([cells.get(i, "") for i in range(width)])
        else:
            rows.append([])

    # Alle Zeilen auf gleiche Breite bringen, leere Schlusszeilen entfernen.
    width = max((len(r) for r in rows), default=0)
    rows = [row + [""] * (width - len(row)) for row in rows]
    while rows and all(cell == "" for cell in rows[-1]):
        rows.pop()
    return rows


def _cell_value(cell, shared):
    cell_type = cell.get("t", "n")
    if cell_type == "inlineStr":
        return _element_text(cell)
    value = ""
    for child in cell:
        if _local(child.tag) == "v":
            value = child.text or ""
            break
    if cell_type == "s":
        try:
            return shared[int(value)]
        except (ValueError, IndexError):
            return value
    if cell_type == "b":
        return "WAHR" if value == "1" else "FALSCH"
    if cell_type in ("str", "e"):
        return value
    return _format_number(value) if value else ""


# ---------------------------------------------------------------------------
# Schreiben
# ---------------------------------------------------------------------------

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
</styleSheet>"""


def _sheet_name_safe(name):
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name).strip() or "Daten"
    return cleaned[:31]


def is_number(value):
    return bool(_NUMBER_RE.match(value)) if isinstance(value, str) else isinstance(value, (int, float))


def _cell_xml(row_num, col_idx, value, style_id=0):
    ref = f"{_col_letter(col_idx)}{row_num}"
    style = f' s="{style_id}"' if style_id else ""
    text = "" if value is None else str(value)
    if text and is_number(text):
        return f'<c r="{ref}"{style}><v>{text}</v></c>'
    if not text:
        return f'<c r="{ref}"{style}/>'
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}"{style} t="inlineStr"><is><t{preserve}>{escape(text)}</t></is></c>'


def write_xlsx(path, headers, rows, sheet_name="Daten"):
    """Tabelle als XLSX schreiben. Kopfzeile fett, Spaltenbreiten angepasst."""
    all_rows = ([list(headers)] if headers else []) + [list(r) for r in rows]
    col_count = max((len(r) for r in all_rows), default=1)

    widths = []
    for col in range(col_count):
        longest = max((len(str(r[col])) for r in all_rows if col < len(r) and r[col] is not None), default=0)
        widths.append(min(max(longest + 2, 9), 60))

    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    parts.append("<cols>")
    for idx, width in enumerate(widths, start=1):
        parts.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
    parts.append("</cols><sheetData>")
    for row_num, row in enumerate(all_rows, start=1):
        style = 1 if (headers and row_num == 1) else 0
        cells = "".join(_cell_xml(row_num, col, value, style) for col, value in enumerate(row))
        parts.append(f'<row r="{row_num}">{cells}</row>')
    parts.append("</sheetData></worksheet>")
    sheet_xml = "".join(parts)

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{escape(_sheet_name_safe(sheet_name))}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _CONTENT_TYPES)
        archive.writestr("_rels/.rels", _ROOT_RELS)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
        archive.writestr("xl/styles.xml", _STYLES)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


# ===========================================================================
# Tabellen lesen, aufbereiten, schreiben
# ===========================================================================

"""Tabellen lesen, aufbereiten und schreiben - reine Standardbibliothek.

Zentrales Modell ist `Table`: eine Kopfzeile plus Datenzeilen (Strings).
Damit lassen sich Preislisten, Stuecklisten und Lieferantendaten zwischen
CSV/TSV/TXT, JSON, XML und XLSX wandeln und dabei normalisieren.
"""



TABLE_READ_EXTS = [".csv", ".tsv", ".txt", ".xlsx", ".json", ".xml"]
TABLE_WRITE_EXTS = [".csv", ".tsv", ".xlsx", ".json", ".xml", ".html", ".md"]
DELIMITER_CHOICES = {";": ";", ",": ",", "Tab": "\t", "|": "|"}
ENCODING_CHOICES = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]

_COMMA_DECIMAL = re.compile(r"^-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d+$")
_DOT_DECIMAL = re.compile(r"^-?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d+$")


@dataclass
class Table:
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)

    @property
    def column_count(self):
        return max([len(self.headers)] + [len(r) for r in self.rows]) if (self.headers or self.rows) else 0

    def normalized(self):
        """Alle Zeilen auf Headerbreite bringen, fehlende Header auffuellen."""
        width = self.column_count
        headers = list(self.headers) + [f"Spalte {i + 1}" for i in range(len(self.headers), width)]
        rows = [[_to_text(cell) for cell in row] + [""] * (width - len(row)) for row in self.rows]
        return Table(headers, rows)


def _to_text(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "WAHR" if value else "FALSCH"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ---------------------------------------------------------------------------
# Lesen
# ---------------------------------------------------------------------------

def decode_bytes(data):
    """Encoding pragmatisch erkennen: erst UTF-8 (mit/ohne BOM), dann cp1252."""
    for encoding in ("utf-8-sig", "utf-8"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("cp1252", errors="replace"), "cp1252"


def read_text_auto(path):
    return decode_bytes(Path(path).read_bytes())


def sniff_delimiter(sample, fallback=";"):
    try:
        return csv.Sniffer().sniff(sample[:4096], delimiters=";,\t|").delimiter
    except csv.Error:
        first_line = sample.splitlines()[0] if sample.splitlines() else ""
        counts = {d: first_line.count(d) for d in (";", ",", "\t", "|")}
        best = max(counts, key=counts.get)
        return best if counts[best] else fallback


def read_table(path):
    """Datei anhand der Endung als Tabelle einlesen."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".csv", ".tsv", ".txt"):
        return _read_delimited(path, ext)
    if ext == ".xlsx":
        rows = read_xlsx(path)
        if not rows:
            return Table()
        return Table([_to_text(c) for c in rows[0]], rows[1:]).normalized()
    if ext == ".json":
        return _read_json(path)
    if ext == ".xml":
        return _read_xml(path)
    raise ValueError(f"Nicht unterstuetztes Tabellenformat: {ext}")


def _read_delimited(path, ext):
    text, _encoding = read_text_auto(path)
    delimiter = "\t" if ext == ".tsv" else sniff_delimiter(text)
    rows = list(csv.reader(text.splitlines(), delimiter=delimiter))
    rows = [row for row in rows if row]
    if not rows:
        return Table()
    return Table([c.strip() for c in rows[0]], rows[1:]).normalized()


def _read_json(path):
    text, _encoding = read_text_auto(path)
    data = json.loads(text)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list) or not data:
        return Table()
    if all(isinstance(item, dict) for item in data):
        headers = []
        for item in data:
            for key in item:
                if key not in headers:
                    headers.append(str(key))
        rows = [[_to_text(item.get(key, "")) for key in headers] for item in data]
        return Table(headers, rows)
    if all(isinstance(item, list) for item in data):
        return Table([_to_text(c) for c in data[0]], data[1:]).normalized()
    return Table(["Wert"], [[_to_text(item)] for item in data])


def _read_xml(path):
    root = ET.parse(path).getroot()
    records = list(root)
    if not records:
        return Table()
    headers = []
    parsed = []
    for record in records:
        values = {}
        for attr, value in record.attrib.items():
            values[attr] = value
        for child in record:
            tag = child.tag.rsplit("}", 1)[-1]
            values[tag] = (child.text or "").strip()
        if not values and (record.text or "").strip():
            values["Wert"] = record.text.strip()
        for key in values:
            if key not in headers:
                headers.append(key)
        parsed.append(values)
    rows = [[_to_text(item.get(key, "")) for key in headers] for item in parsed]
    return Table(headers, rows)


# ---------------------------------------------------------------------------
# Aufbereitung
# ---------------------------------------------------------------------------

@dataclass
class TransformOptions:
    trim: bool = True
    drop_empty_rows: bool = True
    dedupe: bool = False
    decimal: str = "keep"          # keep | comma_to_dot | dot_to_comma
    column_spec: str = ""          # "alt>neu; spalte2" - leer = alle Spalten


def parse_column_spec(spec):
    """'Preis>Netto; SKU' -> [("Preis", "Netto"), ("SKU", "SKU")]."""
    entries = []
    for chunk in re.split(r"[;,]", spec):
        chunk = chunk.strip()
        if not chunk:
            continue
        if ">" in chunk:
            old, new = chunk.split(">", 1)
            entries.append((old.strip(), new.strip() or old.strip()))
        else:
            entries.append((chunk, chunk))
    return entries


def _convert_decimal(cell, mode):
    if mode == "comma_to_dot" and _COMMA_DECIMAL.match(cell):
        return cell.replace(".", "").replace(",", ".")
    if mode == "dot_to_comma" and _DOT_DECIMAL.match(cell):
        return cell.replace(",", "").replace(".", ",")
    return cell


def apply_transform(table, options):
    """Aufbereitung anwenden. Liefert (neue Tabelle, Warnungen)."""
    table = table.normalized()
    warnings = []
    headers = list(table.headers)
    rows = [list(row) for row in table.rows]

    if options.trim:
        headers = [h.strip() for h in headers]
        rows = [[cell.strip() for cell in row] for row in rows]

    if options.decimal != "keep":
        rows = [[_convert_decimal(cell, options.decimal) for cell in row] for row in rows]

    spec = parse_column_spec(options.column_spec)
    if spec:
        lookup = {h.lower(): i for i, h in enumerate(headers)}
        indices, new_headers = [], []
        for old, new in spec:
            idx = lookup.get(old.lower())
            if idx is None:
                warnings.append(f"Spalte nicht gefunden: {old}")
                continue
            indices.append(idx)
            new_headers.append(new)
        if indices:
            headers = new_headers
            rows = [[row[i] for i in indices] for row in rows]
        else:
            warnings.append("Keine der angegebenen Spalten gefunden - alle Spalten beibehalten.")

    if options.drop_empty_rows:
        rows = [row for row in rows if any(cell != "" for cell in row)]

    if options.dedupe:
        seen = set()
        unique = []
        for row in rows:
            key = tuple(row)
            if key not in seen:
                seen.add(key)
                unique.append(row)
        removed = len(rows) - len(unique)
        if removed:
            warnings.append(f"{removed} Duplikat-Zeile(n) entfernt.")
        rows = unique

    return Table(headers, rows), warnings


def merge_tables(named_tables, source_column="Quelle"):
    """Mehrere Tabellen zu einer zusammenfuehren (Header-Vereinigung)."""
    headers = [source_column] if source_column else []
    for _name, table in named_tables:
        for header in table.normalized().headers:
            if header not in headers:
                headers.append(header)
    rows = []
    for name, table in named_tables:
        table = table.normalized()
        index = {h: i for i, h in enumerate(table.headers)}
        for row in table.rows:
            merged = []
            for header in headers:
                if source_column and header == source_column:
                    merged.append(name)
                else:
                    i = index.get(header)
                    merged.append(row[i] if i is not None else "")
            rows.append(merged)
    return Table(headers, rows)


# ---------------------------------------------------------------------------
# Schreiben
# ---------------------------------------------------------------------------

def write_table(table, path, delimiter=";", encoding="utf-8-sig"):
    """Tabelle anhand der Zielendung schreiben."""
    path = Path(path)
    ext = path.suffix.lower()
    table = table.normalized()
    if ext in (".csv", ".txt"):
        _write_delimited(table, path, delimiter, encoding)
    elif ext == ".tsv":
        _write_delimited(table, path, "\t", encoding)
    elif ext == ".xlsx":
        write_xlsx(path, table.headers, table.rows, sheet_name=path.stem)
    elif ext == ".json":
        records = [dict(zip(table.headers, row)) for row in table.rows]
        path.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    elif ext == ".xml":
        path.write_text(_to_xml(table), encoding="utf-8")
    elif ext == ".html":
        path.write_text(_to_html(table, title=path.stem), encoding="utf-8")
    elif ext == ".md":
        path.write_text(_to_markdown(table), encoding="utf-8")
    else:
        raise ValueError(f"Nicht unterstuetztes Zielformat: {ext}")


def _write_delimited(table, path, delimiter, encoding):
    with path.open("w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle, delimiter=delimiter)
        writer.writerow(table.headers)
        writer.writerows(table.rows)


def _xml_tag(name):
    tag = re.sub(r"[^0-9A-Za-z_.-]", "_", name.strip()) or "feld"
    if not re.match(r"^[A-Za-z_]", tag):
        tag = "c_" + tag
    return tag


def _to_xml(table):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<records>"]
    tags = [_xml_tag(h) for h in table.headers]
    for row in table.rows:
        lines.append("  <record>")
        for tag, cell in zip(tags, row):
            lines.append(f"    <{tag}>{escape(cell)}</{tag}>")
        lines.append("  </record>")
    lines.append("</records>")
    return "\n".join(lines) + "\n"


def _to_html(table, title="Tabelle"):
    head_cells = "".join(f"<th>{escape(h)}</th>" for h in table.headers)
    body = []
    for row in table.rows:
        cells = "".join(f"<td>{escape(cell)}</td>" for cell in row)
        body.append(f"<tr>{cells}</tr>")
    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>{escape(title)}</title>
<style>
body {{ font-family: "Segoe UI", system-ui, sans-serif; margin: 2rem; color: #1c2430; background: #f6f8fb; }}
h1 {{ font-size: 1.25rem; }}
p.meta {{ color: #5b6676; font-size: .85rem; }}
table {{ border-collapse: collapse; background: #fff; box-shadow: 0 1px 4px rgba(20,30,50,.12); }}
th, td {{ padding: .45rem .8rem; border: 1px solid #dde3ec; font-size: .9rem; text-align: left; }}
th {{ background: #22304a; color: #fff; position: sticky; top: 0; }}
tr:nth-child(even) td {{ background: #f2f5fa; }}
</style>
</head>
<body>
<h1>{escape(title)}</h1>
<p class="meta">{len(table.rows)} Zeilen, {len(table.headers)} Spalten</p>
<table>
<thead><tr>{head_cells}</tr></thead>
<tbody>
{chr(10).join(body)}
</tbody>
</table>
</body>
</html>
"""


def _to_markdown(table):
    def cell_text(value):
        return value.replace("|", "\\|").replace("\n", " ")

    lines = ["| " + " | ".join(cell_text(h) for h in table.headers) + " |"]
    lines.append("| " + " | ".join("---" for _ in table.headers) + " |")
    for row in table.rows:
        lines.append("| " + " | ".join(cell_text(cell) for cell in row) + " |")
    return "\n".join(lines) + "\n"


# ===========================================================================
# Datei-Werkzeuge: Umbenennen, Archive, Ordnen, Inventar, EML, Encoding
# ===========================================================================

"""Datei-Werkzeuge: Scannen, Umbenennen (mit Undo), Archive, Inventar,
E-Mail-Anhaenge und Text-Encoding - reine Standardbibliothek."""



ARCHIVE_EXTS = [".zip", ".tar", ".gz", ".tgz", ".bz2", ".xz", ".7z", ".rar"]
_SINGLE_COMPRESSED = {".gz": gzip.open, ".bz2": bz2.open, ".xz": lzma.open}


# ---------------------------------------------------------------------------
# Scannen & Pfade
# ---------------------------------------------------------------------------

def parse_ext_filter(text):
    """'csv; xlsx, .pdf' -> {'.csv', '.xlsx', '.pdf'} (leer = kein Filter)."""
    exts = set()
    for chunk in re.split(r"[;,\s]+", text.strip()):
        chunk = chunk.strip().lower().lstrip("*")
        if chunk:
            exts.add(chunk if chunk.startswith(".") else "." + chunk)
    return exts


def iter_files(sources, recursive=True, exts=None, contains="", max_mb=None):
    """Dateien aus Ordnern/Einzeldateien einsammeln und filtern."""
    seen = set()
    result = []
    max_bytes = max_mb * 1024 * 1024 if max_mb else None
    for source in sources:
        source = Path(source)
        if source.is_file():
            candidates = [source]
        elif source.is_dir():
            candidates = sorted(p for p in (source.rglob("*") if recursive else source.iterdir()) if p.is_file())
        else:
            raise ValueError(f"Pfad existiert nicht: {source}")
        for path in candidates:
            if path in seen:
                continue
            if exts and path.suffix.lower() not in exts:
                continue
            if contains and contains.lower() not in path.name.lower():
                continue
            if max_bytes is not None and path.stat().st_size > max_bytes:
                continue
            seen.add(path)
            result.append(path)
    return result


def unique_path(path):
    """Bei Namenskollision '_2', '_3', ... anhaengen."""
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def format_size(num_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < 1024 or unit == "GB":
            return f"{num_bytes:,.0f} {unit}".replace(",", ".") if unit == "B" else f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024


def safe_filename(name, fallback="datei"):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().strip(".")
    return cleaned or fallback


# ---------------------------------------------------------------------------
# Umbenennen mit Vorschau und Undo
# ---------------------------------------------------------------------------

@dataclass
class RenameOptions:
    search: str = ""
    replace: str = ""
    use_regex: bool = False
    prefix: str = ""
    suffix: str = ""
    case: str = "keep"            # keep | lower | upper | title
    numbering: bool = False
    number_start: int = 1
    number_digits: int = 3
    date_stamp: bool = False      # Aenderungsdatum der Datei als Praefix


def build_new_name(path, options, number=None):
    stem = path.stem
    if options.search:
        if options.use_regex:
            try:
                stem = re.sub(options.search, options.replace, stem)
            except re.error as exc:
                raise ValueError(f"Ungueltiger regulaerer Ausdruck: {exc}") from exc
        else:
            stem = stem.replace(options.search, options.replace)
    if options.case == "lower":
        stem = stem.lower()
    elif options.case == "upper":
        stem = stem.upper()
    elif options.case == "title":
        stem = stem.title()
    stem = f"{options.prefix}{stem}{options.suffix}"
    if options.numbering and number is not None:
        stem = f"{number:0{max(1, options.number_digits)}d}_{stem}"
    if options.date_stamp:
        stamp = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d")
        stem = f"{stamp}_{stem}"
    return safe_filename(stem, fallback=path.stem) + path.suffix


def plan_rename(files, options):
    """Liefert Liste (Quelle, neuer Name). Unveraenderte Dateien fallen raus."""
    plan = []
    number = options.number_start
    for path in files:
        new_name = build_new_name(path, options, number)
        if options.numbering:
            number += 1
        if new_name != path.name:
            plan.append((path, new_name))
    return plan


def apply_rename(plan, log=print, stopped=lambda: False, progress=None):
    """Plan ausfuehren, Journal fuer Undo zurueckgeben."""
    journal = []
    for index, (source, new_name) in enumerate(plan):
        if stopped():
            break
        target = unique_path(source.with_name(new_name))
        source.rename(target)
        journal.append({"von": str(source), "nach": str(target)})
        log(f"UMBENANNT: {source.name} -> {target.name}")
        if progress:
            progress(index + 1, len(plan))
    if journal:
        journal_path = unique_path(Path(plan[0][0]).parent / f"_umbenennen_journal_{datetime.now():%Y%m%d_%H%M%S}.json")
        journal_path.write_text(json.dumps(journal, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"Journal gespeichert: {journal_path.name}")
        return journal, journal_path
    return journal, None


def undo_rename(journal, log=print):
    """Journal rueckwaerts abarbeiten."""
    restored = 0
    for entry in reversed(journal):
        target = Path(entry["nach"])
        original = Path(entry["von"])
        if not target.exists():
            log(f"UEBERSPRUNGEN (fehlt): {target.name}")
            continue
        original = unique_path(original) if original.exists() else original
        target.rename(original)
        restored += 1
        log(f"ZURUECK: {target.name} -> {original.name}")
    return restored


# ---------------------------------------------------------------------------
# Archive
# ---------------------------------------------------------------------------

def zip_single(source, target_dir, log=print):
    target = unique_path(target_dir / f"{source.stem}.zip")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(source, arcname=source.name)
    log(f"ZIP: {source.name} -> {target.name}")
    return target


def bundle_archive(files, target_path, base_dir=None, flatten=False, log=print,
                   stopped=lambda: False, progress=None):
    """Alle Dateien in ein ZIP- oder TAR.GZ-Paket packen."""
    target_path = unique_path(target_path)
    names_used = set()

    def arcname_for(path):
        if not flatten and base_dir:
            try:
                return str(path.relative_to(base_dir))
            except ValueError:
                pass
        name = path.name
        counter = 2
        while name in names_used:
            name = f"{path.stem}_{counter}{path.suffix}"
            counter += 1
        names_used.add(name)
        return name

    if target_path.suffix.lower() == ".zip":
        with zipfile.ZipFile(target_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for index, path in enumerate(files):
                if stopped():
                    break
                archive.write(path, arcname=arcname_for(path))
                if progress:
                    progress(index + 1, len(files))
    else:
        with tarfile.open(target_path, "w:gz") as archive:
            for index, path in enumerate(files):
                if stopped():
                    break
                archive.add(path, arcname=arcname_for(path))
                if progress:
                    progress(index + 1, len(files))
    log(f"PAKET: {target_path.name} ({len(files)} Datei(en))")
    return target_path


def split_zip_bundles(files, target_dir, name, limit_mb, log=print,
                      stopped=lambda: False, progress=None):
    """Dateien auf mehrere unabhängige ZIPs mit Maximalgröße verteilen.

    Jedes Teil ist allein entpackbar (kein Multi-Volume-Archiv) - gedacht
    für E-Mail-Anhang-Limits. Die Grenze bezieht sich auf die Rohgröße der
    Eingaben; komprimiert sind die Teile eher kleiner.
    """
    limit = max(1, int(limit_mb)) * 1024 * 1024
    bundles = []
    current, current_size = [], 0
    for path in files:
        size = path.stat().st_size
        if size > limit:
            log(f"HINWEIS: {path.name} ({format_size(size)}) ist größer als das Limit "
                f"und bekommt ein eigenes Teil.")
        if current and current_size + size > limit:
            bundles.append(current)
            current, current_size = [], 0
        current.append(path)
        current_size += size
    if current:
        bundles.append(current)

    written = []
    total = len(files)
    done = 0
    for index, bundle in enumerate(bundles, start=1):
        if stopped():
            break
        target = unique_path(Path(target_dir) / f"{name}_teil{index:02d}.zip")
        names_used = set()
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
            for path in bundle:
                arcname = path.name
                counter = 2
                while arcname in names_used:
                    arcname = f"{path.stem}_{counter}{path.suffix}"
                    counter += 1
                names_used.add(arcname)
                archive.write(path, arcname=arcname)
                done += 1
                if progress:
                    progress(done, total)
        written.append(target)
        log(f"TEIL {index}/{len(bundles)}: {target.name} "
            f"({len(bundle)} Datei(en), {format_size(target.stat().st_size)})")
    return written


def _is_within(base, target):
    try:
        target.resolve().relative_to(base.resolve())
        return True
    except ValueError:
        return False


def extract_archive(archive_path, target_dir, flatten=False, log=print):
    """ZIP/TAR/GZ/BZ2/XZ entpacken - mit Schutz gegen Pfad-Ausbrueche."""
    archive_path = Path(archive_path)
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    suffix = archive_path.suffix.lower()
    extracted = 0

    if suffix == ".zip":
        with zipfile.ZipFile(archive_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                relative = Path(info.filename)
                out_name = relative.name if flatten else relative
                destination = target_dir / out_name
                if not _is_within(target_dir, destination):
                    log(f"UEBERSPRUNGEN (unsicherer Pfad): {info.filename}")
                    continue
                destination = unique_path(destination) if flatten else destination
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as src, destination.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
                extracted += 1
    elif suffix in (".tar", ".tgz") or archive_path.name.lower().endswith((".tar.gz", ".tar.bz2", ".tar.xz")):
        with tarfile.open(archive_path) as archive:
            members = [m for m in archive.getmembers() if m.isfile()]
            for member in members:
                relative = Path(member.name)
                out_name = relative.name if flatten else relative
                destination = target_dir / out_name
                if not _is_within(target_dir, destination):
                    log(f"UEBERSPRUNGEN (unsicherer Pfad): {member.name}")
                    continue
                destination = unique_path(destination) if flatten else destination
                destination.parent.mkdir(parents=True, exist_ok=True)
                source = archive.extractfile(member)
                if source is None:
                    continue
                with source, destination.open("wb") as dst:
                    shutil.copyfileobj(source, dst)
                extracted += 1
    elif suffix in _SINGLE_COMPRESSED:
        opener = _SINGLE_COMPRESSED[suffix]
        destination = unique_path(target_dir / archive_path.stem)
        with opener(archive_path, "rb") as src, destination.open("wb") as dst:
            shutil.copyfileobj(src, dst)
        extracted = 1
    elif suffix in (".7z", ".rar"):
        raise ValueError(f"{suffix.upper()} kann die Standardbibliothek nicht entpacken - bitte als ZIP anliefern lassen.")
    else:
        raise ValueError(f"Unbekanntes Archivformat: {archive_path.name}")

    log(f"ENTPACKT: {archive_path.name} -> {extracted} Datei(en)")
    return extracted


# ---------------------------------------------------------------------------
# Inventar / Pruefbericht
# ---------------------------------------------------------------------------

def sha256_of(path, chunk_size=1024 * 1024):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def build_manifest(files, base_dir=None, with_hash=True, stopped=lambda: False,
                   progress=None, log=print):
    """Manifest-Tabelle: Pfad, Typ, Groesse, Datum, Hash, Duplikat-Markierung."""
    entries = []
    for index, path in enumerate(files):
        if stopped():
            break
        stat = path.stat()
        try:
            relative = str(path.relative_to(base_dir)) if base_dir else path.name
        except ValueError:
            relative = str(path)
        entries.append({
            "Datei": relative,
            "Typ": path.suffix.lower().lstrip(".") or "-",
            "Groesse": format_size(stat.st_size),
            "Bytes": str(stat.st_size),
            "Geaendert": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
            "SHA-256": sha256_of(path) if with_hash else "",
        })
        if progress:
            progress(index + 1, len(files))

    if with_hash:
        counts = {}
        for entry in entries:
            counts[entry["SHA-256"]] = counts.get(entry["SHA-256"], 0) + 1
        duplicates = 0
        for entry in entries:
            is_dup = counts[entry["SHA-256"]] > 1
            entry["Duplikat"] = "JA" if is_dup else ""
            duplicates += 1 if is_dup else 0
        if duplicates:
            log(f"Achtung: {duplicates} Datei(en) mit identischem Inhalt gefunden.")

    headers = ["Datei", "Typ", "Groesse", "Bytes", "Geaendert"] + (["SHA-256", "Duplikat"] if with_hash else [])
    rows = [[entry.get(h, "") for h in headers] for entry in entries]
    return Table(headers, rows)


def write_checksum_file(files, base_dir, target_path, stopped=lambda: False, progress=None):
    """SHA256SUMS.txt im ueblichen Format 'hash  relativer/pfad' schreiben."""
    lines = []
    for index, path in enumerate(files):
        if stopped():
            break
        try:
            relative = path.relative_to(base_dir).as_posix()
        except ValueError:
            relative = path.name
        lines.append(f"{sha256_of(path)}  {relative}")
        if progress:
            progress(index + 1, len(files))
    Path(target_path).write_text("\n".join(lines) + "\n", encoding="utf-8")
    return len(lines)


# ---------------------------------------------------------------------------
# Ordnen: sortieren, glaetten, leere Ordner
# ---------------------------------------------------------------------------

TYPE_FOLDERS = {
    "Tabellen": {".csv", ".tsv", ".xlsx", ".xls", ".xlsm", ".ods", ".json", ".xml"},
    "Dokumente": {".pdf", ".doc", ".docx", ".odt", ".rtf", ".txt", ".md", ".ppt", ".pptx"},
    "CAD": {".step", ".stp", ".iges", ".igs", ".stl", ".obj", ".ply", ".3mf", ".dxf", ".dwg",
            ".sldprt", ".sldasm", ".slddrw", ".ipt", ".iam", ".idw", ".x_t", ".jt", ".glb"},
    "Bilder": {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg", ".heic"},
    "Archive": {".zip", ".tar", ".gz", ".tgz", ".bz2", ".xz", ".7z", ".rar"},
    "E-Mails": {".eml", ".msg"},
}


def folder_for_type(path):
    ext = path.suffix.lower()
    for folder, exts in TYPE_FOLDERS.items():
        if ext in exts:
            return folder
    return "Sonstiges"


def plan_organize(files, base_dir, mode):
    """Plan (Quelle -> Zielpfad) fuer 'typ' oder 'datum'. Nichts wird bewegt."""
    plan = []
    for path in files:
        if mode == "typ":
            folder = folder_for_type(path)
        else:
            folder = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m")
        target = Path(base_dir) / folder / path.name
        if target != path:
            plan.append((path, target))
    return plan


def apply_moves(plan, copy=False, log=print, stopped=lambda: False, progress=None):
    """Verschieben/Kopieren nach Plan. Liefert Journal (fuer Undo bei Verschieben)."""
    journal = []
    verb = "KOPIERT" if copy else "VERSCHOBEN"
    for index, (source, target) in enumerate(plan):
        if stopped():
            break
        target = unique_path(Path(target))
        target.parent.mkdir(parents=True, exist_ok=True)
        if copy:
            shutil.copy2(source, target)
        else:
            shutil.move(str(source), str(target))
            journal.append({"von": str(source), "nach": str(target)})
        log(f"{verb}: {source.name} -> {target.parent.name}/{target.name}")
        if progress:
            progress(index + 1, len(plan))
    return journal


def undo_moves(journal, log=print):
    """Verschiebe-Journal rueckwaerts abarbeiten (auch ueber Laufwerksgrenzen)."""
    restored = 0
    for entry in reversed(journal):
        target = Path(entry["nach"])
        original = Path(entry["von"])
        if not target.exists():
            log(f"UEBERSPRUNGEN (fehlt): {target.name}")
            continue
        original.parent.mkdir(parents=True, exist_ok=True)
        original = unique_path(original) if original.exists() else original
        shutil.move(str(target), str(original))
        restored += 1
        log(f"ZURUECK: {target.name}")
    return restored


def find_empty_dirs(base_dir):
    """Leere Ordner (auch verschachtelt leere) unterhalb von base_dir."""
    base_dir = Path(base_dir)
    empty = []
    for folder in sorted((p for p in base_dir.rglob("*") if p.is_dir()),
                         key=lambda p: len(p.parts), reverse=True):
        entries = list(folder.iterdir())
        if not entries or all(e in empty for e in entries):
            empty.append(folder)
    return empty


# ---------------------------------------------------------------------------
# E-Mail-Anhaenge (.eml)
# ---------------------------------------------------------------------------

def extract_eml_attachments(eml_path, target_dir, per_mail_subfolder=True, log=print):
    """Anhaenge einer .eml-Datei speichern. Liefert Anzahl."""
    eml_path = Path(eml_path)
    with eml_path.open("rb") as handle:
        message = BytesParser(policy=policy.default).parse(handle)

    destination = Path(target_dir)
    if per_mail_subfolder:
        destination = destination / safe_filename(eml_path.stem)
    saved = 0
    for part in message.iter_attachments():
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        name = safe_filename(part.get_filename() or "anhang.bin")
        destination.mkdir(parents=True, exist_ok=True)
        out_path = unique_path(destination / name)
        out_path.write_bytes(payload)
        saved += 1
        log(f"ANHANG: {eml_path.name} -> {out_path.name}")
    if not saved:
        log(f"KEINE ANHAENGE: {eml_path.name}")
    return saved


def _load_eml(eml_path):
    with Path(eml_path).open("rb") as handle:
        return BytesParser(policy=policy.default).parse(handle)


def _eml_header_rows(message):
    rows = []
    for label, key in (("Von", "From"), ("An", "To"), ("Cc", "Cc"),
                       ("Datum", "Date"), ("Betreff", "Subject")):
        value = message.get(key)
        if value:
            rows.append((label, str(value)))
    return rows


def eml_to_text(eml_path, target_path):
    """E-Mail als lesbare Textdatei (Kopf + Textkoerper + Anhangliste)."""
    message = _load_eml(eml_path)
    lines = [f"{label}: {value}" for label, value in _eml_header_rows(message)]
    lines.append("-" * 60)
    body = message.get_body(preferencelist=("plain", "html"))
    if body is not None:
        content = body.get_content()
        if body.get_content_type() == "text/html":
            content = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", content, flags=re.S | re.I)
            content = re.sub(r"<br\s*/?>|</p>|</div>|</tr>", "\n", content, flags=re.I)
            content = re.sub(r"<[^>]+>", "", content)
            content = html_unescape(content)
        lines.append(content.strip())
    attachments = [part.get_filename() or "anhang" for part in message.iter_attachments()]
    if attachments:
        lines.append("")
        lines.append("Anhaenge: " + ", ".join(attachments))
    Path(target_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def eml_to_html(eml_path, target_path):
    """E-Mail als eigenstaendige HTML-Datei - eingebettete Bilder inklusive."""
    from html import escape as html_escape

    message = _load_eml(eml_path)
    header_html = "".join(
        f"<tr><th>{html_escape(label)}</th><td>{html_escape(value)}</td></tr>"
        for label, value in _eml_header_rows(message)
    )

    body = message.get_body(preferencelist=("html", "plain"))
    if body is None:
        body_html = "<p><i>(kein Nachrichtentext)</i></p>"
    elif body.get_content_type() == "text/html":
        body_html = body.get_content()
        # cid:-Bilder als data-URIs einbetten, damit die Datei allein lebensfaehig ist
        cid_map = {}
        for part in message.walk():
            cid = part.get("Content-ID")
            if cid and part.get_content_maintype() == "image":
                payload = part.get_payload(decode=True)
                if payload:
                    uri = f"data:{part.get_content_type()};base64,{base64.b64encode(payload).decode('ascii')}"
                    cid_map[cid.strip('<>')] = uri
        for cid, uri in cid_map.items():
            body_html = body_html.replace(f"cid:{cid}", uri)
        body_html = re.sub(r"<(script)[^>]*>.*?</\1>", "", body_html, flags=re.S | re.I)
    else:
        body_html = f"<pre>{html_escape(body.get_content())}</pre>"

    attachments = [part.get_filename() or "anhang" for part in message.iter_attachments()]
    attachment_html = ""
    if attachments:
        items = "".join(f"<li>{html_escape(name)}</li>" for name in attachments)
        attachment_html = f"<h2>Anhänge ({len(attachments)})</h2><ul>{items}</ul>"

    subject = html_escape(str(message.get("Subject", Path(eml_path).stem)))
    Path(target_path).write_text(f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>{subject}</title>
<style>
body {{ font-family: "Segoe UI", system-ui, sans-serif; margin: 2rem auto; max-width: 60rem;
       color: #1c2430; background: #f6f8fb; }}
.kopf {{ background: #fff; border: 1px solid #dde3ec; border-radius: 6px; padding: 1rem 1.2rem; }}
.kopf table {{ border-collapse: collapse; }}
.kopf th {{ text-align: left; padding: .15rem .9rem .15rem 0; color: #5b6676; font-weight: 600; }}
.kopf td {{ padding: .15rem 0; }}
.inhalt {{ background: #fff; border: 1px solid #dde3ec; border-radius: 6px;
          padding: 1rem 1.2rem; margin-top: 1rem; overflow-x: auto; }}
h2 {{ font-size: 1rem; }} pre {{ white-space: pre-wrap; }}
</style>
</head>
<body>
<div class="kopf"><table>{header_html}</table></div>
<div class="inhalt">{body_html}{attachment_html}</div>
</body>
</html>
""", encoding="utf-8")


# ---------------------------------------------------------------------------
# Text-Encoding & Zeilenenden
# ---------------------------------------------------------------------------

def convert_text_file(source, target, target_encoding="utf-8", newline_mode="keep"):
    """Datei mit neuem Encoding/Zeilenende schreiben (Quelle bleibt unveraendert)."""
    text, detected = decode_bytes(Path(source).read_bytes())
    if newline_mode != "keep":
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if newline_mode == "crlf":
            text = text.replace("\n", "\r\n")
    Path(target).write_bytes(text.encode(target_encoding, errors="replace"))
    return detected


# ===========================================================================
# CAD: Mesh-Formate, HTML-3D-Viewer, STEP/IGES-Bericht, DXF -> SVG
# ===========================================================================

"""CAD-Werkzeuge auf Basis der Standardbibliothek.

1. Mesh-Konvertierung zwischen Neutralformaten:
   STL (ASCII/binär), OBJ, PLY (ASCII/binär LE), 3MF  ->  STL/OBJ/PLY/3MF/GLB
   plus eigenständige HTML-3D-Ansicht (eigener WebGL-Viewer, offline).
2. STEP -> Mesh über den eigenen B-Rep-Tessellierungs-Kern (step_mesh.py).
3. STEP/IGES-Prüfbericht: Header, Schema, Produkte, Einheiten, Entitäten.
"""


MESH_READ_EXTS = [".stl", ".obj", ".ply", ".3mf", ".step", ".stp", ".iges", ".igs"]
MESH_WRITE_EXTS = [".stl", ".obj", ".ply", ".3mf", ".glb", ".html"]
BREP_EXTS = [".step", ".stp", ".iges", ".igs"]


@dataclass
class Mesh:
    vertices: list = field(default_factory=list)   # [(x, y, z), ...]
    faces: list = field(default_factory=list)      # [(i, j, k), ...] 0-basiert
    name: str = "Teil"

    def stats(self):
        return f"{len(self.vertices)} Punkte, {len(self.faces)} Dreiecke"


# ---------------------------------------------------------------------------
# Mesh lesen
# ---------------------------------------------------------------------------

def read_mesh(path, quality="mittel", log=None):
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".stl":
        return _read_stl(path)
    if ext == ".obj":
        return _read_obj(path)
    if ext == ".ply":
        return _read_ply(path)
    if ext == ".3mf":
        return _read_3mf(path)
    if ext in (".step", ".stp"):
        return read_step_mesh(path, quality=quality, log=log)
    if ext in (".iges", ".igs"):
        return read_iges_mesh(path, quality=quality, log=log)
    raise ValueError(f"Kein unterstütztes Mesh-Format: {ext}")


class _VertexPool:
    """Punkte deduplizieren, damit OBJ/3MF/GLB kompakt werden."""

    def __init__(self):
        self.index = {}
        self.vertices = []

    def add(self, x, y, z):
        key = (round(x, 6), round(y, 6), round(z, 6))
        idx = self.index.get(key)
        if idx is None:
            idx = len(self.vertices)
            self.index[key] = idx
            self.vertices.append((x, y, z))
        return idx


def _read_stl(path):
    data = path.read_bytes()
    if len(data) < 84:
        head = data.lstrip()[:80].lower()
        if not head.startswith(b"solid"):
            raise ValueError("STL-Datei ist zu kurz/beschädigt.")
    is_ascii = data.lstrip()[:5].lower() == b"solid" and b"facet" in data[:2048].lower()
    if not is_ascii:
        count = struct.unpack_from("<I", data, 80)[0]
        if 84 + count * 50 <= len(data):
            return _read_stl_binary(data, count, path.stem)
    return _read_stl_ascii(data, path.stem)


def _read_stl_binary(data, count, name):
    pool = _VertexPool()
    faces = []
    offset = 84
    for _ in range(count):
        values = struct.unpack_from("<12f", data, offset)
        offset += 50
        a = pool.add(*values[3:6])
        b = pool.add(*values[6:9])
        c = pool.add(*values[9:12])
        faces.append((a, b, c))
    return Mesh(pool.vertices, faces, name)


def _read_stl_ascii(data, name):
    text = data.decode("ascii", errors="replace")
    pool = _VertexPool()
    faces = []
    corner = []
    for match in re.finditer(r"vertex\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)", text):
        corner.append(pool.add(float(match.group(1)), float(match.group(2)), float(match.group(3))))
        if len(corner) == 3:
            faces.append(tuple(corner))
            corner = []
    if not faces:
        raise ValueError("Keine Dreiecke in der STL-Datei gefunden.")
    return Mesh(pool.vertices, faces, name)


def _read_obj(path):
    vertices = []
    faces = []
    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        parts = raw.split()
        if not parts:
            continue
        if parts[0] == "v" and len(parts) >= 4:
            vertices.append((float(parts[1]), float(parts[2]), float(parts[3])))
        elif parts[0] == "f" and len(parts) >= 4:
            idx = []
            for token in parts[1:]:
                value = int(token.split("/")[0])
                idx.append(value - 1 if value > 0 else len(vertices) + value)
            for k in range(1, len(idx) - 1):      # Polygone als Fächer triangulieren
                faces.append((idx[0], idx[k], idx[k + 1]))
    if not faces:
        raise ValueError("Keine Flächen in der OBJ-Datei gefunden.")
    return Mesh(vertices, faces, path.stem)


_PLY_TYPES = {
    "char": ("b", 1), "int8": ("b", 1), "uchar": ("B", 1), "uint8": ("B", 1),
    "short": ("h", 2), "int16": ("h", 2), "ushort": ("H", 2), "uint16": ("H", 2),
    "int": ("i", 4), "int32": ("i", 4), "uint": ("I", 4), "uint32": ("I", 4),
    "float": ("f", 4), "float32": ("f", 4), "double": ("d", 8), "float64": ("d", 8),
}


def _read_ply(path):
    data = path.read_bytes()
    end = data.find(b"end_header")
    if end == -1:
        raise ValueError("PLY-Header nicht gefunden.")
    header = data[:end].decode("ascii", errors="replace").splitlines()
    body_start = data.find(b"\n", end) + 1

    fmt = "ascii"
    elements = []       # (name, count, [(prop_name, type, list_count_type)])
    for line in header:
        parts = line.split()
        if not parts:
            continue
        if parts[0] == "format":
            fmt = parts[1]
        elif parts[0] == "element":
            elements.append((parts[1], int(parts[2]), []))
        elif parts[0] == "property" and elements:
            if parts[1] == "list":
                elements[-1][2].append((parts[4], parts[3], parts[2]))
            else:
                elements[-1][2].append((parts[2], parts[1], None))

    if fmt == "ascii":
        return _read_ply_ascii(data[body_start:].decode("ascii", errors="replace"), elements, path.stem)
    if fmt == "binary_little_endian":
        return _read_ply_binary(data, body_start, elements, path.stem)
    raise ValueError(f"PLY-Format „{fmt}“ wird nicht unterstützt (nur ascii/binary_little_endian).")


def _read_ply_ascii(text, elements, name):
    tokens = text.split()
    pos = 0
    vertices, faces = [], []
    for elem_name, count, props in elements:
        for _ in range(count):
            if elem_name == "vertex":
                values = {}
                for prop_name, _ptype, list_type in props:
                    if list_type:
                        n = int(tokens[pos]); pos += 1 + n
                    else:
                        values[prop_name] = float(tokens[pos]); pos += 1
                vertices.append((values.get("x", 0.0), values.get("y", 0.0), values.get("z", 0.0)))
            else:
                for prop_name, _ptype, list_type in props:
                    if list_type:
                        n = int(tokens[pos]); pos += 1
                        idx = [int(tokens[pos + i]) for i in range(n)]
                        pos += n
                        if elem_name == "face" and prop_name.startswith("vertex"):
                            for k in range(1, len(idx) - 1):
                                faces.append((idx[0], idx[k], idx[k + 1]))
                    else:
                        pos += 1
    return Mesh(vertices, faces, name)


def _read_ply_binary(data, pos, elements, name):
    vertices, faces = [], []
    for elem_name, count, props in elements:
        for _ in range(count):
            values = {}
            for prop_name, ptype, list_type in props:
                if list_type:
                    cfmt, csize = _PLY_TYPES[list_type]
                    n = struct.unpack_from("<" + cfmt, data, pos)[0]
                    pos += csize
                    ifmt, isize = _PLY_TYPES[ptype]
                    idx = struct.unpack_from(f"<{n}{ifmt}", data, pos)
                    pos += n * isize
                    if elem_name == "face" and prop_name.startswith("vertex"):
                        for k in range(1, len(idx) - 1):
                            faces.append((int(idx[0]), int(idx[k]), int(idx[k + 1])))
                else:
                    vfmt, vsize = _PLY_TYPES[ptype]
                    values[prop_name] = struct.unpack_from("<" + vfmt, data, pos)[0]
                    pos += vsize
            if elem_name == "vertex":
                vertices.append((float(values.get("x", 0.0)), float(values.get("y", 0.0)),
                                 float(values.get("z", 0.0))))
    return Mesh(vertices, faces, name)


def _read_3mf(path):
    with zipfile.ZipFile(path) as archive:
        model_name = "3D/3dmodel.model"
        try:
            rels = ET.fromstring(archive.read("_rels/.rels"))
            for rel in rels:
                if rel.get("Type", "").endswith("3dmodel"):
                    model_name = rel.get("Target", model_name).lstrip("/")
        except KeyError:
            pass
        root = ET.fromstring(archive.read(model_name))

    vertices, faces = [], []
    for mesh_el in root.iter():
        if not mesh_el.tag.endswith("}mesh") and mesh_el.tag != "mesh":
            continue
        offset = len(vertices)
        for node in mesh_el.iter():
            tag = node.tag.rsplit("}", 1)[-1]
            if tag == "vertex":
                vertices.append((float(node.get("x", 0)), float(node.get("y", 0)), float(node.get("z", 0))))
            elif tag == "triangle":
                faces.append((offset + int(node.get("v1")), offset + int(node.get("v2")),
                              offset + int(node.get("v3"))))
    if not faces:
        raise ValueError("Kein Mesh in der 3MF-Datei gefunden.")
    return Mesh(vertices, faces, path.stem)


# ---------------------------------------------------------------------------
# Mesh schreiben
# ---------------------------------------------------------------------------

def write_mesh(mesh, path, stl_binary=True):
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".stl":
        _write_stl_binary(mesh, path) if stl_binary else _write_stl_ascii(mesh, path)
    elif ext == ".obj":
        _write_obj(mesh, path)
    elif ext == ".ply":
        _write_ply(mesh, path)
    elif ext == ".3mf":
        _write_3mf(mesh, path)
    elif ext == ".glb":
        _write_glb(mesh, path)
    elif ext == ".html":
        _write_html_viewer(mesh, path)
    else:
        raise ValueError(f"Kein unterstütztes Mesh-Zielformat: {ext}")


def _face_normal(mesh, face):
    (ax, ay, az), (bx, by, bz), (cx, cy, cz) = (mesh.vertices[i] for i in face)
    ux, uy, uz = bx - ax, by - ay, bz - az
    vx, vy, vz = cx - ax, cy - ay, cz - az
    nx, ny, nz = uy * vz - uz * vy, uz * vx - ux * vz, ux * vy - uy * vx
    length = (nx * nx + ny * ny + nz * nz) ** 0.5 or 1.0
    return nx / length, ny / length, nz / length


def _write_stl_binary(mesh, path):
    with path.open("wb") as handle:
        handle.write(f"{mesh.name[:70]} (Einkauf Data Converter)".encode("ascii", "replace").ljust(80, b" "))
        handle.write(struct.pack("<I", len(mesh.faces)))
        for face in mesh.faces:
            normal = _face_normal(mesh, face)
            coords = []
            for i in face:
                coords.extend(mesh.vertices[i])
            handle.write(struct.pack("<12fH", *normal, *coords, 0))


def _write_stl_ascii(mesh, path):
    lines = [f"solid {mesh.name}"]
    for face in mesh.faces:
        nx, ny, nz = _face_normal(mesh, face)
        lines.append(f"  facet normal {nx:.6e} {ny:.6e} {nz:.6e}")
        lines.append("    outer loop")
        for i in face:
            x, y, z = mesh.vertices[i]
            lines.append(f"      vertex {x:.6e} {y:.6e} {z:.6e}")
        lines.append("    endloop")
        lines.append("  endfacet")
    lines.append(f"endsolid {mesh.name}")
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


def _write_obj(mesh, path):
    lines = [f"# {mesh.name} - Einkauf Data Converter", f"o {mesh.name}"]
    for x, y, z in mesh.vertices:
        lines.append(f"v {x:.6f} {y:.6f} {z:.6f}")
    for a, b, c in mesh.faces:
        lines.append(f"f {a + 1} {b + 1} {c + 1}")
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


def _write_ply(mesh, path):
    lines = [
        "ply", "format ascii 1.0", f"comment {mesh.name} - Einkauf Data Converter",
        f"element vertex {len(mesh.vertices)}",
        "property float x", "property float y", "property float z",
        f"element face {len(mesh.faces)}",
        "property list uchar int vertex_indices", "end_header",
    ]
    for x, y, z in mesh.vertices:
        lines.append(f"{x:.6f} {y:.6f} {z:.6f}")
    for a, b, c in mesh.faces:
        lines.append(f"3 {a} {b} {c}")
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


_3MF_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="model" ContentType="application/vnd.ms-package.3dmanufacturing-3dmodel+xml"/>
</Types>"""

_3MF_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rel0" Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel" Target="/3D/3dmodel.model"/>
</Relationships>"""


def _write_3mf(mesh, path):
    parts = ['<?xml version="1.0" encoding="UTF-8"?>']
    parts.append('<model unit="millimeter" xml:lang="de-DE" '
                 'xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02">')
    parts.append(f'<metadata name="Title">{escape(mesh.name)}</metadata>')
    parts.append('<resources><object id="1" type="model"><mesh><vertices>')
    for x, y, z in mesh.vertices:
        parts.append(f'<vertex x="{x:.6f}" y="{y:.6f}" z="{z:.6f}"/>')
    parts.append("</vertices><triangles>")
    for a, b, c in mesh.faces:
        parts.append(f'<triangle v1="{a}" v2="{b}" v3="{c}"/>')
    parts.append("</triangles></mesh></object></resources>")
    parts.append('<build><item objectid="1"/></build></model>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _3MF_CONTENT_TYPES)
        archive.writestr("_rels/.rels", _3MF_RELS)
        archive.writestr("3D/3dmodel.model", "".join(parts))


def _mesh_buffers(mesh):
    positions = struct.pack(f"<{len(mesh.vertices) * 3}f",
                            *[c for v in mesh.vertices for c in v])
    indices = struct.pack(f"<{len(mesh.faces) * 3}I",
                          *[i for f in mesh.faces for i in f])
    return positions, indices


def _write_glb(mesh, path):
    positions, indices = _mesh_buffers(mesh)
    xs = [v[0] for v in mesh.vertices] or [0.0]
    ys = [v[1] for v in mesh.vertices] or [0.0]
    zs = [v[2] for v in mesh.vertices] or [0.0]
    binary = positions + indices
    if len(binary) % 4:
        binary += b"\0" * (4 - len(binary) % 4)

    gltf = {
        "asset": {"version": "2.0", "generator": "Einkauf Data Converter"},
        "scene": 0,
        "scenes": [{"nodes": [0]}],
        "nodes": [{"mesh": 0, "name": mesh.name}],
        "meshes": [{"name": mesh.name,
                    "primitives": [{"attributes": {"POSITION": 0}, "indices": 1, "material": 0}]}],
        "materials": [{"name": "Metall", "pbrMetallicRoughness":
                       {"baseColorFactor": [0.55, 0.62, 0.70, 1.0], "metallicFactor": 0.3, "roughnessFactor": 0.6}}],
        "buffers": [{"byteLength": len(binary)}],
        "bufferViews": [
            {"buffer": 0, "byteOffset": 0, "byteLength": len(positions), "target": 34962},
            {"buffer": 0, "byteOffset": len(positions), "byteLength": len(indices), "target": 34963},
        ],
        "accessors": [
            {"bufferView": 0, "componentType": 5126, "count": len(mesh.vertices), "type": "VEC3",
             "min": [min(xs), min(ys), min(zs)], "max": [max(xs), max(ys), max(zs)]},
            {"bufferView": 1, "componentType": 5125, "count": len(mesh.faces) * 3, "type": "SCALAR"},
        ],
    }
    json_bytes = json.dumps(gltf, separators=(",", ":")).encode("utf-8")
    if len(json_bytes) % 4:
        json_bytes += b" " * (4 - len(json_bytes) % 4)

    with path.open("wb") as handle:
        total = 12 + 8 + len(json_bytes) + 8 + len(binary)
        handle.write(struct.pack("<4sII", b"glTF", 2, total))
        handle.write(struct.pack("<I4s", len(json_bytes), b"JSON"))
        handle.write(json_bytes)
        handle.write(struct.pack("<I4s", len(binary), b"BIN\0"))
        handle.write(binary)


_VIEWER_TEMPLATE = """<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__ · 3D-Ansicht</title>
<style>
html, body { margin: 0; height: 100%; background: #eef1f6; font-family: "Segoe UI", system-ui, sans-serif; }
#top { position: fixed; inset: 0 0 auto 0; padding: 10px 16px; background: #16233c; color: #fff;
       display: flex; justify-content: space-between; align-items: baseline; }
#top b { font-size: 15px; } #top span { color: #a9b8d4; font-size: 12px; }
canvas { display: block; width: 100vw; height: 100vh; }
#hint { position: fixed; bottom: 10px; left: 16px; color: #5b6676; font-size: 12px; }
</style>
</head>
<body>
<div id="top"><b>__TITLE__</b><span>__STATS__ · Einkauf Data Converter</span></div>
<canvas id="c"></canvas>
<div id="hint">Ziehen = drehen · Rad = zoomen · Doppelklick = zurücksetzen</div>
<script>
"use strict";
function b64ToBuf(s){const b=atob(s),a=new Uint8Array(b.length);for(let i=0;i<b.length;i++)a[i]=b.charCodeAt(i);return a.buffer;}
const POS=new Float32Array(b64ToBuf("__POS__"));
const IDX=new Uint32Array(b64ToBuf("__IDX__"));
const canvas=document.getElementById("c");
const gl=canvas.getContext("webgl2",{antialias:true});
if(!gl){document.body.innerHTML="<p style='padding:80px 20px'>WebGL2 wird von diesem Browser nicht unterstützt.</p>";}
const VS=`#version 300 es
in vec3 p; uniform mat4 mvp, mv; out vec3 v;
void main(){ v=(mv*vec4(p,1.)).xyz; gl_Position=mvp*vec4(p,1.); }`;
const FS=`#version 300 es
precision highp float; in vec3 v; out vec4 o;
void main(){
  vec3 n=normalize(cross(dFdx(v),dFdy(v)));
  vec3 l=normalize(-v);
  float d=max(dot(n,l),0.0);
  vec3 base=vec3(0.45,0.55,0.68);
  vec3 c=base*(0.30+0.70*d)+vec3(1.0)*pow(d,48.0)*0.35;
  o=vec4(c,1.0);
}`;
function shader(type,src){const s=gl.createShader(type);gl.shaderSource(s,src);gl.compileShader(s);
  if(!gl.getShaderParameter(s,gl.COMPILE_STATUS))throw gl.getShaderInfoLog(s);return s;}
const prog=gl.createProgram();
gl.attachShader(prog,shader(gl.VERTEX_SHADER,VS));
gl.attachShader(prog,shader(gl.FRAGMENT_SHADER,FS));
gl.linkProgram(prog); gl.useProgram(prog);
const vao=gl.createVertexArray(); gl.bindVertexArray(vao);
const vb=gl.createBuffer(); gl.bindBuffer(gl.ARRAY_BUFFER,vb); gl.bufferData(gl.ARRAY_BUFFER,POS,gl.STATIC_DRAW);
const loc=gl.getAttribLocation(prog,"p"); gl.enableVertexAttribArray(loc); gl.vertexAttribPointer(loc,3,gl.FLOAT,false,0,0);
const ib=gl.createBuffer(); gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER,ib); gl.bufferData(gl.ELEMENT_ARRAY_BUFFER,IDX,gl.STATIC_DRAW);
gl.enable(gl.DEPTH_TEST);
let lo=[1e30,1e30,1e30],hi=[-1e30,-1e30,-1e30];
for(let i=0;i<POS.length;i+=3)for(let k=0;k<3;k++){lo[k]=Math.min(lo[k],POS[i+k]);hi[k]=Math.max(hi[k],POS[i+k]);}
const mid=[(lo[0]+hi[0])/2,(lo[1]+hi[1])/2,(lo[2]+hi[2])/2];
const radius=Math.max(hi[0]-lo[0],hi[1]-lo[1],hi[2]-lo[2])||1;
let theta=0.7,phi=1.05,dist=radius*2.2;
function mul(a,b){const r=new Float32Array(16);for(let i=0;i<4;i++)for(let j=0;j<4;j++){let s=0;for(let k=0;k<4;k++)s+=a[k*4+j]*b[i*4+k];r[i*4+j]=s;}return r;}
function persp(fov,ar,n,f){const t=1/Math.tan(fov/2);return new Float32Array([t/ar,0,0,0, 0,t,0,0, 0,0,(f+n)/(n-f),-1, 0,0,2*f*n/(n-f),0]);}
function view(){
  const cx=mid[0]+dist*Math.sin(phi)*Math.cos(theta),
        cy=mid[1]+dist*Math.sin(phi)*Math.sin(theta),
        cz=mid[2]+dist*Math.cos(phi);
  let zx=cx-mid[0],zy=cy-mid[1],zz=cz-mid[2];
  const zl=Math.hypot(zx,zy,zz); zx/=zl;zy/=zl;zz/=zl;
  let xx=-zy,xy=zx,xz=0; const xl=Math.hypot(xx,xy,xz)||1; xx/=xl;xy/=xl;xz/=xl;
  const yx=zy*xz-zz*xy,yy=zz*xx-zx*xz,yz=zx*xy-zy*xx;
  return new Float32Array([xx,yx,zx,0, xy,yy,zy,0, xz,yz,zz,0,
    -(xx*cx+xy*cy+xz*cz),-(yx*cx+yy*cy+yz*cz),-(zx*cx+zy*cy+zz*cz),1]);
}
function draw(){
  const w=canvas.clientWidth*devicePixelRatio,h=canvas.clientHeight*devicePixelRatio;
  if(canvas.width!==w||canvas.height!==h){canvas.width=w;canvas.height=h;gl.viewport(0,0,w,h);}
  gl.clearColor(0.933,0.945,0.965,1); gl.clear(gl.COLOR_BUFFER_BIT|gl.DEPTH_BUFFER_BIT);
  const mv=view(),p=persp(0.9,w/h,radius*0.01,radius*40);
  gl.uniformMatrix4fv(gl.getUniformLocation(prog,"mv"),false,mv);
  gl.uniformMatrix4fv(gl.getUniformLocation(prog,"mvp"),false,mul(p,mv));
  gl.drawElements(gl.TRIANGLES,IDX.length,gl.UNSIGNED_INT,0);
  requestAnimationFrame(draw);
}
let drag=null;
canvas.addEventListener("pointerdown",e=>{drag=[e.clientX,e.clientY];canvas.setPointerCapture(e.pointerId);});
canvas.addEventListener("pointermove",e=>{if(!drag)return;
  theta-=(e.clientX-drag[0])*0.008; phi-=(e.clientY-drag[1])*0.008;
  phi=Math.min(Math.max(phi,0.05),Math.PI-0.05); drag=[e.clientX,e.clientY];});
canvas.addEventListener("pointerup",()=>drag=null);
canvas.addEventListener("wheel",e=>{e.preventDefault();dist*=e.deltaY>0?1.1:0.9;
  dist=Math.min(Math.max(dist,radius*0.3),radius*20);},{passive:false});
canvas.addEventListener("dblclick",()=>{theta=0.7;phi=1.05;dist=radius*2.2;});
draw();
</script>
</body>
</html>
"""


def _write_html_viewer(mesh, path):
    positions, indices = _mesh_buffers(mesh)
    html = (_VIEWER_TEMPLATE
            .replace("__TITLE__", escape(mesh.name))
            .replace("__STATS__", mesh.stats())
            .replace("__POS__", base64.b64encode(positions).decode("ascii"))
            .replace("__IDX__", base64.b64encode(indices).decode("ascii")))
    Path(path).write_text(html, encoding="utf-8")


# ---------------------------------------------------------------------------
# STEP / IGES Prüfbericht
# ---------------------------------------------------------------------------

def step_info(path):
    text = Path(path).read_text(encoding="latin-1", errors="replace")
    info = {"Format": "STEP"}
    match = re.search(r"FILE_SCHEMA\s*\(\s*\(\s*'([^']*)'", text)
    if match:
        info["Schema"] = match.group(1).split("{")[0].strip()
    match = re.search(r"FILE_NAME\s*\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*\(([^)]*)\)\s*,\s*\(([^)]*)\)", text, re.S)
    if match:
        info["Zeitstempel"] = match.group(2)
        info["Autor"] = ", ".join(re.findall(r"'([^']*)'", match.group(3)))[:120]
        info["Organisation"] = ", ".join(re.findall(r"'([^']*)'", match.group(4)))[:120]
    products = []
    for m in re.finditer(r"=\s*PRODUCT\s*\(\s*'([^']*)'\s*,\s*'([^']*)'", text):
        label = m.group(2).strip() or m.group(1).strip()
        if label and label not in products:
            products.append(label)
    info["Produkte"] = "; ".join(products[:10]) + (" …" if len(products) > 10 else "")
    info["Anzahl Produkte"] = str(len(products))
    if re.search(r"\(\s*\.MILLI\.\s*,\s*\.METRE\.\s*\)", text):
        info["Einheit"] = "mm"
    elif re.search(r"CONVERSION_BASED_UNIT\s*\(\s*'INCH", text, re.I):
        info["Einheit"] = "inch"
    elif re.search(r"\(\s*\$\s*,\s*\.METRE\.\s*\)|SI_UNIT\s*\(\s*\$\s*,\s*\.METRE\.", text):
        info["Einheit"] = "m"
    entities = Counter(re.findall(r"#\d+\s*=\s*([A-Z0-9_]+)\s*\(", text))
    info["Entitäten"] = str(sum(entities.values()))
    info["Häufigste Entitäten"] = ", ".join(f"{name} ({count})" for name, count in entities.most_common(5))
    has_brep = any(k in entities for k in ("ADVANCED_BREP_SHAPE_REPRESENTATION", "MANIFOLD_SOLID_BREP", "CLOSED_SHELL"))
    info["Geometrie"] = "B-Rep-Volumenmodell" if has_brep else ("Flächen/Drahtmodell" if entities else "unklar")
    return info


_IGES_UNITS = {1: "inch", 2: "mm", 3: "(Datei)", 4: "ft", 5: "mi", 6: "m", 7: "km",
               8: "mil", 9: "µm", 10: "cm", 11: "µinch"}


def _iges_fields(global_text):
    """Globale IGES-Sektion: Hollerith-Strings (nH...) und Rohfelder parsen."""
    fields = []
    pos = 0
    length = len(global_text)
    while pos < length and len(fields) < 26:
        match = re.match(r"\s*(\d+)H", global_text[pos:])
        if match:
            n = int(match.group(1))
            start = pos + match.end()
            fields.append(global_text[start:start + n])
            pos = start + n
            while pos < length and global_text[pos] not in ",;":
                pos += 1
            pos += 1
        else:
            end_comma = global_text.find(",", pos)
            end_semi = global_text.find(";", pos)
            candidates = [e for e in (end_comma, end_semi) if e != -1]
            end = min(candidates) if candidates else length
            fields.append(global_text[pos:end].strip())
            pos = end + 1
    return fields


def iges_info(path):
    info = {"Format": "IGES"}
    lines = Path(path).read_text(encoding="latin-1", errors="replace").splitlines()
    global_lines = [line[:72] for line in lines if len(line) > 72 and line[72] == "G"]
    directory = [line for line in lines if len(line) > 72 and line[72] == "D"]
    try:
        fields = _iges_fields("".join(global_lines))
        info["Produkt"] = fields[2] if len(fields) > 2 else ""
        info["System"] = fields[4] if len(fields) > 4 else ""
        info["Preprozessor"] = fields[5] if len(fields) > 5 else ""
        if len(fields) > 13 and fields[13].strip().isdigit():
            info["Einheit"] = _IGES_UNITS.get(int(fields[13]), fields[14] if len(fields) > 14 else "?")
        if len(fields) > 17:
            info["Zeitstempel"] = fields[17]
        if len(fields) > 21:
            info["Autor"] = fields[20]
            info["Organisation"] = fields[21]
    except Exception:
        info["Hinweis"] = "Globale Sektion nur teilweise lesbar."
    entity_names = {100: "Kreisbogen", 102: "Kurvenzug", 104: "Kegelschnitt", 106: "Punktfolge",
                    108: "Ebene", 110: "Linie", 112: "Spline-Kurve", 114: "Spline-Fläche",
                    116: "Punkt", 118: "Regelfläche", 120: "Rotationsfläche", 122: "Translationsfläche",
                    124: "Transformation", 126: "NURBS-Kurve", 128: "NURBS-Fläche", 140: "Offset-Fläche",
                    142: "Kurve auf Fläche", 144: "Getrimmte Fläche", 186: "B-Rep-Solid",
                    308: "Subfigur", 314: "Farbe", 402: "Gruppe", 406: "Eigenschaft", 408: "Subfigur-Instanz"}
    counter = Counter()
    for index in range(0, len(directory) - 1, 2):
        raw = directory[index][:8].strip()
        if raw.isdigit():
            code = int(raw)
            counter[entity_names.get(code, f"Typ {code}")] += 1
    info["Entitäten"] = str(sum(counter.values()))
    info["Häufigste Entitäten"] = ", ".join(f"{name} ({count})" for name, count in counter.most_common(5))
    info["Geometrie"] = "B-Rep-Solid" if "B-Rep-Solid" in counter else \
        ("Flächenmodell" if any("Fläche" in k for k in counter) else "Kurven/Draht")
    return info


def brep_info(path):
    ext = Path(path).suffix.lower()
    if ext in (".step", ".stp"):
        return step_info(path)
    if ext in (".iges", ".igs"):
        return iges_info(path)
    raise ValueError(f"Kein STEP/IGES: {ext}")


# ---------------------------------------------------------------------------
# DXF -> SVG (2D-Zeichnungen sichtbar machen)
# ---------------------------------------------------------------------------

def _dxf_pairs(text):
    lines = text.splitlines()
    for i in range(0, len(lines) - 1, 2):
        try:
            yield int(lines[i].strip()), lines[i + 1].strip()
        except ValueError:
            continue


def dxf_to_svg(path, target_path):
    """ASCII-DXF (LINE, CIRCLE, ARC, LWPOLYLINE, POLYLINE, TEXT) als SVG.

    Liefert die Anzahl gezeichneter Elemente. Y-Achse wird für SVG gespiegelt.
    """
    text = Path(path).read_text(encoding="latin-1", errors="replace")
    entities = []
    section = None
    current = None

    def flush():
        if current and current.get("typ"):
            entities.append(current)

    for code, value in _dxf_pairs(text):
        if code == 2 and section == "pending":
            section = value.upper()
            continue
        if code == 0:
            if value.upper() == "SECTION":
                section = "pending"
                continue
            if value.upper() == "ENDSEC":
                flush()
                current = None
                section = None
                continue
            if section == "ENTITIES":
                flush()
                current = {"typ": value.upper()}
                continue
        if current is not None and section == "ENTITIES":
            current.setdefault(code, []).append(value)
    flush()

    def fget(entity, code, index=0, default=0.0):
        try:
            return float(entity[code][index])
        except (KeyError, IndexError, ValueError):
            return default

    shapes = []
    points_bounds = []

    def note(x, y):
        points_bounds.append((x, y))

    import math as _math
    for entity in entities:
        typ = entity.get("typ")
        if typ == "LINE":
            x1, y1 = fget(entity, 10), fget(entity, 20)
            x2, y2 = fget(entity, 11), fget(entity, 21)
            shapes.append(("line", (x1, y1, x2, y2)))
            note(x1, y1); note(x2, y2)
        elif typ == "CIRCLE":
            cx, cy, r = fget(entity, 10), fget(entity, 20), fget(entity, 40)
            shapes.append(("circle", (cx, cy, r)))
            note(cx - r, cy - r); note(cx + r, cy + r)
        elif typ == "ARC":
            cx, cy, r = fget(entity, 10), fget(entity, 20), fget(entity, 40)
            a0 = _math.radians(fget(entity, 50))
            a1 = _math.radians(fget(entity, 51))
            if a1 <= a0:
                a1 += 2 * _math.pi
            shapes.append(("arc", (cx, cy, r, a0, a1)))
            note(cx - r, cy - r); note(cx + r, cy + r)
        elif typ in ("LWPOLYLINE", "POLYLINE"):
            xs = entity.get(10, [])
            ys = entity.get(20, [])
            pts = [(float(x), float(y)) for x, y in zip(xs, ys)]
            if len(pts) >= 2:
                closed = False
                try:
                    closed = int(float(entity.get(70, ["0"])[0])) & 1 == 1
                except ValueError:
                    pass
                shapes.append(("poly", (pts, closed)))
                for p in pts:
                    note(*p)
        elif typ == "VERTEX":
            if shapes and shapes[-1][0] == "poly":
                pts, closed = shapes[-1][1]
                pts.append((fget(entity, 10), fget(entity, 20)))
                note(*pts[-1])
        elif typ in ("TEXT", "MTEXT"):
            x, y = fget(entity, 10), fget(entity, 20)
            h = fget(entity, 40, default=2.5) or 2.5
            content = (entity.get(1, [""])[0]).replace("\\P", " ")
            content = re.sub(r"\[A-Za-z][^;]*;", "", content).replace("{", "").replace("}", "")
            if content:
                shapes.append(("text", (x, y, h, content)))
                note(x, y); note(x + h * 0.7 * len(content), y + h)

    if not shapes:
        raise ValueError("Keine darstellbaren 2D-Elemente in der DXF-Datei gefunden.")

    min_x = min(p[0] for p in points_bounds)
    max_x = max(p[0] for p in points_bounds)
    min_y = min(p[1] for p in points_bounds)
    max_y = max(p[1] for p in points_bounds)
    width = max(max_x - min_x, 1e-6)
    height = max(max_y - min_y, 1e-6)
    margin = 0.03 * max(width, height)
    stroke = max(width, height) / 400

    def ty(y):                                        # DXF-Y nach SVG-Y spiegeln
        return (max_y + min_y) - y

    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" '
             f'viewBox="{min_x - margin:.3f} {min_y - margin:.3f} '
             f'{width + 2 * margin:.3f} {height + 2 * margin:.3f}">',
             f'<g fill="none" stroke="#1c2430" stroke-width="{stroke:.4f}" '
             f'stroke-linecap="round" stroke-linejoin="round">']
    for kind, payload in shapes:
        if kind == "line":
            x1, y1, x2, y2 = payload
            parts.append(f'<line x1="{x1:.3f}" y1="{ty(y1):.3f}" x2="{x2:.3f}" y2="{ty(y2):.3f}"/>')
        elif kind == "circle":
            cx, cy, r = payload
            parts.append(f'<circle cx="{cx:.3f}" cy="{ty(cy):.3f}" r="{r:.3f}"/>')
        elif kind == "arc":
            cx, cy, r, a0, a1 = payload
            x1 = cx + r * _math.cos(a0); y1 = cy + r * _math.sin(a0)
            x2 = cx + r * _math.cos(a1); y2 = cy + r * _math.sin(a1)
            large = 1 if (a1 - a0) > _math.pi else 0
            parts.append(f'<path d="M {x1:.3f} {ty(y1):.3f} '
                         f'A {r:.3f} {r:.3f} 0 {large} 0 {x2:.3f} {ty(y2):.3f}"/>')
        elif kind == "poly":
            pts, closed = payload
            coords = " ".join(f"{x:.3f},{ty(y):.3f}" for x, y in pts)
            tag = "polygon" if closed else "polyline"
            parts.append(f'<{tag} points="{coords}"/>')
        elif kind == "text":
            x, y, h, content = payload
            parts.append(f'<text x="{x:.3f}" y="{ty(y):.3f}" font-size="{h:.3f}" '
                         f'fill="#1c2430" stroke="none" '
                         f'font-family="Segoe UI, sans-serif">{escape(content)}</text>')
    parts.append("</g></svg>")
    Path(target_path).write_text("\n".join(parts), encoding="utf-8")
    return len(shapes)


# ===========================================================================
# STEP-B-Rep-Tessellierungs-Kern
# ===========================================================================

"""STEP-B-Rep-Tessellierungs-Kern - reine Standardbibliothek.

Wandelt STEP-Dateien (ISO 10303-21, AP203/AP214/AP242) in Dreiecksnetze:

    Part-21-Parser  ->  Topologie (Shell/Face/Loop/Edge)  ->  Flächen-Evaluatoren
    (Ebene, Zylinder, Kegel, Kugel, Torus, NURBS, Extrusion, Rotation)  ->
    Kanten-Diskretisierung  ->  Trimmung im UV-Raum  ->  Ear-Clipping +
    Verfeinerung  ->  Mesh

Grenzen (bewusst): Baugruppen-Transformationen werden nicht angewendet
(Einzelteile sind exakt, Baugruppen liegen ggf. übereinander); Trimmkurven auf
voll umlaufenden Flächen (z. B. Bohrung quer durch eine Zylinderwand) werden
als Band angenähert. Ergebnis ist ein Sichtmodell/Fertigungs-Mesh, kein exaktes
B-Rep.
"""



TWO_PI = 2.0 * math.pi
QUALITY_SEGMENTS = {"grob": 24, "mittel": 48, "fein": 96}

_FACE_TYPES = {"ADVANCED_FACE", "FACE_SURFACE"}
_MAX_FACE_TRIS = 40000
_MAX_TOTAL_TRIS = 800000


# ---------------------------------------------------------------------------
# Vektor-Helfer
# ---------------------------------------------------------------------------

def _add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _mul(a, s):
    return (a[0] * s, a[1] * s, a[2] * s)


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])


def _norm(a):
    length = math.sqrt(_dot(a, a))
    return (a[0] / length, a[1] / length, a[2] / length) if length else (0.0, 0.0, 1.0)


def _any_perp(z):
    axis = (1.0, 0.0, 0.0) if abs(z[0]) <= min(abs(z[1]), abs(z[2])) else \
           (0.0, 1.0, 0.0) if abs(z[1]) <= abs(z[2]) else (0.0, 0.0, 1.0)
    return _norm(_cross(z, axis))


# ---------------------------------------------------------------------------
# Part-21-Parser
# ---------------------------------------------------------------------------

class Ref(int):
    """Entity-Verweis #123."""


def _split_statements(data):
    statements = []
    buffer = []
    in_string = False
    for ch in data:
        if ch == "'":
            in_string = not in_string
            buffer.append(ch)
        elif ch == ";" and not in_string:
            statements.append("".join(buffer))
            buffer = []
        else:
            buffer.append(ch)
    return statements


_NUM_RE = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")


def _parse_value(text, pos):
    while pos < len(text) and text[pos] in " \t\r\n":
        pos += 1
    if pos >= len(text):
        return None, pos
    ch = text[pos]
    if ch == "(":
        values = []
        pos += 1
        while True:
            while pos < len(text) and text[pos] in " \t\r\n":
                pos += 1
            if pos >= len(text) or text[pos] == ")":
                return values, pos + 1
            value, pos = _parse_value(text, pos)
            values.append(value)
            while pos < len(text) and text[pos] in " \t\r\n":
                pos += 1
            if pos < len(text) and text[pos] == ",":
                pos += 1
    if ch == "#":
        match = re.match(r"#(\d+)", text[pos:])
        return Ref(match.group(1)), pos + match.end()
    if ch == "'":
        end = pos + 1
        while end < len(text):
            if text[end] == "'":
                if end + 1 < len(text) and text[end + 1] == "'":
                    end += 2
                    continue
                break
            end += 1
        return text[pos + 1:end].replace("''", "'"), end + 1
    if ch == ".":
        match = re.match(r"\.[A-Z0-9_]+\.", text[pos:])
        if match:
            return match.group(0), pos + match.end()
    if ch == "$":
        return None, pos + 1
    if ch == "*":
        return "*", pos + 1
    match = _NUM_RE.match(text, pos)
    if match:
        raw = match.group(0)
        value = float(raw) if any(c in raw for c in ".eE") else int(raw)
        return value, match.end()
    match = re.match(r"[A-Za-z0-9_]+", text[pos:])
    if match:
        return match.group(0), pos + match.end()
    return None, pos + 1


def _parse_record(text, pos):
    match = re.match(r"\s*([A-Za-z0-9_]+)\s*", text[pos:])
    if not match:
        return None, pos
    name = match.group(1).upper()
    pos += match.end()
    if pos < len(text) and text[pos] == "(":
        args, pos = _parse_value(text, pos)
    else:
        args = []
    return (name, args), pos


class StepFile:
    def __init__(self, text):
        text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
        match = re.search(r"\bDATA\s*;(.*?)\bENDSEC\s*;", text, re.S | re.I)
        if not match:
            raise ValueError("Keine DATA-Sektion - ist das eine STEP-Datei?")
        self.entities = {}
        for statement in _split_statements(match.group(1)):
            m = re.match(r"\s*#(\d+)\s*=\s*(.*)", statement, re.S)
            if not m:
                continue
            eid = int(m.group(1))
            body = m.group(2).strip()
            if body.startswith("("):                       # komplexe Entität
                records = {}
                pos = 1
                while pos < len(body) and not body[pos:].lstrip().startswith(")"):
                    record, pos = _parse_record(body, pos)
                    if record is None:
                        break
                    records[record[0]] = record[1]
                self.entities[eid] = ("&COMPLEX", records)
            else:
                record, _pos = _parse_record(body, 0)
                if record:
                    self.entities[eid] = record

    def get(self, ref):
        return self.entities.get(int(ref), ("?", []))

    def name_of(self, ref):
        return self.get(ref)[0]

    def by_type(self, *names):
        wanted = set(names)
        result = []
        for eid, (name, args) in self.entities.items():
            if name in wanted:
                result.append(eid)
            elif name == "&COMPLEX" and wanted.intersection(args):
                result.append(eid)
        return result


def _true(flag):
    return flag == ".T."


# ---------------------------------------------------------------------------
# B-Spline (de Boor, homogen für rationale Kurven/Flächen)
# ---------------------------------------------------------------------------

def _expand_knots(mults, knots):
    expanded = []
    for mult, knot in zip(mults, knots):
        expanded.extend([float(knot)] * int(mult))
    return expanded


def _find_span(degree, knots, count, t):
    low, high = degree, count
    t = min(max(t, knots[low]), knots[high])
    for k in range(low, high):
        if knots[k] <= t < knots[k + 1]:
            return k, t
    return high - 1, t


def _deboor(degree, knots, ctrl4, t):
    """Homogene de-Boor-Auswertung. ctrl4: [(wx, wy, wz, w), ...]."""
    span, t = _find_span(degree, knots, len(ctrl4), t)
    d = [list(ctrl4[j]) for j in range(span - degree, span + 1)]
    for r in range(1, degree + 1):
        for j in range(degree, r - 1, -1):
            i = span - degree + j
            denom = knots[i + degree - r + 1] - knots[i]
            alpha = 0.0 if denom == 0 else (t - knots[i]) / denom
            for c in range(4):
                d[j][c] = d[j - 1][c] * (1 - alpha) + d[j][c] * alpha
    return d[degree]


def _homogeneous(points, weights):
    result = []
    for point, weight in zip(points, weights):
        result.append((point[0] * weight, point[1] * weight, point[2] * weight, weight))
    return result


class BSplineCurve:
    def __init__(self, degree, points, weights, knots):
        self.degree = degree
        self.ctrl = _homogeneous(points, weights)
        self.knots = knots
        self.t0, self.t1 = knots[degree], knots[len(points)]

    def ev(self, t):
        x, y, z, w = _deboor(self.degree, self.knots, self.ctrl, t)
        w = w or 1.0
        return (x / w, y / w, z / w)


class BSplineSurface:
    def __init__(self, u_degree, v_degree, grid, weights, u_knots, v_knots):
        self.u_degree, self.v_degree = u_degree, v_degree
        self.rows = [_homogeneous(row, wrow) for row, wrow in zip(grid, weights)]
        self.u_knots, self.v_knots = u_knots, v_knots
        self.u0, self.u1 = u_knots[u_degree], u_knots[len(grid)]
        self.v0, self.v1 = v_knots[v_degree], v_knots[len(grid[0])]

    def ev(self, u, v):
        column = [_deboor(self.v_degree, self.v_knots, row, v) for row in self.rows]
        x, y, z, w = _deboor(self.u_degree, self.u_knots, column, u)
        w = w or 1.0
        return (x / w, y / w, z / w)


# ---------------------------------------------------------------------------
# Flächen-Beschreibung
# ---------------------------------------------------------------------------

class Surface:
    """ev(u, v) -> 3D, uv_of(p) -> (u, v); Perioden und Krümmungsinfo."""

    def __init__(self, ev, uv_of, u_period=None, v_period=None,
                 curved_u=True, curved_v=True, kind="?", u_scale=1.0, v_scale=1.0):
        self.ev = ev
        self.uv_of = uv_of
        self.u_period = u_period
        self.v_period = v_period
        self.curved_u = curved_u
        self.curved_v = curved_v
        self.kind = kind
        # Umrechnung UV-Länge -> ungefähre 3D-Länge (für die Verfeinerung)
        self.u_scale = u_scale
        self.v_scale = v_scale


def _generic_uv_of(ev, u0, u1, v0, v1, nu=24, nv=24):
    """Inversion über Rastersuche + lokale Verfeinerung (für NURBS & Co.)."""
    du = (u1 - u0) / nu
    dv = (v1 - v0) / nv
    grid = []
    for i in range(nu + 1):
        for j in range(nv + 1):
            u = u0 + i * du
            v = v0 + j * dv
            grid.append((ev(u, v), u, v))

    def uv_of(p):
        best_u, best_v = grid[0][1], grid[0][2]
        best_d = float("inf")
        for point, u, v in grid:
            d = (point[0] - p[0]) ** 2 + (point[1] - p[1]) ** 2 + (point[2] - p[2]) ** 2
            if d < best_d:
                best_d, best_u, best_v = d, u, v
        step_u, step_v = du, dv
        for _ in range(18):
            improved = False
            for cu, cv in ((best_u + step_u, best_v), (best_u - step_u, best_v),
                           (best_u, best_v + step_v), (best_u, best_v - step_v),
                           (best_u + step_u, best_v + step_v), (best_u - step_u, best_v - step_v),
                           (best_u + step_u, best_v - step_v), (best_u - step_u, best_v + step_v)):
                cu = min(max(cu, u0), u1)
                cv = min(max(cv, v0), v1)
                point = ev(cu, cv)
                d = (point[0] - p[0]) ** 2 + (point[1] - p[1]) ** 2 + (point[2] - p[2]) ** 2
                if d < best_d:
                    best_d, best_u, best_v = d, cu, cv
                    improved = True
            if not improved:
                step_u /= 2
                step_v /= 2
        return (best_u, best_v)

    return uv_of


# ---------------------------------------------------------------------------
# Kern
# ---------------------------------------------------------------------------

class StepTessellator:
    def __init__(self, step, segments=48, log=None):
        self.step = step
        self.segments = segments
        self.log = log or (lambda message: None)
        self.pool = _VertexPool()
        self.tris = []
        self.skipped = Counter()
        self.faces_done = 0

    # -- Grundelemente -----------------------------------------------------

    def point(self, ref):
        name, args = self.step.get(ref)
        if name == "VERTEX_POINT":
            return self.point(args[1])
        return tuple(float(c) for c in args[1][:3])

    def direction(self, ref):
        _name, args = self.step.get(ref)
        return _norm(tuple(float(c) for c in args[1][:3]))

    def axis2(self, ref):
        _name, args = self.step.get(ref)
        origin = self.point(args[1]) if args[1] else (0.0, 0.0, 0.0)
        z = self.direction(args[2]) if len(args) > 2 and args[2] else (0.0, 0.0, 1.0)
        if len(args) > 3 and args[3]:
            x = self.direction(args[3])
            x = _norm(_sub(x, _mul(z, _dot(x, z))))
        else:
            x = _any_perp(z)
        return origin, x, _cross(z, x), z

    # -- Kurven ------------------------------------------------------------

    def _bspline_curve(self, name, args, records=None):
        if records is not None:
            base = records.get("B_SPLINE_CURVE", [])
            degree = int(base[0])
            point_refs = base[1]
            knot_rec = records.get("B_SPLINE_CURVE_WITH_KNOTS", [])
            mults, knot_values = knot_rec[0], knot_rec[1]
            weight_rec = records.get("RATIONAL_B_SPLINE_CURVE", [[1.0] * len(point_refs)])
            weights = [float(w) for w in weight_rec[0]]
        else:
            degree = int(args[1])
            point_refs = args[2]
            mults, knot_values = args[6], args[7]
            weights = [1.0] * len(point_refs)
        points = [self.point(r) for r in point_refs]
        knots = _expand_knots(mults, knot_values)
        return BSplineCurve(degree, points, weights, knots)

    def curve_of(self, ref):
        """(art, daten) für eine Kurven-Entität."""
        name, args = self.step.get(ref)
        if name == "&COMPLEX":
            if "B_SPLINE_CURVE" in args:
                return ("bspline", self._bspline_curve(name, None, records=args))
            return ("?", None)
        if name == "LINE":
            return ("line", None)
        if name in ("CIRCLE", "ELLIPSE"):
            placement = self.axis2(args[1])
            if name == "CIRCLE":
                a = b = float(args[2])
            else:
                a, b = float(args[2]), float(args[3])
            return ("arc", (placement, a, b))
        if name == "B_SPLINE_CURVE_WITH_KNOTS":
            return ("bspline", self._bspline_curve(name, args))
        if name == "TRIMMED_CURVE":
            return self.curve_of(args[1])
        return ("?", None)

    def edge_points(self, edge_ref):
        """Punkte einer EDGE_CURVE in ihrer natürlichen Richtung (start -> end)."""
        _name, args = self.step.get(edge_ref)
        p_start = self.point(args[1])
        p_end = self.point(args[2])
        same_sense = _true(args[4]) if len(args) > 4 else True
        kind, data = self.curve_of(args[3])

        if kind == "arc":
            (origin, x_axis, y_axis, _z), a, b = data
            def angle_of(p):
                d = _sub(p, origin)
                return math.atan2(_dot(d, y_axis) / b, _dot(d, x_axis) / a)
            t0, t1 = angle_of(p_start), angle_of(p_end)
            closed = math.dist(p_start, p_end) < 1e-9
            if closed:
                t0, t1 = 0.0, TWO_PI
            elif same_sense:
                while t1 <= t0 + 1e-12:
                    t1 += TWO_PI
            else:
                while t1 >= t0 - 1e-12:
                    t1 -= TWO_PI
            steps = max(2, int(round(self.segments * abs(t1 - t0) / TWO_PI)))
            points = []
            for i in range(steps + 1):
                t = t0 + (t1 - t0) * i / steps
                points.append(_add(origin, _add(_mul(x_axis, a * math.cos(t)),
                                                _mul(y_axis, b * math.sin(t)))))
            points[0], points[-1] = p_start, p_end
            return points

        if kind == "bspline":
            curve = data
            samples = max(16, self.segments)
            raw = [curve.ev(curve.t0 + (curve.t1 - curve.t0) * i / samples)
                   for i in range(samples + 1)]
            def nearest(p):
                return min(range(len(raw)), key=lambda i: math.dist(raw[i], p))
            i0, i1 = nearest(p_start), nearest(p_end)
            closed = math.dist(raw[0], raw[-1]) < 1e-9 and math.dist(p_start, p_end) < 1e-9
            if closed or i0 == i1:
                points = raw if same_sense else raw[::-1]
            elif i0 < i1:
                points = raw[i0:i1 + 1]
                if not same_sense:
                    points = points[::-1]
            else:
                points = raw[i1:i0 + 1][::-1]
                if not same_sense:
                    points = points[::-1]
            points = list(points)
            points[0], points[-1] = p_start, p_end
            return points

        return [p_start, p_end]

    def loop_points(self, loop_ref):
        _name, args = self.step.get(loop_ref)
        if not args or not isinstance(args[1], list):
            return []
        points = []
        for oriented_ref in args[1]:
            _oname, oargs = self.step.get(oriented_ref)
            segment = self.edge_points(oargs[3])
            if not _true(oargs[4]):
                segment = segment[::-1]
            points.extend(segment[:-1])
        return points

    # -- Flächen -----------------------------------------------------------

    def surface_of(self, ref):
        name, args = self.step.get(ref)

        if name == "&COMPLEX" and "B_SPLINE_SURFACE" in args:
            return self._make_bspline_surface(records=args)

        if name == "PLANE":
            origin, x, y, _z = self.axis2(args[1])
            def ev(u, v, o=origin, X=x, Y=y):
                return _add(o, _add(_mul(X, u), _mul(Y, v)))
            def uv_of(p, o=origin, X=x, Y=y):
                d = _sub(p, o)
                return (_dot(d, X), _dot(d, Y))
            return Surface(ev, uv_of, curved_u=False, curved_v=False, kind="plane")

        if name == "CYLINDRICAL_SURFACE":
            origin, x, y, z = self.axis2(args[1])
            radius = float(args[2])
            def ev(u, v, o=origin, X=x, Y=y, Z=z, r=radius):
                return _add(o, _add(_mul(X, r * math.cos(u)),
                                    _add(_mul(Y, r * math.sin(u)), _mul(Z, v))))
            def uv_of(p, o=origin, X=x, Y=y, Z=z):
                d = _sub(p, o)
                return (math.atan2(_dot(d, Y), _dot(d, X)) % TWO_PI, _dot(d, Z))
            return Surface(ev, uv_of, u_period=TWO_PI, curved_v=False,
                           kind="cylinder", u_scale=radius)

        if name == "CONICAL_SURFACE":
            origin, x, y, z = self.axis2(args[1])
            radius, semi_angle = float(args[2]), float(args[3])
            tan_a = math.tan(semi_angle)
            def ev(u, v, o=origin, X=x, Y=y, Z=z, r=radius, t=tan_a):
                rv = r + v * t
                return _add(o, _add(_mul(X, rv * math.cos(u)),
                                    _add(_mul(Y, rv * math.sin(u)), _mul(Z, v))))
            def uv_of(p, o=origin, X=x, Y=y, Z=z):
                d = _sub(p, o)
                return (math.atan2(_dot(d, Y), _dot(d, X)) % TWO_PI, _dot(d, Z))
            return Surface(ev, uv_of, u_period=TWO_PI, curved_v=False,
                           kind="cone", u_scale=max(radius, 1e-6))

        if name == "SPHERICAL_SURFACE":
            origin, x, y, z = self.axis2(args[1])
            radius = float(args[2])
            def ev(u, v, o=origin, X=x, Y=y, Z=z, r=radius):
                cv = math.cos(v)
                return _add(o, _add(_mul(X, r * cv * math.cos(u)),
                                    _add(_mul(Y, r * cv * math.sin(u)), _mul(Z, r * math.sin(v)))))
            def uv_of(p, o=origin, X=x, Y=y, Z=z, r=radius):
                d = _sub(p, o)
                dz = max(-1.0, min(1.0, _dot(d, Z) / (r or 1.0)))
                return (math.atan2(_dot(d, Y), _dot(d, X)) % TWO_PI, math.asin(dz))
            return Surface(ev, uv_of, u_period=TWO_PI, kind="sphere",
                           u_scale=radius, v_scale=radius)

        if name == "TOROIDAL_SURFACE":
            origin, x, y, z = self.axis2(args[1])
            major, minor = float(args[2]), float(args[3])
            def ev(u, v, o=origin, X=x, Y=y, Z=z, R=major, r=minor):
                ring = R + r * math.cos(v)
                return _add(o, _add(_mul(X, ring * math.cos(u)),
                                    _add(_mul(Y, ring * math.sin(u)), _mul(Z, r * math.sin(v)))))
            def uv_of(p, o=origin, X=x, Y=y, Z=z, R=major):
                d = _sub(p, o)
                u = math.atan2(_dot(d, Y), _dot(d, X)) % TWO_PI
                w = math.hypot(_dot(d, X), _dot(d, Y)) - R
                return (u, math.atan2(_dot(d, Z), w) % TWO_PI)
            return Surface(ev, uv_of, u_period=TWO_PI, v_period=TWO_PI,
                           kind="torus", u_scale=major + minor, v_scale=minor)

        if name == "B_SPLINE_SURFACE_WITH_KNOTS":
            return self._make_bspline_surface(args=args)

        if name == "SURFACE_OF_LINEAR_EXTRUSION":
            kind, data = self.curve_of(args[1])
            _vname, vargs = self.step.get(args[2])
            direction = _mul(self.direction(vargs[1]), float(vargs[2]))
            curve_ev, t0, t1 = self._curve_evaluator(kind, data, args[1])
            if curve_ev is None:
                return None
            def ev(u, v, ce=curve_ev, d=direction):
                return _add(ce(u), _mul(d, v))
            extent = math.dist(curve_ev(t0), curve_ev(t1)) or 1.0
            uv_of = _generic_uv_of(ev, t0, t1, -2.0, 2.0)
            return Surface(ev, uv_of, kind="extrusion", curved_v=False,
                           u_scale=extent / max(t1 - t0, 1e-9))

        if name == "SURFACE_OF_REVOLUTION":
            kind, data = self.curve_of(args[1])
            _aname, aargs = self.step.get(args[2])          # AXIS1_PLACEMENT
            axis_origin = self.point(aargs[1])
            axis_dir = self.direction(aargs[2]) if len(aargs) > 2 and aargs[2] else (0.0, 0.0, 1.0)
            curve_ev, t0, t1 = self._curve_evaluator(kind, data, args[1])
            if curve_ev is None:
                return None
            def ev(u, v, ce=curve_ev, o=axis_origin, z=axis_dir):
                p = _sub(ce(v), o)
                par = _mul(z, _dot(p, z))
                perp = _sub(p, par)
                ortho = _cross(z, perp)
                cos_u, sin_u = math.cos(u), math.sin(u)
                rotated = _add(_add(_mul(perp, cos_u), _mul(ortho, sin_u)), par)
                return _add(o, rotated)
            radius = max(math.dist(curve_ev((t0 + t1) / 2), axis_origin), 1e-6)
            uv_of = _generic_uv_of(ev, 0.0, TWO_PI, t0, t1)
            return Surface(ev, uv_of, u_period=TWO_PI, kind="revolution", u_scale=radius)

        return None

    def _curve_evaluator(self, kind, data, ref):
        if kind == "bspline":
            return data.ev, data.t0, data.t1
        if kind == "arc":
            (origin, x_axis, y_axis, _z), a, b = data
            def ev(t, o=origin, X=x_axis, Y=y_axis, A=a, B=b):
                return _add(o, _add(_mul(X, A * math.cos(t)), _mul(Y, B * math.sin(t))))
            return ev, 0.0, TWO_PI
        if kind == "line":
            _name, args = self.step.get(ref)
            origin = self.point(args[1])
            _vname, vargs = self.step.get(args[2])
            direction = _mul(self.direction(vargs[1]), float(vargs[2]))
            def ev(t, o=origin, d=direction):
                return _add(o, _mul(d, t))
            return ev, -1.0, 1.0
        return None, 0.0, 1.0

    def _make_bspline_surface(self, args=None, records=None):
        if records is not None:
            base = records.get("B_SPLINE_SURFACE", [])
            u_degree, v_degree = int(base[0]), int(base[1])
            grid_refs = base[2]
            knot_rec = records.get("B_SPLINE_SURFACE_WITH_KNOTS", [])
            u_mults, v_mults, u_knot_values, v_knot_values = knot_rec[0], knot_rec[1], knot_rec[2], knot_rec[3]
            weight_rec = records.get("RATIONAL_B_SPLINE_SURFACE")
            weights = ([[float(w) for w in row] for row in weight_rec[0]] if weight_rec
                       else [[1.0] * len(grid_refs[0]) for _ in grid_refs])
        else:
            u_degree, v_degree = int(args[1]), int(args[2])
            grid_refs = args[3]
            u_mults, v_mults, u_knot_values, v_knot_values = args[8], args[9], args[10], args[11]
            weights = [[1.0] * len(grid_refs[0]) for _ in grid_refs]
        grid = [[self.point(r) for r in row] for row in grid_refs]
        surface = BSplineSurface(u_degree, v_degree, grid, weights,
                                 _expand_knots(u_mults, u_knot_values),
                                 _expand_knots(v_mults, v_knot_values))
        corner_a = surface.ev(surface.u0, surface.v0)
        corner_b = surface.ev(surface.u1, surface.v1)
        extent = max(math.dist(corner_a, corner_b), 1e-6)
        uv_of = _generic_uv_of(surface.ev, surface.u0, surface.u1, surface.v0, surface.v1)
        return Surface(surface.ev, uv_of, kind="nurbs",
                       u_scale=extent / max(surface.u1 - surface.u0, 1e-9),
                       v_scale=extent / max(surface.v1 - surface.v0, 1e-9))

    # -- Trimmung & Triangulierung ----------------------------------------

    def tessellate_face(self, face_ref):
        name, args = self.step.get(face_ref)
        if name == "&COMPLEX":
            for face_type in _FACE_TYPES:
                if face_type in args:
                    args = args[face_type]
                    break
        bounds, surface_ref, same_sense = args[1], args[2], _true(args[3])
        surface = self.surface_of(surface_ref)
        if surface is None:
            self.skipped[self.step.name_of(surface_ref)] += 1
            return

        loops = []
        for bound_ref in bounds:
            _bname, bargs = self.step.get(bound_ref)
            points = self.loop_points(bargs[1])
            if len(points) >= 3:
                if not _true(bargs[2]):
                    points = points[::-1]
                loops.append(points)
        if not loops:
            return

        triangles_uv = self._face_uv_triangles(surface, loops)
        if not triangles_uv:
            return
        emitted = 0
        for a, b, c in triangles_uv:
            if len(self.tris) >= _MAX_TOTAL_TRIS:
                break
            pa = self.pool.add(*surface.ev(*a))
            pb = self.pool.add(*surface.ev(*b))
            pc = self.pool.add(*surface.ev(*c))
            if pa == pb or pb == pc or pa == pc:
                continue
            self.tris.append((pa, pb, pc) if same_sense else (pa, pc, pb))
            emitted += 1
        if emitted:
            self.faces_done += 1

    def _face_uv_triangles(self, surface, loops):
        uv_loops = []
        for loop in loops:
            uvs = [surface.uv_of(p) for p in loop]
            uvs, _winding = _unwrap_loop(uvs, surface.u_period, surface.v_period)
            uv_loops.append(uvs)
        return face_triangles_from_uv(surface, uv_loops, self.segments)

    # -- Gesamtablauf ------------------------------------------------------

    def run(self, name):
        face_refs = self.step.by_type(*_FACE_TYPES)
        if not face_refs:
            raise ValueError("Kein B-Rep-Inhalt (keine ADVANCED_FACE-Entitäten) gefunden.")
        for ref in face_refs:
            try:
                self.tessellate_face(Ref(ref))
            except Exception:
                self.skipped["Fehler bei Fläche"] += 1
        if not self.tris:
            raise ValueError("Keine Fläche konnte tesselliert werden.")
        if self.skipped:
            details = ", ".join(f"{k} ({v})" for k, v in self.skipped.most_common())
            self.log(f"Übersprungen: {details}")
        self.log(f"Tesselliert: {self.faces_done}/{len(face_refs)} Flächen, "
                 f"{len(self.tris)} Dreiecke.")
        return Mesh(self.pool.vertices, self.tris, name)


# ---------------------------------------------------------------------------
# UV-Geometrie: Entrollen, Fläche, Triangulierung, Verfeinerung
# (Modulebene, damit auch der IGES-Kern sie nutzen kann)
# ---------------------------------------------------------------------------

def face_triangles_from_uv(surface, uv_loops, segments):
    """Getrimmte Fläche aus (bereits entrollten) UV-Loops triangulieren."""
    wrapped = False
    for loop in uv_loops:
        _re, winding = _unwrap_loop(loop, surface.u_period, surface.v_period)
        if winding != (0, 0):
            wrapped = True
    if wrapped or not uv_loops:
        return grid_triangles(surface, uv_loops, segments)
    target = _refine_target(surface, segments)
    outer = max(uv_loops, key=lambda l: abs(_polygon_area(l)))
    holes = [l for l in uv_loops if l is not outer]
    if _polygon_area(outer) < 0:
        outer = outer[::-1]
    holes = [h[::-1] if _polygon_area(h) > 0 else h for h in holes]
    triangles = _triangulate_polygon(outer, holes)
    if target:
        triangles = _refine_triangles(triangles, target, _MAX_FACE_TRIS)
    return triangles


def _refine_target(surface, segments):
    if not surface.curved_u and not surface.curved_v:
        return None
    chord = TWO_PI / segments
    target_u = chord if surface.curved_u else float("inf")
    target_v = chord if surface.curved_v else float("inf")
    if surface.kind == "nurbs":
        # NURBS-Domäne ist nicht in Radiant: Ziel = Domänenanteil
        target_u = (surface.u_scale and 1.0 / (segments / 4)) or target_u
        target_v = (surface.v_scale and 1.0 / (segments / 4)) or target_v
    return (target_u, target_v)


def grid_triangles(surface, uv_loops, segments, domain=None):
    """Voll umlaufende Flächen (Zylinderwand, Torus, Kugelzone) als Band."""
    all_u = [uv[0] for loop in uv_loops for uv in loop]
    all_v = [uv[1] for loop in uv_loops for uv in loop]
    if domain and not all_u:
        u_lo, u_hi, v_lo, v_hi = domain
    else:
        if surface.u_period:
            u_lo, u_hi = 0.0, surface.u_period
        else:
            u_lo, u_hi = min(all_u), max(all_u)
        if surface.v_period and (max(all_v) - min(all_v)) > surface.v_period * 0.98:
            v_lo, v_hi = 0.0, surface.v_period
        else:
            v_lo, v_hi = min(all_v), max(all_v)
    if u_hi - u_lo < 1e-12 or v_hi - v_lo < 1e-12:
        return []
    if surface.u_period:
        n_u = max(4, int(round(segments * (u_hi - u_lo) / TWO_PI)))
    elif surface.curved_u:
        n_u = max(8, segments // 4)
    else:
        n_u = max(2, segments // 8)
    if surface.curved_v:
        span = (v_hi - v_lo) / (surface.v_period or TWO_PI)
        n_v = max(4, int(round(segments * span))) if surface.v_period else max(8, segments // 4)
    else:
        n_v = 1
    triangles = []
    for i in range(n_u):
        u_a = u_lo + (u_hi - u_lo) * i / n_u
        u_b = u_lo + (u_hi - u_lo) * (i + 1) / n_u
        for j in range(n_v):
            v_a = v_lo + (v_hi - v_lo) * j / n_v
            v_b = v_lo + (v_hi - v_lo) * (j + 1) / n_v
            triangles.append(((u_a, v_a), (u_b, v_a), (u_b, v_b)))
            triangles.append(((u_a, v_a), (u_b, v_b), (u_a, v_b)))
    return triangles

def _unwrap_loop(uvs, u_period, v_period):
    """Periodische Koordinaten stetig machen; liefert (Loop, Windungszahlen)."""
    result = [uvs[0]]
    for u, v in uvs[1:]:
        pu, pv = result[-1]
        if u_period:
            delta = u - pu
            u = pu + delta - u_period * round(delta / u_period)
        if v_period:
            delta = v - pv
            v = pv + delta - v_period * round(delta / v_period)
        result.append((u, v))
    wind_u = wind_v = 0
    if u_period:
        delta = uvs[0][0] - result[-1][0]
        delta -= u_period * round(delta / u_period)
        wind_u = round((result[-1][0] + delta - result[0][0]) / u_period)
    if v_period:
        delta = uvs[0][1] - result[-1][1]
        delta -= v_period * round(delta / v_period)
        wind_v = round((result[-1][1] + delta - result[0][1]) / v_period)
    return result, (wind_u, wind_v)


def _polygon_area(points):
    total = 0.0
    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        total += x1 * y2 - x2 * y1
    return total / 2.0


def _point_in_triangle(p, a, b, c):
    def sign(p1, p2, p3):
        return (p1[0] - p3[0]) * (p2[1] - p3[1]) - (p2[0] - p3[0]) * (p1[1] - p3[1])
    d1, d2, d3 = sign(p, a, b), sign(p, b, c), sign(p, c, a)
    has_neg = (d1 < 0) or (d2 < 0) or (d3 < 0)
    has_pos = (d1 > 0) or (d2 > 0) or (d3 > 0)
    return not (has_neg and has_pos)


def _decimate(points, limit=1200):
    if len(points) <= limit:
        return points
    step = len(points) / limit
    return [points[int(i * step)] for i in range(limit)]


def _triangulate_polygon(outer, holes):
    """Ear-Clipping mit Loch-Brücken. outer CCW, Löcher CW."""
    outer = _decimate(list(outer))
    polygon = outer
    for hole in sorted((_decimate(list(h)) for h in holes),
                       key=lambda h: -max(p[0] for p in h)):
        bridge_h = max(range(len(hole)), key=lambda i: hole[i][0])
        hp = hole[bridge_h]
        bridge_o = min(range(len(polygon)),
                       key=lambda i: (polygon[i][0] - hp[0]) ** 2 + (polygon[i][1] - hp[1]) ** 2)
        polygon = (polygon[:bridge_o + 1]
                   + hole[bridge_h:] + hole[:bridge_h + 1]
                   + polygon[bridge_o:])

    indices = list(range(len(polygon)))
    triangles = []
    guard = 0
    while len(indices) > 3 and guard < 20000:
        guard += 1
        ear_found = False
        for k in range(len(indices)):
            i_prev = indices[k - 1]
            i_cur = indices[k]
            i_next = indices[(k + 1) % len(indices)]
            a, b, c = polygon[i_prev], polygon[i_cur], polygon[i_next]
            cross = (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])
            if cross <= 1e-14:
                continue
            contains = False
            for other in indices:
                if other in (i_prev, i_cur, i_next):
                    continue
                if _point_in_triangle(polygon[other], a, b, c):
                    contains = True
                    break
            if contains:
                continue
            triangles.append((a, b, c))
            del indices[k]
            ear_found = True
            break
        if not ear_found:                          # degeneriert: Fächer-Fallback
            for k in range(1, len(indices) - 1):
                triangles.append((polygon[indices[0]], polygon[indices[k]], polygon[indices[k + 1]]))
            return triangles
    if len(indices) == 3:
        triangles.append((polygon[indices[0]], polygon[indices[1]], polygon[indices[2]]))
    return triangles


def _refine_triangles(triangles, target, max_tris):
    """Gleichmäßige 4-fach-Unterteilung, bis Kanten die Zielweite erreichen."""
    target_u, target_v = target

    def worst(tris):
        value = 0.0
        for a, b, c in tris:
            for p, q in ((a, b), (b, c), (c, a)):
                value = max(value, math.hypot((p[0] - q[0]) / target_u if target_u != float("inf") else 0.0,
                                              (p[1] - q[1]) / target_v if target_v != float("inf") else 0.0))
        return value

    for _ in range(5):
        if worst(triangles) <= 1.0 or len(triangles) * 4 > max_tris:
            break
        refined = []
        for a, b, c in triangles:
            ab = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
            bc = ((b[0] + c[0]) / 2, (b[1] + c[1]) / 2)
            ca = ((c[0] + a[0]) / 2, (c[1] + a[1]) / 2)
            refined.extend([(a, ab, ca), (ab, b, bc), (ca, bc, c), (ab, bc, ca)])
        triangles = refined
    return triangles


# ---------------------------------------------------------------------------
# Öffentliche API
# ---------------------------------------------------------------------------

def read_step_mesh(path, quality="mittel", log=None):
    """STEP-Datei tessellieren. quality: grob | mittel | fein."""
    path = Path(path)
    segments = QUALITY_SEGMENTS.get(quality, 48)
    text = path.read_text(encoding="latin-1", errors="replace")
    step = StepFile(text)
    products = step.by_type("PRODUCT")
    if log and len(products) > 1:
        log(f"Hinweis: {len(products)} Produkte in der Datei - Baugruppen-"
            "Transformationen werden nicht angewendet (Teile ggf. am Ursprung).")
    tess = StepTessellator(step, segments=segments, log=log)
    return tess.run(path.stem)


# ===========================================================================
# IGES-Erweiterung des Tessellierungs-Kerns
# ===========================================================================

"""IGES -> Mesh: Erweiterung des Tessellierungs-Kerns.

Unterstützte Entitäten:
  Kurven:  100 (Kreisbogen), 102 (Verbundkurve), 106 (Punktfolge),
           110 (Linie), 126 (NURBS-Kurve)
  Flächen: 120 (Rotationsfläche), 122 (Extrusionsfläche/Tabulated Cylinder),
           128 (NURBS-Fläche)
  Trimmung: 144 (Getrimmte Fläche) über 142 (Kurve auf Fläche) -
           bevorzugt über die Parameterraum-Kurve (B-Pointer), sonst
           Modellraum-Kurve + UV-Inversion
  124 (Transformationsmatrix) wird auf alle ausgewerteten Punkte angewandt.

Nicht referenzierte, ungetrimmte Flächen werden über die volle Domäne
tesselliert. Nicht unterstützte Typen werden gezählt und geloggt.
"""




class IgesFile:
    def __init__(self, text):
        self.directory = []
        self.plines = []
        param_delim, record_delim = ",", ";"
        d_lines = []
        for line in text.splitlines():
            if len(line) < 73:
                continue
            section = line[72]
            if section == "G" and line[:72].lstrip().startswith("1H"):
                head = line.lstrip()
                param_delim = head[2] if len(head) > 2 else ","
            elif section == "D":
                d_lines.append(line[:72])
            elif section == "P":
                self.plines.append(line[:64])
        self.param_delim = param_delim
        self.record_delim = record_delim

        def field(raw, index):
            chunk = raw[(index - 1) * 8:index * 8].strip()
            try:
                return int(chunk)
            except ValueError:
                return 0

        for i in range(0, len(d_lines) - 1, 2):
            line1, line2 = d_lines[i], d_lines[i + 1]
            self.directory.append({
                "type": field(line1, 1),
                "pptr": field(line1, 2),
                "trans": field(line1, 7),
                "count": field(line2, 4),
                "form": field(line2, 5),
            })

    def entry(self, de_pointer):
        """DE-Zeiger (ungerade Sequenznummer) -> Verzeichniseintrag."""
        index = (int(de_pointer) - 1) // 2
        if 0 <= index < len(self.directory):
            return self.directory[index]
        return None

    def params(self, entry):
        start = entry["pptr"] - 1
        text = "".join(self.plines[start:start + max(1, entry["count"])])
        text = text.split(self.record_delim)[0]
        fields = []
        for chunk in text.split(self.param_delim):
            chunk = chunk.strip()
            if not chunk:
                fields.append(0.0)
                continue
            try:
                fields.append(float(chunk))
            except ValueError:
                fields.append(chunk)
        return fields[1:]                     # Feld 0 = Entitätstyp


class IgesTessellator:
    def __init__(self, iges, segments=48, log=None):
        self.iges = iges
        self.segments = segments
        self.log = log or (lambda message: None)
        self.pool = _VertexPool()
        self.tris = []
        self.skipped = Counter()

    # -- Transformationen --------------------------------------------------

    def transform_of(self, entry, depth=0):
        if not entry or not entry.get("trans") or depth > 8:
            return None
        t_entry = self.iges.entry(entry["trans"])
        if not t_entry or t_entry["type"] != 124:
            return None
        p = self.iges.params(t_entry)
        rotation = ((p[0], p[1], p[2]), (p[4], p[5], p[6]), (p[8], p[9], p[10]))
        translation = (p[3], p[7], p[11])
        parent = self.transform_of(t_entry, depth + 1)

        def apply(point, R=rotation, T=translation, outer=parent):
            x = R[0][0] * point[0] + R[0][1] * point[1] + R[0][2] * point[2] + T[0]
            y = R[1][0] * point[0] + R[1][1] * point[1] + R[1][2] * point[2] + T[1]
            z = R[2][0] * point[0] + R[2][1] * point[1] + R[2][2] * point[2] + T[2]
            return outer((x, y, z)) if outer else (x, y, z)
        return apply

    # -- Kurven ------------------------------------------------------------

    def curve_points(self, de_pointer, samples=None):
        """3D-Punktfolge einer Kurven-Entität (Transformation angewandt)."""
        entry = self.iges.entry(de_pointer)
        if entry is None:
            return []
        samples = samples or max(8, self.segments // 2)
        p = self.iges.params(entry)
        etype = entry["type"]
        points = []

        if etype == 110:
            points = [(p[0], p[1], p[2]), (p[3], p[4], p[5])]
        elif etype == 100:
            zt, xc, yc, x1, y1, x2, y2 = p[:7]
            a0 = math.atan2(y1 - yc, x1 - xc)
            a1 = math.atan2(y2 - yc, x2 - xc)
            if a1 <= a0 + 1e-12:
                a1 += TWO_PI
            radius = math.hypot(x1 - xc, y1 - yc)
            steps = max(4, int(round(self.segments * (a1 - a0) / TWO_PI)))
            points = [(xc + radius * math.cos(a0 + (a1 - a0) * i / steps),
                       yc + radius * math.sin(a0 + (a1 - a0) * i / steps), zt)
                      for i in range(steps + 1)]
        elif etype == 102:
            count = int(p[0])
            for sub in p[1:1 + count]:
                seg = self.curve_points(int(sub), samples)
                if points and seg and math.dist(points[-1], seg[0]) < 1e-9:
                    seg = seg[1:]
                points.extend(seg)
            return points                      # Unterkurven tragen ihre Transformationen selbst
        elif etype == 106:
            ip = int(p[0])
            count = int(p[1])
            data = p[2:]
            if ip == 1:                        # gemeinsame z-Ebene
                z = data[0]
                points = [(data[1 + 2 * i], data[2 + 2 * i], z) for i in range(count)]
            elif ip == 2:
                points = [(data[3 * i], data[3 * i + 1], data[3 * i + 2]) for i in range(count)]
            else:
                points = [(data[6 * i], data[6 * i + 1], data[6 * i + 2]) for i in range(count)]
        elif etype == 126:
            curve = self._nurbs_curve(p)
            points = [curve.ev(curve.t0 + (curve.t1 - curve.t0) * i / samples)
                      for i in range(samples + 1)]
        else:
            self.skipped[f"Kurve Typ {etype}"] += 1
            return []

        transform = self.transform_of(entry)
        if transform:
            points = [transform(pt) for pt in points]
        return points

    @staticmethod
    def _n126_lengths(p):
        upper = int(p[0])
        degree = int(p[1])
        return (upper + degree + 2, upper + 1)

    def _nurbs_curve(self, p):
        upper = int(p[0])
        degree = int(p[1])
        knot_count, point_count = self._n126_lengths(p)
        base = 6
        knots = [float(v) for v in p[base:base + knot_count]]
        weights = [float(v) for v in p[base + knot_count:base + knot_count + point_count]]
        coords = p[base + knot_count + point_count:base + knot_count + point_count + 3 * point_count]
        points = [(float(coords[3 * i]), float(coords[3 * i + 1]), float(coords[3 * i + 2]))
                  for i in range(point_count)]
        return BSplineCurve(degree, points, weights, knots)

    def curve_uv_points(self, de_pointer, samples=None):
        """Parameterraum-Kurve (im 142-B-Zeiger): x,y der Punkte = u,v."""
        return [(pt[0], pt[1]) for pt in self.curve_points(de_pointer, samples)]

    # -- Flächen -----------------------------------------------------------

    def surface_of(self, de_pointer):
        entry = self.iges.entry(de_pointer)
        if entry is None:
            return None
        p = self.iges.params(entry)
        etype = entry["type"]
        transform = self.transform_of(entry)

        def wrap(ev):
            if transform is None:
                return ev
            def transformed(u, v, base=ev, tf=transform):
                return tf(base(u, v))
            return transformed

        if etype == 128:
            k1, k2 = int(p[0]), int(p[1])
            m1, m2 = int(p[2]), int(p[3])
            nk1 = k1 + m1 + 2
            nk2 = k2 + m2 + 2
            n_pts = (k1 + 1) * (k2 + 1)
            base = 9
            u_knots = [float(v) for v in p[base:base + nk1]]
            v_knots = [float(v) for v in p[base + nk1:base + nk1 + nk2]]
            w_start = base + nk1 + nk2
            weights_flat = [float(v) for v in p[w_start:w_start + n_pts]]
            c_start = w_start + n_pts
            coords = p[c_start:c_start + 3 * n_pts]
            # Index 1 (u) läuft am schnellsten
            grid = [[None] * (k2 + 1) for _ in range(k1 + 1)]
            weights = [[1.0] * (k2 + 1) for _ in range(k1 + 1)]
            for j in range(k2 + 1):
                for i in range(k1 + 1):
                    flat = j * (k1 + 1) + i
                    grid[i][j] = (float(coords[3 * flat]), float(coords[3 * flat + 1]),
                                  float(coords[3 * flat + 2]))
                    weights[i][j] = weights_flat[flat]
            spline = BSplineSurface(m1, m2, grid, weights, u_knots, v_knots)
            ev = wrap(spline.ev)
            corner_a = ev(spline.u0, spline.v0)
            corner_b = ev(spline.u1, spline.v1)
            extent = max(math.dist(corner_a, corner_b), 1e-6)
            flat_surface = (m1 == 1 and m2 == 1 and k1 == 1 and k2 == 1)
            surface = Surface(ev, _generic_uv_of(ev, spline.u0, spline.u1, spline.v0, spline.v1),
                              kind="nurbs", curved_u=not flat_surface, curved_v=not flat_surface,
                              u_scale=extent / max(spline.u1 - spline.u0, 1e-9),
                              v_scale=extent / max(spline.v1 - spline.v0, 1e-9))
            surface.domain = (spline.u0, spline.u1, spline.v0, spline.v1)
            return surface

        if etype == 120:
            axis_pts = self.curve_points(int(p[0]), samples=1)
            if len(axis_pts) < 2:
                return None
            origin = axis_pts[0]
            direction = _norm(_sub(axis_pts[1], axis_pts[0]))
            gen_entry = self.iges.entry(int(p[1]))
            gen_samples = max(8, self.segments // 2)
            gen_points = self.curve_points(int(p[1]), gen_samples)
            if len(gen_points) < 2:
                return None
            sa, ta = float(p[2]), float(p[3])
            if ta <= sa:
                ta += TWO_PI

            def ev(u, v, o=origin, z=direction, pts=gen_points):
                t = min(max(v, 0.0), 1.0) * (len(pts) - 1)
                i = min(int(t), len(pts) - 2)
                frac = t - i
                point = _add(_mul(pts[i], 1 - frac), _mul(pts[i + 1], frac))
                d = _sub(point, o)
                par = _mul(z, _dot(d, z))
                perp = _sub(d, par)
                ortho = _cross(z, perp)
                return _add(o, _add(_add(_mul(perp, math.cos(u)), _mul(ortho, math.sin(u))), par))

            radius = max(max(math.dist(pt, origin) for pt in gen_points), 1e-6)
            full = abs((ta - sa) - TWO_PI) < 1e-3
            surface = Surface(ev, _generic_uv_of(ev, sa, ta, 0.0, 1.0),
                              u_period=TWO_PI if full else None,
                              kind="revolution", curved_v=False, u_scale=radius)
            surface.domain = (sa, ta, 0.0, 1.0)
            return surface

        if etype == 122:
            base_samples = max(8, self.segments // 2)
            base_points = self.curve_points(int(p[0]), base_samples)
            if len(base_points) < 2:
                return None
            direction = _sub((float(p[1]), float(p[2]), float(p[3])), base_points[0])

            def ev(u, v, pts=base_points, d=direction):
                t = min(max(u, 0.0), 1.0) * (len(pts) - 1)
                i = min(int(t), len(pts) - 2)
                frac = t - i
                point = _add(_mul(pts[i], 1 - frac), _mul(pts[i + 1], frac))
                return _add(point, _mul(d, v))

            extent = max(math.dist(base_points[0], base_points[-1]), 1e-6)
            surface = Surface(ev, _generic_uv_of(ev, 0.0, 1.0, 0.0, 1.0),
                              kind="extrusion", curved_v=False, u_scale=extent)
            surface.domain = (0.0, 1.0, 0.0, 1.0)
            return surface

        self.skipped[f"Fläche Typ {etype}"] += 1
        return None

    # -- Getrimmte Flächen -------------------------------------------------

    def _boundary_uv(self, cos_pointer, surface):
        """142 (Kurve auf Fläche) -> UV-Loop."""
        entry = self.iges.entry(int(cos_pointer))
        if entry is None or entry["type"] != 142:
            return []
        p = self.iges.params(entry)
        b_pointer = int(p[2]) if len(p) > 2 else 0
        c_pointer = int(p[3]) if len(p) > 3 else 0
        if b_pointer:
            uvs = self.curve_uv_points(b_pointer)
            if uvs:
                return uvs
        if c_pointer:
            points = self.curve_points(c_pointer)
            uvs = [surface.uv_of(pt) for pt in points]
            uvs, _w = _unwrap_loop(uvs, surface.u_period, surface.v_period)
            return uvs
        return []

    def tessellate_trimmed(self, entry):
        p = self.iges.params(entry)
        surface = self.surface_of(int(p[0]))
        if surface is None:
            return
        outer_flag = int(p[1])
        hole_count = int(p[2])
        outer_pointer = int(p[3]) if len(p) > 3 else 0
        loops = []
        if outer_flag and outer_pointer:
            outer = self._boundary_uv(outer_pointer, surface)
            if len(outer) >= 3:
                loops.append(outer)
        for k in range(hole_count):
            if 4 + k < len(p):
                hole = self._boundary_uv(int(p[4 + k]), surface)
                if len(hole) >= 3:
                    loops.append(hole)
        if loops:
            triangles = face_triangles_from_uv(surface, loops, self.segments)
        else:
            triangles = grid_triangles(surface, [], self.segments,
                                       domain=getattr(surface, "domain", None))
        self._emit(surface, triangles)

    def tessellate_untrimmed(self, de_pointer):
        surface = self.surface_of(de_pointer)
        if surface is None:
            return
        triangles = grid_triangles(surface, [], self.segments,
                                   domain=getattr(surface, "domain", None))
        self._emit(surface, triangles)

    def _emit(self, surface, triangles):
        for a, b, c in triangles:
            pa = self.pool.add(*surface.ev(*a))
            pb = self.pool.add(*surface.ev(*b))
            pc = self.pool.add(*surface.ev(*c))
            if pa != pb and pb != pc and pa != pc:
                self.tris.append((pa, pb, pc))

    # -- Gesamtablauf ------------------------------------------------------

    def run(self, name):
        trimmed = []
        referenced = set()
        for index, entry in enumerate(self.iges.directory):
            if entry["type"] == 144:
                trimmed.append(entry)
                p = self.iges.params(entry)
                referenced.add(int(p[0]))
                for pointer in p[3:4 + int(p[2])]:
                    cos = self.iges.entry(int(pointer)) if pointer else None
                    if cos and cos["type"] == 142:
                        cp = self.iges.params(cos)
                        for sub in cp[1:4]:
                            if isinstance(sub, float) and sub:
                                referenced.add(int(sub))

        for entry in trimmed:
            try:
                self.tessellate_trimmed(entry)
            except Exception:
                self.skipped["Fehler bei Fläche"] += 1

        for index, entry in enumerate(self.iges.directory):
            de_pointer = 2 * index + 1
            if entry["type"] in (120, 122, 128) and de_pointer not in referenced:
                try:
                    self.tessellate_untrimmed(de_pointer)
                except Exception:
                    self.skipped["Fehler bei Fläche"] += 1

        if not self.tris:
            raise ValueError("Keine tessellierbare Fläche gefunden (unterstützt: "
                             "120/122/128/144; B-Rep-Solids 186 brauchen 144-Flächen).")
        if self.skipped:
            details = ", ".join(f"{k} ({v})" for k, v in self.skipped.most_common())
            self.log(f"Übersprungen: {details}")
        self.log(f"Tesselliert: {len(self.tris)} Dreiecke.")
        return Mesh(self.pool.vertices, self.tris, name)


def read_iges_mesh(path, quality="mittel", log=None):
    path = Path(path)
    segments = QUALITY_SEGMENTS.get(quality, 48)
    text = path.read_text(encoding="latin-1", errors="replace")
    iges = IgesFile(text)
    if not iges.directory:
        raise ValueError("Keine IGES-Verzeichnissektion gefunden.")
    return IgesTessellator(iges, segments=segments, log=log).run(path.stem)


# ===========================================================================
# PDF: mergen, splitten, drehen, umsortieren, Zonen abdecken
# ===========================================================================

"""PDF-Werkzeuge auf Basis der Standardbibliothek: mergen, splitten, Zonen abdecken.

Eigener Mini-Parser/-Writer für das PDF-Objektmodell (COS): Objekte werden per
Scan eingesammelt (robust auch bei kaputten xref-Tabellen), komprimierte
Objekt-Streams (/ObjStm, Flate) werden entpackt. Beim Schreiben wird der
Objektgraph der benötigten Seiten kopiert und neu nummeriert.

Grenzen: verschlüsselte PDFs werden abgelehnt; „Abdecken" legt eine deckende
Fläche ÜBER den Inhalt (der Text darunter bleibt technisch extrahierbar -
für echte Schwärzung Dokument neu erzeugen lassen).
"""



class PdfError(ValueError):
    pass


class PName(str):
    """PDF-Name /Xyz."""


class PRef(int):
    """Indirekter Verweis n 0 R (Generation wird ignoriert)."""


_WHITESPACE = b"\x00\t\n\x0c\r "
_DELIMITER = b"()<>[]{}/%"


# ---------------------------------------------------------------------------
# Lexer / Objekt-Parser
# ---------------------------------------------------------------------------

class _Lexer:
    def __init__(self, data, pos=0):
        self.data = data
        self.pos = pos

    def _skip_ws(self):
        data, pos = self.data, self.pos
        while pos < len(data):
            ch = data[pos:pos + 1]
            if ch in _WHITESPACE:
                pos += 1
            elif ch == b"%":
                end = data.find(b"\n", pos)
                pos = len(data) if end == -1 else end + 1
            else:
                break
        self.pos = pos

    def parse(self):
        self._skip_ws()
        data, pos = self.data, self.pos
        if pos >= len(data):
            raise PdfError("Unerwartetes Dateiende.")
        ch = data[pos:pos + 1]

        if data[pos:pos + 2] == b"<<":
            self.pos = pos + 2
            result = {}
            while True:
                self._skip_ws()
                if self.data[self.pos:self.pos + 2] == b">>":
                    self.pos += 2
                    return result
                key = self.parse()
                if not isinstance(key, PName):
                    raise PdfError("Dictionary-Schlüssel ist kein Name.")
                result[str(key)] = self.parse()

        if ch == b"[":
            self.pos = pos + 1
            result = []
            while True:
                self._skip_ws()
                if self.data[self.pos:self.pos + 1] == b"]":
                    self.pos += 1
                    return result
                result.append(self.parse())

        if ch == b"/":
            match = re.match(rb"/([^\x00\t\n\x0c\r ()<>\[\]{}/%]*)", data[pos:])
            self.pos = pos + match.end()
            name = match.group(1)
            # #xx-Escapes in Namen auflösen
            name = re.sub(rb"#([0-9A-Fa-f]{2})", lambda m: bytes([int(m.group(1), 16)]), name)
            return PName(name.decode("latin-1"))

        if ch == b"(":
            depth = 0
            out = bytearray()
            i = pos
            while i < len(data):
                c = data[i:i + 1]
                if c == b"\\":
                    nxt = data[i + 1:i + 2]
                    mapping = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b",
                               b"f": b"\x0c", b"(": b"(", b")": b")", b"\\": b"\\"}
                    if nxt in mapping:
                        out += mapping[nxt]
                        i += 2
                        continue
                    octal = re.match(rb"\\([0-7]{1,3})", data[i:])
                    if octal:
                        out.append(int(octal.group(1), 8) & 0xFF)
                        i += octal.end()
                        continue
                    i += 1
                    continue
                if c == b"(":
                    depth += 1
                    if depth > 1:
                        out += c
                elif c == b")":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                    out += c
                else:
                    out += c
                i += 1
            self.pos = i
            return bytes(out)

        if ch == b"<":
            end = data.find(b">", pos)
            hexstr = re.sub(rb"[^0-9A-Fa-f]", b"", data[pos + 1:end])
            if len(hexstr) % 2:
                hexstr += b"0"
            self.pos = end + 1
            return bytes.fromhex(hexstr.decode("ascii"))

        match = re.match(rb"[+-]?(\d+\.\d*|\.\d+|\d+)", data[pos:])
        if match:
            raw = match.group(0)
            self.pos = pos + match.end()
            if b"." not in raw:
                # Verweis "n g R"?
                ref = re.match(rb"\s+(\d+)\s+R(?![A-Za-z0-9])", data[self.pos:])
                if ref:
                    self.pos += ref.end()
                    return PRef(int(raw))
                return int(raw)
            return float(raw)

        for literal, value in ((b"true", True), (b"false", False), (b"null", None)):
            if data[pos:pos + len(literal)] == literal:
                self.pos = pos + len(literal)
                return value
        raise PdfError(f"Unbekanntes Token an Position {pos}: {data[pos:pos + 20]!r}")


class _PdfObject:
    __slots__ = ("value", "stream")

    def __init__(self, value, stream=None):
        self.value = value
        self.stream = stream


# ---------------------------------------------------------------------------
# Dokument einlesen
# ---------------------------------------------------------------------------

class PdfDocument:
    def __init__(self, path):
        data = Path(path).read_bytes()
        if b"%PDF" not in data[:1024]:
            raise PdfError("Keine PDF-Datei.")
        if re.search(rb"/Encrypt\s+\d+\s+\d+\s+R", data) or b"/Encrypt<<" in data:
            raise PdfError("PDF ist verschlüsselt - bitte entsperrte Fassung anliefern lassen.")
        self.objects = {}
        self._scan_objects(data)
        self._expand_object_streams()
        if not self.objects:
            raise PdfError("Keine PDF-Objekte gefunden.")

    def _scan_objects(self, data):
        skip_until = 0
        for match in re.finditer(rb"(\d+)\s+\d+\s+obj\b", data):
            if match.start() < skip_until:
                continue
            num = int(match.group(1))
            lexer = _Lexer(data, match.end())
            try:
                value = lexer.parse()
            except PdfError:
                continue
            stream = None
            rest = data[lexer.pos:lexer.pos + 20]
            ws = len(rest) - len(rest.lstrip(_WHITESPACE))
            if data[lexer.pos + ws:lexer.pos + ws + 6] == b"stream":
                start = lexer.pos + ws + 6
                if data[start:start + 2] == b"\r\n":
                    start += 2
                elif data[start:start + 1] in (b"\n", b"\r"):
                    start += 1
                length = value.get("Length") if isinstance(value, dict) else None
                end = -1
                if isinstance(length, int) and start + length <= len(data):
                    after = data[start + length:start + length + 20].lstrip(b"\r\n \t")
                    if after.startswith(b"endstream"):      # /Length ist verlässlich
                        stream = data[start:start + length]
                        end = data.find(b"endstream", start + length)
                if end == -1:
                    end = data.find(b"endstream", start)
                    stream = data[start:end].rstrip(b"\r\n") if end != -1 else data[start:]
                skip_until = (end if end != -1 else len(data)) + 9
            else:
                skip_until = lexer.pos
            # Spätere Vorkommen (inkrementelle Updates) überschreiben frühere.
            self.objects[num] = _PdfObject(value, stream)

        # /Length als Verweis nachträglich auflösen
        for obj in self.objects.values():
            if obj.stream is not None and isinstance(obj.value, dict):
                length = obj.value.get("Length")
                if isinstance(length, PRef):
                    resolved = self.resolve(length)
                    if isinstance(resolved, int) and resolved <= len(obj.stream):
                        obj.stream = obj.stream[:resolved]

    def _expand_object_streams(self):
        for container in list(self.objects.values()):
            value = container.value
            if not (isinstance(value, dict) and value.get("Type") == "ObjStm" and container.stream):
                continue
            filters = value.get("Filter")
            if filters not in ("FlateDecode", ["FlateDecode"], PName("FlateDecode")):
                continue
            try:
                payload = zlib.decompress(container.stream)
            except zlib.error:
                continue
            count = self.resolve(value.get("N", 0)) or 0
            first = self.resolve(value.get("First", 0)) or 0
            header = payload[:first].split()
            for i in range(int(count)):
                try:
                    num = int(header[2 * i])
                    offset = int(header[2 * i + 1])
                    inner = _Lexer(payload, first + offset).parse()
                except (IndexError, ValueError, PdfError):
                    continue
                if num not in self.objects:
                    self.objects[num] = _PdfObject(inner)

    def resolve(self, value):
        seen = set()
        while isinstance(value, PRef):
            if int(value) in seen or int(value) not in self.objects:
                return None
            seen.add(int(value))
            value = self.objects[int(value)].value
        return value

    # -- Seitenbaum --------------------------------------------------------

    _INHERITED = ("Resources", "MediaBox", "CropBox", "Rotate")

    def pages(self):
        catalog = None
        for obj in self.objects.values():
            if isinstance(obj.value, dict) and obj.value.get("Type") == "Catalog" \
                    and "Pages" in obj.value:
                catalog = obj.value
        if catalog is None:
            raise PdfError("Kein Dokumentkatalog gefunden.")
        result = []

        def walk(ref, inherited, depth=0):
            if depth > 60:
                return
            node = self.resolve(ref)
            if not isinstance(node, dict):
                return
            merged = dict(inherited)
            for key in self._INHERITED:
                if key in node:
                    merged[key] = node[key]
            if node.get("Type") == "Page" or ("Contents" in node and "Kids" not in node):
                result.append((ref, merged))
            else:
                for kid in self.resolve(node.get("Kids")) or []:
                    walk(kid, merged, depth + 1)

        walk(catalog["Pages"], {})
        if not result:
            raise PdfError("PDF enthält keine Seiten.")
        return result


# ---------------------------------------------------------------------------
# Schreiben
# ---------------------------------------------------------------------------

def _serialize(value):
    if isinstance(value, PRef):
        return f"{int(value)} 0 R".encode("ascii")
    if isinstance(value, PName):
        escaped = re.sub(r"([^\x21-\x7e]|[()<>\[\]{}/%#])",
                         lambda m: f"#{ord(m.group(1)):02X}", str(value))
        return b"/" + escaped.encode("ascii")
    if value is True:
        return b"true"
    if value is False:
        return b"false"
    if value is None:
        return b"null"
    if isinstance(value, bool):
        return b"true" if value else b"false"
    if isinstance(value, int):
        return str(value).encode("ascii")
    if isinstance(value, float):
        return f"{value:.4f}".rstrip("0").rstrip(".").encode("ascii") or b"0"
    if isinstance(value, bytes):
        return b"<" + value.hex().encode("ascii") + b">"
    if isinstance(value, str):
        return b"<" + value.encode("latin-1", "replace").hex().encode("ascii") + b">"
    if isinstance(value, list):
        return b"[" + b" ".join(_serialize(v) for v in value) + b"]"
    if isinstance(value, dict):
        parts = [b"<<"]
        for key, item in value.items():
            parts.append(_serialize(PName(key)) + b" " + _serialize(item))
        parts.append(b">>")
        return b"\n".join(parts)
    raise PdfError(f"Nicht serialisierbar: {type(value)}")


class PdfWriter:
    def __init__(self):
        self.items = [None]                      # Index = neue Objektnummer

    def reserve(self):
        self.items.append(None)
        return PRef(len(self.items) - 1)

    def set(self, ref, value, stream=None):
        if stream is not None and isinstance(value, dict):
            value = dict(value)
            value["Length"] = len(stream)
        self.items[int(ref)] = (value, stream)
        return ref

    def add(self, value, stream=None):
        return self.set(self.reserve(), value, stream)

    def copy_object(self, doc, ref, memo, strip_keys=()):
        key = (id(doc), int(ref))
        if key in memo:
            return memo[key]
        new_ref = self.reserve()
        memo[key] = new_ref
        source = doc.objects.get(int(ref))
        if source is None:
            self.set(new_ref, None)
            return new_ref
        value = source.value
        if strip_keys and isinstance(value, dict):
            value = {k: v for k, v in value.items() if k not in strip_keys}
        self.set(new_ref, self._deep_copy(doc, value, memo), source.stream)
        return new_ref

    def _deep_copy(self, doc, value, memo):
        if isinstance(value, PRef):
            return self.copy_object(doc, value, memo)
        if isinstance(value, dict):
            return {k: self._deep_copy(doc, v, memo) for k, v in value.items()}
        if isinstance(value, list):
            return [self._deep_copy(doc, v, memo) for v in value]
        return value

    def import_page(self, doc, page_ref, inherited, memo):
        page = doc.resolve(page_ref)
        merged = dict(page)
        for key, val in inherited.items():
            merged.setdefault(key, val)
        merged.pop("Parent", None)
        merged["Type"] = PName("Page")
        new_ref = self.reserve()
        memo[(id(doc), int(page_ref))] = new_ref
        self.set(new_ref, self._deep_copy(doc, merged, memo))
        return new_ref

    def save(self, path, page_refs):
        pages_ref = self.reserve()
        for ref in page_refs:
            value, stream = self.items[int(ref)]
            value["Parent"] = pages_ref
            self.items[int(ref)] = (value, stream)
        self.set(pages_ref, {"Type": PName("Pages"),
                             "Kids": list(page_refs), "Count": len(page_refs)})
        catalog_ref = self.add({"Type": PName("Catalog"), "Pages": pages_ref})

        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for num in range(1, len(self.items)):
            offsets.append(len(out))
            value, stream = self.items[num] or (None, None)
            out += f"{num} 0 obj\n".encode("ascii")
            out += _serialize(value)
            if stream is not None:
                out += b"\nstream\n" + stream + b"\nendstream"
            out += b"\nendobj\n"
        xref_pos = len(out)
        out += f"xref\n0 {len(self.items)}\n".encode("ascii")
        out += b"0000000000 65535 f \n"
        for offset in offsets[1:]:
            out += f"{offset:010d} 00000 n \n".encode("ascii")
        out += b"trailer\n" + _serialize({"Size": len(self.items), "Root": catalog_ref})
        out += f"\nstartxref\n{xref_pos}\n%%EOF\n".encode("ascii")
        Path(path).write_bytes(bytes(out))


# ---------------------------------------------------------------------------
# Öffentliche Funktionen
# ---------------------------------------------------------------------------

def pdf_page_count(path):
    return len(PdfDocument(path).pages())


def parse_page_ranges(text, total):
    """'1-3,7' -> [0, 1, 2, 6]; leer = alle Seiten."""
    text = text.strip()
    if not text:
        return list(range(total))
    result = []
    for chunk in text.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", chunk)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
        elif chunk.isdigit():
            start = end = int(chunk)
        else:
            raise PdfError(f"Ungültiger Seitenbereich: „{chunk}“ (erwartet z. B. 1-3,7)")
        for page in range(start, end + 1):
            if 1 <= page <= total and (page - 1) not in result:
                result.append(page - 1)
    if not result:
        raise PdfError("Der Seitenbereich trifft keine Seite.")
    return result


def merge_pdfs(paths, target_path):
    """Mehrere PDFs in Reihenfolge zu einer Datei zusammenführen."""
    writer = PdfWriter()
    page_refs = []
    for path in paths:
        doc = PdfDocument(path)
        memo = {}
        for page_ref, inherited in doc.pages():
            page_refs.append(writer.import_page(doc, page_ref, inherited, memo))
    writer.save(target_path, page_refs)
    return len(page_refs)


def split_pdf(path, target_dir, ranges_text="", single_pages=True, log=None):
    """PDF zerlegen: jede Seite einzeln oder ein Auszug als eine Datei."""
    doc = PdfDocument(path)
    pages = doc.pages()
    indexes = parse_page_ranges(ranges_text, len(pages))
    target_dir = Path(target_dir)
    stem = Path(path).stem
    written = []
    if single_pages:
        for index in indexes:
            writer = PdfWriter()
            memo = {}
            ref = writer.import_page(doc, pages[index][0], pages[index][1], memo)
            out = target_dir / f"{stem}_S{index + 1:03d}.pdf"
            writer.save(out, [ref])
            written.append(out)
            if log:
                log(f"SEITE: {out.name}")
    else:
        writer = PdfWriter()
        memo = {}
        refs = [writer.import_page(doc, pages[i][0], pages[i][1], memo) for i in indexes]
        out = target_dir / f"{stem}_auszug.pdf"
        writer.save(out, refs)
        written.append(out)
        if log:
            log(f"AUSZUG: {out.name} ({len(refs)} Seiten)")
    return written


def parse_page_order(text, total):
    """Wie parse_page_ranges, aber Reihenfolge und Duplikate bleiben erhalten."""
    text = text.strip()
    if not text:
        return list(range(total))
    result = []
    for chunk in text.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", chunk)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            step = 1 if end >= start else -1
            pages = range(start, end + step, step)
        elif chunk.isdigit():
            pages = [int(chunk)]
        else:
            raise PdfError(f"Ungültige Seitenangabe: „{chunk}“ (erwartet z. B. 3,1-2)")
        for page in pages:
            if 1 <= page <= total:
                result.append(page - 1)
    if not result:
        raise PdfError("Die Seitenangabe trifft keine Seite.")
    return result


def rotate_pdf(path, target_path, angle, ranges_text="", log=None):
    """Seiten um 90/180/270 Grad drehen (im Uhrzeigersinn)."""
    if angle % 90:
        raise PdfError("Drehwinkel muss ein Vielfaches von 90 sein.")
    doc = PdfDocument(path)
    pages = doc.pages()
    indexes = set(parse_page_ranges(ranges_text, len(pages)))
    writer = PdfWriter()
    memo = {}
    page_refs = []
    rotated = 0
    for index, (page_ref, inherited) in enumerate(pages):
        new_ref = writer.import_page(doc, page_ref, inherited, memo)
        page_refs.append(new_ref)
        if index in indexes:
            value, stream = writer.items[int(new_ref)]
            current = doc.resolve(inherited.get("Rotate")) or value.get("Rotate") or 0
            value["Rotate"] = (int(current) + int(angle)) % 360
            writer.items[int(new_ref)] = (value, stream)
            rotated += 1
            if log:
                log(f"GEDREHT: Seite {index + 1} um {angle}°")
    writer.save(target_path, page_refs)
    return rotated


def reorder_pdf(path, target_path, order_text, log=None):
    """Seiten in neuer Reihenfolge schreiben (z. B. „3,1-2“; Duplikate erlaubt)."""
    doc = PdfDocument(path)
    pages = doc.pages()
    order = parse_page_order(order_text, len(pages))
    writer = PdfWriter()
    memo = {}
    page_refs = [writer.import_page(doc, pages[i][0], pages[i][1], memo) for i in order]
    writer.save(target_path, page_refs)
    if log:
        log(f"UMSORTIERT: {len(order)} Seiten in Reihenfolge {', '.join(str(i + 1) for i in order[:12])}"
            + (" …" if len(order) > 12 else ""))
    return len(order)


def redact_pdf(path, target_path, zone, color="weiss", ranges_text="", log=None):
    """Zone (x, y, breite, hoehe in % von links oben) deckend überlagern.

    Achtung: rein visuelle Abdeckung - darunterliegender Text bleibt in der
    Datei erhalten.
    """
    doc = PdfDocument(path)
    pages = doc.pages()
    indexes = set(parse_page_ranges(ranges_text, len(pages)))
    zx, zy, zw, zh = (max(0.0, min(100.0, float(v))) / 100.0 for v in zone)
    rgb = "1 1 1" if color == "weiss" else "0 0 0"

    writer = PdfWriter()
    memo = {}
    page_refs = []
    covered = 0
    for index, (page_ref, inherited) in enumerate(pages):
        new_ref = writer.import_page(doc, page_ref, inherited, memo)
        page_refs.append(new_ref)
        if index not in indexes:
            continue
        value, stream = writer.items[int(new_ref)]
        media = [doc.resolve(v) if isinstance(v, PRef) else v
                 for v in (doc.resolve(inherited.get("MediaBox")) or [0, 0, 595, 842])]
        x0, y0, x1, y1 = (float(v) for v in media)
        width, height = x1 - x0, y1 - y0
        rect_x = x0 + zx * width
        rect_y = y1 - (zy + zh) * height
        overlay = (f"Q q {rgb} rg {rect_x:.2f} {rect_y:.2f} "
                   f"{zw * width:.2f} {zh * height:.2f} re f Q").encode("ascii")
        pre_ref = writer.add({}, b"q")
        post_ref = writer.add({}, overlay)
        contents = value.get("Contents")
        content_list = contents if isinstance(contents, list) else ([contents] if contents else [])
        value["Contents"] = [pre_ref] + content_list + [post_ref]
        writer.items[int(new_ref)] = (value, stream)
        covered += 1
        if log:
            log(f"ABGEDECKT: Seite {index + 1}")
    writer.save(target_path, page_refs)
    return covered


# ===========================================================================
# Bilder: PNG/BMP-Codec, zuschneiden, skalieren, Wasserzeichen, DPI
# ===========================================================================

"""Bild-Werkzeuge auf Basis der Standardbibliothek: PNG- und BMP-Codec,
Skalieren (bilinear), Wasserzeichen (eigener 5x7-Pixelfont), DPI setzen.

JPEG/TIFF/HEIC brauchen echte Codecs und bleiben bewusst außen vor.
"""


IMG_READ_EXTS = [".png", ".bmp"]
IMG_WRITE_EXTS = [".png", ".bmp"]


@dataclass
class Image:
    width: int
    height: int
    mode: str                       # "RGB" oder "RGBA"
    data: bytearray = field(repr=False, default_factory=bytearray)
    dpi: int | None = None

    @property
    def channels(self):
        return 4 if self.mode == "RGBA" else 3

    def stats(self):
        return f"{self.width}×{self.height} {self.mode}" + (f", {self.dpi} dpi" if self.dpi else "")


# ---------------------------------------------------------------------------
# PNG
# ---------------------------------------------------------------------------

_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _paeth(a, b, c):
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    if pa <= pb and pa <= pc:
        return a
    return b if pb <= pc else c


def read_png(path):
    data = Path(path).read_bytes()
    if not data.startswith(_PNG_SIG):
        raise ValueError("Keine PNG-Datei.")
    pos = 8
    width = height = 0
    bit_depth = color_type = interlace = 0
    palette = b""
    trns = b""
    idat = bytearray()
    dpi = None
    while pos + 8 <= len(data):
        length, ctype = struct.unpack_from(">I4s", data, pos)
        chunk = data[pos + 8:pos + 8 + length]
        pos += 12 + length
        if ctype == b"IHDR":
            width, height, bit_depth, color_type, _comp, _filt, interlace = \
                struct.unpack(">IIBBBBB", chunk)
        elif ctype == b"PLTE":
            palette = chunk
        elif ctype == b"tRNS":
            trns = chunk
        elif ctype == b"pHYs":
            ppx, _ppy, unit = struct.unpack(">IIB", chunk)
            if unit == 1 and ppx:
                dpi = round(ppx * 0.0254)
        elif ctype == b"IDAT":
            idat += chunk
        elif ctype == b"IEND":
            break
    if bit_depth != 8:
        raise ValueError(f"PNG mit Bittiefe {bit_depth} wird nicht unterstützt (nur 8 Bit).")
    if interlace:
        raise ValueError("Interlaced-PNG (Adam7) wird nicht unterstützt.")
    samples = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}.get(color_type)
    if samples is None:
        raise ValueError(f"Unbekannter PNG-Farbtyp {color_type}.")

    raw = zlib.decompress(bytes(idat))
    stride = width * samples
    out = bytearray(height * stride)
    prev = bytearray(stride)
    src = 0
    for row in range(height):
        filter_type = raw[src]
        src += 1
        line = bytearray(raw[src:src + stride])
        src += stride
        if filter_type == 1:
            for i in range(samples, stride):
                line[i] = (line[i] + line[i - samples]) & 0xFF
        elif filter_type == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 0xFF
        elif filter_type == 3:
            for i in range(stride):
                left = line[i - samples] if i >= samples else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif filter_type == 4:
            for i in range(stride):
                left = line[i - samples] if i >= samples else 0
                up_left = prev[i - samples] if i >= samples else 0
                line[i] = (line[i] + _paeth(left, prev[i], up_left)) & 0xFF
        out[row * stride:(row + 1) * stride] = line
        prev = line

    if color_type == 2:
        return Image(width, height, "RGB", out, dpi)
    if color_type == 6:
        return Image(width, height, "RGBA", out, dpi)
    if color_type == 0:
        rgb = bytearray(width * height * 3)
        for i, value in enumerate(out):
            rgb[i * 3:i * 3 + 3] = bytes((value, value, value))
        return Image(width, height, "RGB", rgb, dpi)
    if color_type == 4:
        rgba = bytearray(width * height * 4)
        for i in range(width * height):
            g, a = out[i * 2], out[i * 2 + 1]
            rgba[i * 4:i * 4 + 4] = bytes((g, g, g, a))
        return Image(width, height, "RGBA", rgba, dpi)
    # Palette
    has_alpha = bool(trns)
    channels = 4 if has_alpha else 3
    result = bytearray(width * height * channels)
    for i, index in enumerate(out):
        r, g, b = palette[index * 3:index * 3 + 3]
        if has_alpha:
            a = trns[index] if index < len(trns) else 255
            result[i * 4:i * 4 + 4] = bytes((r, g, b, a))
        else:
            result[i * 3:i * 3 + 3] = bytes((r, g, b))
    return Image(width, height, "RGBA" if has_alpha else "RGB", result, dpi)


def _png_chunk(ctype, payload):
    return (struct.pack(">I", len(payload)) + ctype + payload
            + struct.pack(">I", zlib.crc32(ctype + payload) & 0xFFFFFFFF))


def write_png(image, path):
    color_type = 6 if image.mode == "RGBA" else 2
    samples = image.channels
    stride = image.width * samples
    raw = bytearray()
    for row in range(image.height):
        raw.append(0)
        raw += image.data[row * stride:(row + 1) * stride]
    out = bytearray(_PNG_SIG)
    out += _png_chunk(b"IHDR", struct.pack(">IIBBBBB", image.width, image.height,
                                           8, color_type, 0, 0, 0))
    if image.dpi:
        ppm = round(image.dpi / 0.0254)
        out += _png_chunk(b"pHYs", struct.pack(">IIB", ppm, ppm, 1))
    out += _png_chunk(b"IDAT", zlib.compress(bytes(raw), 8))
    out += _png_chunk(b"IEND", b"")
    Path(path).write_bytes(bytes(out))


# ---------------------------------------------------------------------------
# BMP (24 Bit lesen/schreiben, 32 Bit lesen)
# ---------------------------------------------------------------------------

def read_bmp(path):
    data = Path(path).read_bytes()
    if data[:2] != b"BM":
        raise ValueError("Keine BMP-Datei.")
    pixel_offset = struct.unpack_from("<I", data, 10)[0]
    header_size = struct.unpack_from("<I", data, 14)[0]
    if header_size < 40:
        raise ValueError("BMP-Headervariante wird nicht unterstützt.")
    width, height = struct.unpack_from("<ii", data, 18)
    planes, bpp = struct.unpack_from("<HH", data, 26)
    compression = struct.unpack_from("<I", data, 30)[0]
    ppm_x = struct.unpack_from("<i", data, 38)[0]
    if compression not in (0, 3) or bpp not in (24, 32):
        raise ValueError(f"BMP mit {bpp} Bit/Kompression {compression} wird nicht unterstützt.")
    flip = height > 0
    height = abs(height)
    dpi = round(ppm_x * 0.0254) if ppm_x > 0 else None
    src_stride = ((width * bpp // 8) + 3) & ~3
    mode = "RGBA" if bpp == 32 else "RGB"
    channels = 4 if bpp == 32 else 3
    out = bytearray(width * height * channels)
    for row in range(height):
        src_row = (height - 1 - row) if flip else row
        base = pixel_offset + src_row * src_stride
        for col in range(width):
            offset = base + col * (bpp // 8)
            b, g, r = data[offset], data[offset + 1], data[offset + 2]
            dst = (row * width + col) * channels
            if bpp == 32:
                out[dst:dst + 4] = bytes((r, g, b, data[offset + 3]))
            else:
                out[dst:dst + 3] = bytes((r, g, b))
    return Image(width, height, mode, out, dpi)


def write_bmp(image, path):
    rgb = flatten_to_rgb(image)
    stride = (rgb.width * 3 + 3) & ~3
    ppm = round((rgb.dpi or 96) / 0.0254)
    pixel_data = bytearray()
    for row in range(rgb.height - 1, -1, -1):
        line = bytearray()
        for col in range(rgb.width):
            offset = (row * rgb.width + col) * 3
            r, g, b = rgb.data[offset:offset + 3]
            line += bytes((b, g, r))
        line += b"\0" * (stride - len(line))
        pixel_data += line
    header = struct.pack("<2sIHHI", b"BM", 54 + len(pixel_data), 0, 0, 54)
    info = struct.pack("<IiiHHIIiiII", 40, rgb.width, rgb.height, 1, 24, 0,
                       len(pixel_data), ppm, ppm, 0, 0)
    Path(path).write_bytes(header + info + bytes(pixel_data))


def flatten_to_rgb(image, background=(255, 255, 255)):
    """RGBA auf Hintergrund legen (für BMP/undurchsichtige Ziele)."""
    if image.mode == "RGB":
        return image
    out = bytearray(image.width * image.height * 3)
    for i in range(image.width * image.height):
        r, g, b, a = image.data[i * 4:i * 4 + 4]
        out[i * 3] = (r * a + background[0] * (255 - a)) // 255
        out[i * 3 + 1] = (g * a + background[1] * (255 - a)) // 255
        out[i * 3 + 2] = (b * a + background[2] * (255 - a)) // 255
    return Image(image.width, image.height, "RGB", out, image.dpi)


def read_image(path):
    ext = Path(path).suffix.lower()
    if ext == ".png":
        return read_png(path)
    if ext == ".bmp":
        return read_bmp(path)
    if ext in (".jpg", ".jpeg"):
        return read_jpeg(path)
    if ext == ".gif":
        return read_gif(path)
    if ext in (".tif", ".tiff"):
        return read_tiff(path)
    raise ValueError(f"Kein unterstütztes Bildformat: {ext} (PNG/BMP/JPG/GIF/TIFF).")


def write_image(image, path, jpeg_quality=80, optimize_png=False):
    ext = Path(path).suffix.lower()
    if ext == ".png":
        if not (optimize_png and write_png_palette(image, path)):
            write_png(image, path)
    elif ext == ".bmp":
        write_bmp(image, path)
    elif ext in (".jpg", ".jpeg"):
        write_jpeg(image, path, quality=jpeg_quality)
    else:
        raise ValueError(f"Kein unterstütztes Bild-Zielformat: {ext} (PNG/BMP/JPG).")


# ---------------------------------------------------------------------------
# Operationen
# ---------------------------------------------------------------------------

def crop(image, zone):
    """Zone (x, y, Breite, Höhe in % von links oben) ausschneiden."""
    zx, zy, zw, zh = (max(0.0, min(100.0, float(v))) / 100.0 for v in zone)
    x0 = int(image.width * zx)
    y0 = int(image.height * zy)
    width = max(1, min(int(image.width * zw), image.width - x0))
    height = max(1, min(int(image.height * zh), image.height - y0))
    channels = image.channels
    out = bytearray(width * height * channels)
    for row in range(height):
        src = ((y0 + row) * image.width + x0) * channels
        dst = row * width * channels
        out[dst:dst + width * channels] = image.data[src:src + width * channels]
    return Image(width, height, image.mode, out, image.dpi)


def resize(image, new_width, new_height=None):
    """Bilinear skalieren; new_height leer = proportional."""
    new_width = max(1, int(new_width))
    if new_height is None:
        new_height = max(1, round(image.height * new_width / image.width))
    channels = image.channels
    out = bytearray(new_width * new_height * channels)
    x_ratio = (image.width - 1) / max(1, new_width - 1) if new_width > 1 else 0
    y_ratio = (image.height - 1) / max(1, new_height - 1) if new_height > 1 else 0
    src = image.data
    for y in range(new_height):
        fy = y * y_ratio
        y0 = int(fy)
        y1 = min(y0 + 1, image.height - 1)
        wy = fy - y0
        for x in range(new_width):
            fx = x * x_ratio
            x0 = int(fx)
            x1 = min(x0 + 1, image.width - 1)
            wx = fx - x0
            dst = (y * new_width + x) * channels
            a = (y0 * image.width + x0) * channels
            b = (y0 * image.width + x1) * channels
            c = (y1 * image.width + x0) * channels
            d = (y1 * image.width + x1) * channels
            for ch in range(channels):
                top = src[a + ch] * (1 - wx) + src[b + ch] * wx
                bottom = src[c + ch] * (1 - wx) + src[d + ch] * wx
                out[dst + ch] = int(top * (1 - wy) + bottom * wy)
    return Image(new_width, new_height, image.mode, out, image.dpi)


# 5x7-Pixelfont (5-Bit-Zeilen, MSB links)
_FONT = {
    "A": (0x0E, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11), "B": (0x1E, 0x11, 0x11, 0x1E, 0x11, 0x11, 0x1E),
    "C": (0x0E, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0E), "D": (0x1E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1E),
    "E": (0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x1F), "F": (0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x10),
    "G": (0x0E, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0F), "H": (0x11, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11),
    "I": (0x0E, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0E), "J": (0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0C),
    "K": (0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11), "L": (0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1F),
    "M": (0x11, 0x1B, 0x15, 0x15, 0x11, 0x11, 0x11), "N": (0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11),
    "O": (0x0E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E), "P": (0x1E, 0x11, 0x11, 0x1E, 0x10, 0x10, 0x10),
    "Q": (0x0E, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0D), "R": (0x1E, 0x11, 0x11, 0x1E, 0x14, 0x12, 0x11),
    "S": (0x0F, 0x10, 0x10, 0x0E, 0x01, 0x01, 0x1E), "T": (0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04),
    "U": (0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E), "V": (0x11, 0x11, 0x11, 0x11, 0x11, 0x0A, 0x04),
    "W": (0x11, 0x11, 0x11, 0x15, 0x15, 0x1B, 0x11), "X": (0x11, 0x11, 0x0A, 0x04, 0x0A, 0x11, 0x11),
    "Y": (0x11, 0x11, 0x0A, 0x04, 0x04, 0x04, 0x04), "Z": (0x1F, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1F),
    "0": (0x0E, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0E), "1": (0x04, 0x0C, 0x04, 0x04, 0x04, 0x04, 0x0E),
    "2": (0x0E, 0x11, 0x01, 0x06, 0x08, 0x10, 0x1F), "3": (0x1F, 0x02, 0x04, 0x02, 0x01, 0x11, 0x0E),
    "4": (0x02, 0x06, 0x0A, 0x12, 0x1F, 0x02, 0x02), "5": (0x1F, 0x10, 0x1E, 0x01, 0x01, 0x11, 0x0E),
    "6": (0x06, 0x08, 0x10, 0x1E, 0x11, 0x11, 0x0E), "7": (0x1F, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08),
    "8": (0x0E, 0x11, 0x11, 0x0E, 0x11, 0x11, 0x0E), "9": (0x0E, 0x11, 0x11, 0x0F, 0x01, 0x02, 0x0C),
    " ": (0, 0, 0, 0, 0, 0, 0), "-": (0, 0, 0, 0x1F, 0, 0, 0),
    ".": (0, 0, 0, 0, 0, 0x0C, 0x0C), "/": (0x01, 0x01, 0x02, 0x04, 0x08, 0x10, 0x10),
    ":": (0, 0x0C, 0x0C, 0, 0x0C, 0x0C, 0), "!": (0x04, 0x04, 0x04, 0x04, 0x04, 0, 0x04),
}
_FONT_MAP = {"Ä": "A", "Ö": "O", "Ü": "U", "ß": "S"}


def watermark(image, text, opacity=0.45, color=(90, 90, 90)):
    """Text zentriert als halbtransparentes Wasserzeichen einblenden."""
    text = "".join(_FONT_MAP.get(c, c) for c in text.upper())
    text = "".join(c if c in _FONT else " " for c in text) or "ENTWURF"
    glyph_width = 6                                    # 5 Pixel + 1 Lücke
    scale = max(1, int(image.width * 0.7 / (len(text) * glyph_width)))
    total_width = len(text) * glyph_width * scale
    total_height = 7 * scale
    origin_x = (image.width - total_width) // 2
    origin_y = (image.height - total_height) // 2
    channels = image.channels
    data = image.data
    for index, char in enumerate(text):
        rows = _FONT[char]
        for gy in range(7):
            for gx in range(5):
                if not (rows[gy] >> (4 - gx)) & 1:
                    continue
                for sy in range(scale):
                    y = origin_y + gy * scale + sy
                    if not 0 <= y < image.height:
                        continue
                    for sx in range(scale):
                        x = origin_x + (index * glyph_width + gx) * scale + sx
                        if not 0 <= x < image.width:
                            continue
                        offset = (y * image.width + x) * channels
                        for ch in range(3):
                            old = data[offset + ch]
                            data[offset + ch] = int(old * (1 - opacity) + color[ch] * opacity)
    return image


# ===========================================================================
# JPEG-Codec (Baseline, eigener Encoder/Decoder), GIF- und TIFF-Leser,
# PNG-Optimierung (Palette) und PDF-Konvertierung (Text/Bilder/aus Bildern)
# ===========================================================================

_ZIGZAG = (0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5,
           12, 19, 26, 33, 40, 48, 41, 34, 27, 20, 13, 6, 7, 14, 21, 28,
           35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51,
           58, 59, 52, 45, 38, 31, 39, 46, 53, 60, 61, 54, 47, 55, 62, 63)

_JPEG_Q_LUMA = (16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55,
                14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62,
                18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92,
                49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99)
_JPEG_Q_CHROMA = (17, 18, 24, 47, 99, 99, 99, 99, 18, 21, 26, 66, 99, 99, 99, 99,
                  24, 26, 56, 99, 99, 99, 99, 99, 47, 66, 99, 99, 99, 99, 99, 99,
                  99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99,
                  99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99)

_DC_LUMA_BITS = (0, 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0)
_DC_LUMA_VALS = tuple(range(12))
_DC_CHROMA_BITS = (0, 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0)
_DC_CHROMA_VALS = tuple(range(12))
_AC_LUMA_BITS = (0, 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 125)
_AC_LUMA_VALS = (
    0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07,
    0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08, 0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0,
    0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28,
    0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49,
    0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
    0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
    0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7,
    0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5,
    0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2,
    0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8,
    0xF9, 0xFA)
_AC_CHROMA_BITS = (0, 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 119)
_AC_CHROMA_VALS = (
    0x00, 0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61, 0x71,
    0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xA1, 0xB1, 0xC1, 0x09, 0x23, 0x33, 0x52, 0xF0,
    0x15, 0x62, 0x72, 0xD1, 0x0A, 0x16, 0x24, 0x34, 0xE1, 0x25, 0xF1, 0x17, 0x18, 0x19, 0x1A, 0x26,
    0x27, 0x28, 0x29, 0x2A, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
    0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68,
    0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87,
    0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5,
    0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3,
    0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA,
    0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8,
    0xF9, 0xFA)

_COS = [[math.cos((2 * x + 1) * u * math.pi / 16) for x in range(8)] for u in range(8)]
_CU = [1 / math.sqrt(2)] + [1.0] * 7


def _scale_quant(table, quality):
    quality = max(1, min(95, int(quality)))
    scale = 5000 / quality if quality < 50 else 200 - 2 * quality
    return [max(1, min(255, int((value * scale + 50) // 100))) for value in table]


class _JpegBitWriter:
    def __init__(self):
        self.out = bytearray()
        self.acc = 0
        self.count = 0

    def write(self, value, length):
        self.acc = (self.acc << length) | (value & ((1 << length) - 1))
        self.count += length
        while self.count >= 8:
            self.count -= 8
            byte = (self.acc >> self.count) & 0xFF
            self.out.append(byte)
            if byte == 0xFF:
                self.out.append(0x00)

    def flush(self):
        while self.count % 8:                 # mit 1-Bits bis zur Bytegrenze
            self.write(1, 1)


def _huff_codes(bits, vals):
    codes = {}
    code = 0
    index = 0
    for length in range(1, 17):
        for _ in range(bits[length]):
            codes[vals[index]] = (code, length)
            index += 1
            code += 1
        code <<= 1
    return codes


def _fdct_quant(block, quant):
    shifted = [v - 128.0 for v in block]
    tmp = [0.0] * 64
    for u in range(8):
        cos_u = _COS[u]
        for y in range(8):
            total = 0.0
            base = y * 8
            for x in range(8):
                total += shifted[base + x] * cos_u[x]
            tmp[u * 8 + y] = total
    out = [0] * 64
    for u in range(8):
        for v in range(8):
            cos_v = _COS[v]
            total = 0.0
            for y in range(8):
                total += tmp[u * 8 + y] * cos_v[y]
            value = 0.25 * _CU[u] * _CU[v] * total
            out[v * 8 + u] = int(round(value / quant[v * 8 + u]))
    return out


def write_jpeg(image, path, quality=80):
    """Baseline-JPEG schreiben (4:4:4). quality 1-95."""
    rgb = flatten_to_rgb(image)
    width, height = rgb.width, rgb.height
    data = rgb.data

    q_luma = _scale_quant(_JPEG_Q_LUMA, quality)
    q_chroma = _scale_quant(_JPEG_Q_CHROMA, quality)
    dc_l = _huff_codes(_DC_LUMA_BITS, _DC_LUMA_VALS)
    ac_l = _huff_codes(_AC_LUMA_BITS, _AC_LUMA_VALS)
    dc_c = _huff_codes(_DC_CHROMA_BITS, _DC_CHROMA_VALS)
    ac_c = _huff_codes(_AC_CHROMA_BITS, _AC_CHROMA_VALS)

    blocks_x = (width + 7) // 8
    blocks_y = (height + 7) // 8
    writer = _JpegBitWriter()
    prev_dc = [0, 0, 0]

    def encode_block(block, quant, dc_table, ac_table, comp):
        coeffs = _fdct_quant(block, quant)
        diff = coeffs[0] - prev_dc[comp]
        prev_dc[comp] = coeffs[0]
        magnitude = abs(diff)
        size = magnitude.bit_length()
        code, length = dc_table[size]
        writer.write(code, length)
        if size:
            writer.write(diff if diff > 0 else diff + (1 << size) - 1, size)
        run = 0
        last_nonzero = 0
        zigzagged = [coeffs[_ZIGZAG[k]] for k in range(64)]
        for k in range(63, 0, -1):
            if zigzagged[k]:
                last_nonzero = k
                break
        for k in range(1, last_nonzero + 1):
            value = zigzagged[k]
            if value == 0:
                run += 1
                continue
            while run > 15:
                code, length = ac_table[0xF0]
                writer.write(code, length)
                run -= 16
            size = abs(value).bit_length()
            code, length = ac_table[(run << 4) | size]
            writer.write(code, length)
            writer.write(value if value > 0 else value + (1 << size) - 1, size)
            run = 0
        if last_nonzero < 63:
            code, length = ac_table[0x00]
            writer.write(code, length)

    for by in range(blocks_y):
        for bx in range(blocks_x):
            y_block = [0.0] * 64
            cb_block = [0.0] * 64
            cr_block = [0.0] * 64
            for yy in range(8):
                sy = min(by * 8 + yy, height - 1)
                row_base = sy * width
                for xx in range(8):
                    sx = min(bx * 8 + xx, width - 1)
                    offset = (row_base + sx) * 3
                    r, g, b = data[offset], data[offset + 1], data[offset + 2]
                    idx = yy * 8 + xx
                    y_block[idx] = 0.299 * r + 0.587 * g + 0.114 * b
                    cb_block[idx] = -0.168736 * r - 0.331264 * g + 0.5 * b + 128
                    cr_block[idx] = 0.5 * r - 0.418688 * g - 0.081312 * b + 128
            encode_block(y_block, q_luma, dc_l, ac_l, 0)
            encode_block(cb_block, q_chroma, dc_c, ac_c, 1)
            encode_block(cr_block, q_chroma, dc_c, ac_c, 2)
    writer.flush()

    def segment(marker, payload):
        return bytes((0xFF, marker)) + struct.pack(">H", len(payload) + 2) + payload

    dpi = rgb.dpi or 96
    out = bytearray(b"\xff\xd8")
    out += segment(0xE0, b"JFIF\x00\x01\x01\x01" + struct.pack(">HH", dpi, dpi) + b"\x00\x00")
    out += segment(0xDB, bytes([0]) + bytes(q_luma[_ZIGZAG[k]] for k in range(64)))
    out += segment(0xDB, bytes([1]) + bytes(q_chroma[_ZIGZAG[k]] for k in range(64)))
    out += segment(0xC0, struct.pack(">BHHB", 8, height, width, 3)
                   + bytes((1, 0x11, 0)) + bytes((2, 0x11, 1)) + bytes((3, 0x11, 1)))
    for table_class, table_id, bits, vals in ((0, 0, _DC_LUMA_BITS, _DC_LUMA_VALS),
                                              (1, 0, _AC_LUMA_BITS, _AC_LUMA_VALS),
                                              (0, 1, _DC_CHROMA_BITS, _DC_CHROMA_VALS),
                                              (1, 1, _AC_CHROMA_BITS, _AC_CHROMA_VALS)):
        out += segment(0xC4, bytes([(table_class << 4) | table_id])
                       + bytes(bits[1:]) + bytes(vals))
    out += segment(0xDA, bytes((3, 1, 0x00, 2, 0x11, 3, 0x11, 0, 63, 0)))
    out += writer.out
    out += b"\xff\xd9"
    Path(path).write_bytes(bytes(out))


# -- JPEG lesen -------------------------------------------------------------

class _JpegBitReader:
    def __init__(self, data, pos):
        self.data = data
        self.pos = pos
        self.acc = 0
        self.count = 0
        self.marker = None

    def _fill(self):
        while self.count < 25:
            if self.pos >= len(self.data):
                self.acc = (self.acc << 8) | 0
                self.count += 8
                continue
            byte = self.data[self.pos]
            if byte == 0xFF:
                nxt = self.data[self.pos + 1] if self.pos + 1 < len(self.data) else 0xD9
                if nxt == 0x00:
                    self.pos += 2
                    self.acc = (self.acc << 8) | 0xFF
                    self.count += 8
                    continue
                self.marker = nxt
                self.acc = (self.acc << 8) | 0
                self.count += 8
                continue
            self.pos += 1
            self.acc = (self.acc << 8) | byte
            self.count += 8

    def bit(self):
        if self.count == 0:
            self._fill()
        self.count -= 1
        return (self.acc >> self.count) & 1

    def receive(self, length):
        value = 0
        for _ in range(length):
            value = (value << 1) | self.bit()
        return value

    def restart(self):
        # Byte-Grenze + Restart-Marker ueberspringen
        self.acc = 0
        self.count = 0
        while self.pos + 1 < len(self.data):
            if self.data[self.pos] == 0xFF and 0xD0 <= self.data[self.pos + 1] <= 0xD7:
                self.pos += 2
                self.marker = None
                return
            self.pos += 1


def _extend(value, size):
    return value if value >= (1 << (size - 1)) else value - (1 << size) + 1


def _idct(coeffs):
    tmp = [0.0] * 64
    for y in range(8):
        for u in range(8):
            total = 0.0
            for v in range(8):
                total += _CU[v] * coeffs[v * 8 + u] * _COS[v][y]
            tmp[u * 8 + y] = total
    out = [0] * 64
    for y in range(8):
        for x in range(8):
            total = 0.0
            for u in range(8):
                total += _CU[u] * tmp[u * 8 + y] * _COS[u][x]
            value = int(total * 0.25 + 128.5)
            out[y * 8 + x] = 0 if value < 0 else 255 if value > 255 else value
    return out


def read_jpeg(path):
    """Baseline-JPEG (SOF0/SOF1) dekodieren, inkl. 4:2:0/4:2:2 und Restart-Markern."""
    data = Path(path).read_bytes()
    if data[:2] != b"\xff\xd8":
        raise ValueError("Keine JPEG-Datei.")
    pos = 2
    quant = {}
    huff = {}
    frame = None
    restart_interval = 0
    dpi = None
    while pos < len(data):
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        pos += 2
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            continue
        if marker == 0xD9:
            break
        length = struct.unpack_from(">H", data, pos)[0]
        payload = data[pos + 2:pos + length]
        if marker == 0xDB:
            offset = 0
            while offset < len(payload):
                precision = payload[offset] >> 4
                table_id = payload[offset] & 0x0F
                offset += 1
                if precision:
                    values = list(struct.unpack_from(f">{64}H", payload, offset))
                    offset += 128
                else:
                    values = list(payload[offset:offset + 64])
                    offset += 64
                table = [0] * 64
                for k in range(64):
                    table[_ZIGZAG[k]] = values[k]
                quant[table_id] = table
        elif marker == 0xC4:
            offset = 0
            while offset < len(payload):
                table_class = payload[offset] >> 4
                table_id = payload[offset] & 0x0F
                counts = list(payload[offset + 1:offset + 17])
                total = sum(counts)
                symbols = list(payload[offset + 17:offset + 17 + total])
                offset += 17 + total
                table = {}
                code = 0
                index = 0
                for bit_length in range(1, 17):
                    for _ in range(counts[bit_length - 1]):
                        table[(bit_length, code)] = symbols[index]
                        index += 1
                        code += 1
                    code <<= 1
                huff[(table_class, table_id)] = table
        elif marker in (0xC0, 0xC1):
            precision, height, width, ncomp = struct.unpack_from(">BHHB", payload, 0)
            if precision != 8:
                raise ValueError("Nur 8-Bit-JPEG wird unterstützt.")
            components = []
            for c in range(ncomp):
                cid, sampling, qid = payload[6 + c * 3:9 + c * 3]
                components.append({"id": cid, "h": sampling >> 4, "v": sampling & 0x0F, "q": qid})
            frame = {"w": width, "h": height, "comps": components}
        elif marker == 0xC2:
            raise ValueError("Progressives JPEG wird nicht unterstützt - Datei einmal "
                             "als Baseline speichern (oder PNG anliefern).")
        elif marker in (0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            raise ValueError("Dieses JPEG-Verfahren wird nicht unterstützt (nur Baseline).")
        elif marker == 0xDD:
            restart_interval = struct.unpack_from(">H", payload, 0)[0]
        elif marker == 0xE0 and payload[:5] == b"JFIF\x00":
            unit = payload[7]
            xd = struct.unpack_from(">H", payload, 8)[0]
            if unit == 1 and xd:
                dpi = xd
            elif unit == 2 and xd:
                dpi = round(xd * 2.54)
        elif marker == 0xDA:
            ncomp = payload[0]
            scan = []
            for c in range(ncomp):
                cid, tables = payload[1 + c * 2:3 + c * 2]
                scan.append({"id": cid, "dc": tables >> 4, "ac": tables & 0x0F})
            pos += length
            return _decode_scan(data, pos, frame, scan, quant, huff, restart_interval, dpi)
        pos += length
    raise ValueError("Kein Bildinhalt im JPEG gefunden.")


def _decode_scan(data, pos, frame, scan, quant, huff, restart_interval, dpi):
    if frame is None:
        raise ValueError("JPEG ohne Frame-Header.")
    width, height = frame["w"], frame["h"]
    comps = frame["comps"]
    if len(comps) not in (1, 3):
        raise ValueError("Nur Graustufen- und YCbCr-JPEG werden unterstützt (kein CMYK).")
    hmax = max(c["h"] for c in comps)
    vmax = max(c["v"] for c in comps)
    mcus_x = (width + 8 * hmax - 1) // (8 * hmax)
    mcus_y = (height + 8 * vmax - 1) // (8 * vmax)
    scan_by_id = {s["id"]: s for s in scan}
    planes = []
    for comp in comps:
        pw = mcus_x * comp["h"] * 8
        ph = mcus_y * comp["v"] * 8
        planes.append(bytearray(pw * ph))
        comp["pw"], comp["ph"] = pw, ph

    reader = _JpegBitReader(data, pos)
    prev_dc = [0] * len(comps)

    def decode_huff(table):
        code = 0
        for bit_length in range(1, 17):
            code = (code << 1) | reader.bit()
            symbol = table.get((bit_length, code))
            if symbol is not None:
                return symbol
        raise ValueError("Defekter Huffman-Strom.")

    mcu_index = 0
    for my in range(mcus_y):
        for mx in range(mcus_x):
            if restart_interval and mcu_index and mcu_index % restart_interval == 0:
                reader.restart()
                prev_dc = [0] * len(comps)
            mcu_index += 1
            for ci, comp in enumerate(comps):
                s = scan_by_id.get(comp["id"], scan[ci])
                dc_table = huff[(0, s["dc"])]
                ac_table = huff[(1, s["ac"])]
                qt = quant[comp["q"]]
                for by in range(comp["v"]):
                    for bx in range(comp["h"]):
                        coeffs = [0] * 64
                        size = decode_huff(dc_table)
                        diff = _extend(reader.receive(size), size) if size else 0
                        prev_dc[ci] += diff
                        coeffs[0] = prev_dc[ci] * qt[0]
                        k = 1
                        while k < 64:
                            symbol = decode_huff(ac_table)
                            if symbol == 0x00:
                                break
                            run = symbol >> 4
                            size = symbol & 0x0F
                            if size == 0:
                                if run == 15:
                                    k += 16
                                    continue
                                break
                            k += run
                            if k > 63:
                                break
                            coeffs[_ZIGZAG[k]] = _extend(reader.receive(size), size) * qt[_ZIGZAG[k]]
                            k += 1
                        pixels = _idct(coeffs)
                        plane = planes[ci]
                        pw = comp["pw"]
                        ox = (mx * comp["h"] + bx) * 8
                        oy = (my * comp["v"] + by) * 8
                        for yy in range(8):
                            row = (oy + yy) * pw + ox
                            plane[row:row + 8] = bytes(pixels[yy * 8:yy * 8 + 8])

    out = bytearray(width * height * 3)
    if len(comps) == 1:
        plane = planes[0]
        pw = comps[0]["pw"]
        for y in range(height):
            for x in range(width):
                value = plane[y * pw + x]
                offset = (y * width + x) * 3
                out[offset] = out[offset + 1] = out[offset + 2] = value
        return Image(width, height, "RGB", out, dpi)

    y_plane, cb_plane, cr_plane = planes
    yc, cbc, crc = comps
    for y in range(height):
        y_row = y * yc["pw"]
        cb_row = (y * cbc["v"] // vmax) * cbc["pw"]
        cr_row = (y * crc["v"] // vmax) * crc["pw"]
        for x in range(width):
            lum = y_plane[y_row + x * yc["h"] // hmax]
            cb = cb_plane[cb_row + x * cbc["h"] // hmax] - 128
            cr = cr_plane[cr_row + x * crc["h"] // hmax] - 128
            r = lum + 1.402 * cr
            g = lum - 0.344136 * cb - 0.714136 * cr
            b = lum + 1.772 * cb
            offset = (y * width + x) * 3
            out[offset] = 0 if r < 0 else 255 if r > 255 else int(r)
            out[offset + 1] = 0 if g < 0 else 255 if g > 255 else int(g)
            out[offset + 2] = 0 if b < 0 else 255 if b > 255 else int(b)
    return Image(width, height, "RGB", out, dpi)


def jpeg_size(data):
    """(Breite, Hoehe, Komponenten) aus den JPEG-Markern lesen."""
    pos = 2
    while pos < len(data) - 8:
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        if marker in (0xC0, 0xC1, 0xC2):
            _p, height, width, ncomp = struct.unpack_from(">BHHB", data, pos + 4)
            return width, height, ncomp
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            pos += 2
            continue
        pos += 2 + struct.unpack_from(">H", data, pos + 2)[0]
    raise ValueError("JPEG-Groesse nicht lesbar.")


# -- GIF lesen ----------------------------------------------------------------

def read_gif(path):
    """Erstes Bild einer GIF-Datei dekodieren (LZW, Transparenz)."""
    data = Path(path).read_bytes()
    if data[:3] != b"GIF":
        raise ValueError("Keine GIF-Datei.")
    width, height = struct.unpack_from("<HH", data, 6)
    flags = data[10]
    pos = 13
    global_palette = b""
    if flags & 0x80:
        size = 2 << (flags & 0x07)
        global_palette = data[pos:pos + size * 3]
        pos += size * 3
    transparent = None
    while pos < len(data):
        block = data[pos]
        if block == 0x21:                                # Extension
            label = data[pos + 1]
            pos += 2
            if label == 0xF9 and data[pos] >= 4:
                if data[pos + 1] & 1:
                    transparent = data[pos + 4]
            while data[pos]:
                pos += 1 + data[pos]
            pos += 1
        elif block == 0x2C:                              # Image Descriptor
            ix, iy, iw, ih = struct.unpack_from("<HHHH", data, pos + 1)
            iflags = data[pos + 9]
            pos += 10
            palette = global_palette
            if iflags & 0x80:
                size = 2 << (iflags & 0x07)
                palette = data[pos:pos + size * 3]
                pos += size * 3
            interlaced = bool(iflags & 0x40)
            min_code = data[pos]
            pos += 1
            chunks = []
            while data[pos]:
                count = data[pos]
                chunks.append(data[pos + 1:pos + 1 + count])
                pos += 1 + count
            pos += 1
            indices = _gif_lzw(b"".join(chunks), min_code, iw * ih)
            if interlaced:
                rows = []
                for start, step in ((0, 8), (4, 8), (2, 4), (1, 2)):
                    rows.extend(range(start, ih, step))
                ordered = bytearray(iw * ih)
                for source_row, target_row in enumerate(rows):
                    ordered[target_row * iw:(target_row + 1) * iw] = \
                        indices[source_row * iw:(source_row + 1) * iw]
                indices = ordered
            has_alpha = transparent is not None
            channels = 4 if has_alpha else 3
            out = bytearray(width * height * channels)
            if has_alpha:
                for i in range(0, len(out), 4):
                    out[i + 3] = 255
            for row in range(ih):
                for col in range(iw):
                    idx = indices[row * iw + col]
                    tx, ty = ix + col, iy + row
                    if tx >= width or ty >= height:
                        continue
                    offset = (ty * width + tx) * channels
                    base = idx * 3
                    if base + 2 < len(palette):
                        out[offset:offset + 3] = palette[base:base + 3]
                    if has_alpha:
                        out[offset + 3] = 0 if idx == transparent else 255
            return Image(width, height, "RGBA" if has_alpha else "RGB", out)
        else:
            break
    raise ValueError("Kein Bild in der GIF-Datei gefunden.")


def _gif_lzw(data, min_code, expected):
    clear = 1 << min_code
    end = clear + 1
    table = [bytes([i]) for i in range(clear)] + [b"", b""]
    width = min_code + 1
    out = bytearray()
    acc = 0
    bits = 0
    prev = None
    for byte in data:
        acc |= byte << bits
        bits += 8
        while bits >= width:
            code = acc & ((1 << width) - 1)
            acc >>= width
            bits -= width
            if code == clear:
                table = [bytes([i]) for i in range(clear)] + [b"", b""]
                width = min_code + 1
                prev = None
                continue
            if code == end:
                return out
            if prev is None:
                entry = table[code]
            elif code < len(table):
                entry = table[code]
                table.append(prev + entry[:1])
            else:
                entry = prev + prev[:1]
                table.append(entry)
            out += entry
            prev = entry
            if len(table) >= (1 << width) and width < 12:
                width += 1
            if len(out) >= expected:
                return out
    return out


# -- TIFF lesen ---------------------------------------------------------------

def read_tiff(path):
    """Baseline-TIFF: 8 Bit, unkomprimiert/PackBits/LZW, Grau/RGB/Palette."""
    data = Path(path).read_bytes()
    if data[:2] == b"II":
        endian = "<"
    elif data[:2] == b"MM":
        endian = ">"
    else:
        raise ValueError("Keine TIFF-Datei.")
    if struct.unpack_from(endian + "H", data, 2)[0] != 42:
        raise ValueError("Kein klassisches TIFF (BigTIFF wird nicht unterstützt).")
    ifd = struct.unpack_from(endian + "I", data, 4)[0]
    count = struct.unpack_from(endian + "H", data, ifd)[0]
    tags = {}
    type_sizes = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8}
    for i in range(count):
        base = ifd + 2 + i * 12
        tag, ftype, num = struct.unpack_from(endian + "HHI", data, base)
        size = type_sizes.get(ftype, 1) * num
        offset = base + 8 if size <= 4 else struct.unpack_from(endian + "I", data, base + 8)[0]
        fmt = {1: "B", 3: "H", 4: "I"}.get(ftype)
        if fmt:
            tags[tag] = list(struct.unpack_from(endian + str(num) + fmt, data, offset))

    def tag(tid, default=None):
        values = tags.get(tid)
        return values[0] if values else default

    width = tag(256)
    height = tag(257)
    bits = tags.get(258, [8])
    compression = tag(259, 1)
    photometric = tag(262, 1)
    samples = tag(277, 1)
    rows_per_strip = tag(278, height)
    predictor = tag(317, 1)
    if width is None or height is None:
        raise ValueError("TIFF ohne Bildmaße.")
    if any(b != 8 for b in bits):
        raise ValueError(f"TIFF mit {bits} Bit wird nicht unterstützt (nur 8 Bit).")
    if tag(284, 1) != 1:
        raise ValueError("TIFF mit Planar-Layout wird nicht unterstützt.")
    if compression not in (1, 5, 32773):
        raise ValueError(f"TIFF-Kompression {compression} wird nicht unterstützt "
                         "(nur unkomprimiert, LZW, PackBits).")

    raw = bytearray()
    offsets = tags.get(273, [])
    counts = tags.get(279, [len(data)] * len(offsets))
    for offset, length in zip(offsets, counts):
        chunk = data[offset:offset + length]
        if compression == 32773:
            raw += _packbits(chunk)
        elif compression == 5:
            raw += _tiff_lzw(chunk)
        else:
            raw += chunk
    stride = width * samples
    if predictor == 2:
        for row in range(height):
            base = row * stride
            for i in range(samples, stride):
                raw[base + i] = (raw[base + i] + raw[base + i - samples]) & 0xFF

    if photometric == 3:                                 # Palette
        colormap = tags.get(320, [])
        third = len(colormap) // 3
        out = bytearray(width * height * 3)
        for i in range(width * height):
            idx = raw[i]
            out[i * 3] = colormap[idx] >> 8
            out[i * 3 + 1] = colormap[third + idx] >> 8
            out[i * 3 + 2] = colormap[2 * third + idx] >> 8
        return Image(width, height, "RGB", out)
    if samples == 1:
        out = bytearray(width * height * 3)
        for i in range(width * height):
            value = raw[i] if photometric != 0 else 255 - raw[i]
            out[i * 3] = out[i * 3 + 1] = out[i * 3 + 2] = value
        return Image(width, height, "RGB", out)
    if samples == 3:
        return Image(width, height, "RGB", raw[:width * height * 3])
    if samples == 4:
        return Image(width, height, "RGBA", raw[:width * height * 4])
    raise ValueError(f"TIFF mit {samples} Kanälen wird nicht unterstützt.")


def _packbits(data):
    out = bytearray()
    i = 0
    while i < len(data):
        n = data[i]
        i += 1
        if n < 128:
            out += data[i:i + n + 1]
            i += n + 1
        elif n > 128:
            out += bytes([data[i]]) * (257 - n)
            i += 1
    return out


def _tiff_lzw(data):
    clear, end = 256, 257
    width = 9
    table = [bytes([i]) for i in range(256)] + [b"", b""]
    out = bytearray()
    acc = 0
    bits = 0
    prev = None
    for byte in data:
        acc = (acc << 8) | byte
        bits += 8
        while bits >= width:
            code = (acc >> (bits - width)) & ((1 << width) - 1)
            bits -= width
            if code == clear:
                table = [bytes([i]) for i in range(256)] + [b"", b""]
                width = 9
                prev = None
                continue
            if code == end:
                return out
            if prev is None:
                entry = table[code]
            elif code < len(table):
                entry = table[code]
                table.append(prev + entry[:1])
            else:
                entry = prev + prev[:1]
                table.append(entry)
            out += entry
            prev = entry
            if len(table) >= (1 << width) - 1 and width < 12:  # TIFF: early change
                width += 1
    return out


# -- PNG-Optimierung (Palette) --------------------------------------------------

def write_png_palette(image, path):
    """PNG mit Farbpalette schreiben (deutlich kleiner bei <=256 Farben).

    Liefert True bei Erfolg, False wenn das Bild nicht geeignet ist
    (zu viele Farben oder echte Teiltransparenz).
    """
    rgb = image
    if image.mode == "RGBA":
        if any(image.data[i + 3] != 255 for i in range(0, len(image.data), 4)):
            return False
        rgb = flatten_to_rgb(image)
    colors = {}
    pixel_count = rgb.width * rgb.height
    for i in range(pixel_count):
        key = bytes(rgb.data[i * 3:i * 3 + 3])
        colors[key] = colors.get(key, 0) + 1
        if len(colors) > 4096:
            return False
    palette = _median_cut(colors, 256) if len(colors) > 256 else list(colors)
    lookup = {}
    indices = bytearray(pixel_count)
    for i in range(pixel_count):
        key = bytes(rgb.data[i * 3:i * 3 + 3])
        idx = lookup.get(key)
        if idx is None:
            idx = min(range(len(palette)),
                      key=lambda p: (palette[p][0] - key[0]) ** 2
                      + (palette[p][1] - key[1]) ** 2 + (palette[p][2] - key[2]) ** 2)
            lookup[key] = idx
        indices[i] = idx

    raw = bytearray()
    for row in range(rgb.height):
        raw.append(0)
        raw += indices[row * rgb.width:(row + 1) * rgb.width]
    out = bytearray(_PNG_SIG)
    out += _png_chunk(b"IHDR", struct.pack(">IIBBBBB", rgb.width, rgb.height, 8, 3, 0, 0, 0))
    out += _png_chunk(b"PLTE", b"".join(bytes(c) for c in palette))
    if rgb.dpi:
        ppm = round(rgb.dpi / 0.0254)
        out += _png_chunk(b"pHYs", struct.pack(">IIB", ppm, ppm, 1))
    out += _png_chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    out += _png_chunk(b"IEND", b"")
    Path(path).write_bytes(bytes(out))
    return True


def _median_cut(color_counts, max_colors):
    boxes = [list(color_counts.items())]
    while len(boxes) < max_colors:
        boxes.sort(key=lambda box: -max(
            max(c[0][ch] for c in box) - min(c[0][ch] for c in box) for ch in range(3)))
        box = boxes[0]
        if len(box) <= 1:
            break
        spans = [max(c[0][ch] for c in box) - min(c[0][ch] for c in box) for ch in range(3)]
        channel = spans.index(max(spans))
        box.sort(key=lambda c: c[0][channel])
        middle = len(box) // 2
        boxes[0:1] = [box[:middle], box[middle:]]
    palette = []
    for box in boxes:
        total = sum(count for _c, count in box) or 1
        palette.append(tuple(sum(color[ch] * count for color, count in box) // total
                             for ch in range(3)))
    return palette


# -- PDF-Konvertierung ----------------------------------------------------------

def images_to_pdf(entries, target_path, log=None):
    """Bilder zu einer PDF. entries: Liste aus Path (JPEG wird 1:1 eingebettet)
    oder (name, Image)."""
    writer = PdfWriter()
    page_refs = []
    for entry in entries:
        if isinstance(entry, Path):
            name = entry.name
            ext = entry.suffix.lower()
            if ext in (".jpg", ".jpeg"):
                blob = entry.read_bytes()
                width, height, ncomp = jpeg_size(blob)
                colorspace = PName("DeviceRGB") if ncomp == 3 else PName("DeviceGray")
                xobj = writer.add({"Type": PName("XObject"), "Subtype": PName("Image"),
                                   "Width": width, "Height": height,
                                   "ColorSpace": colorspace, "BitsPerComponent": 8,
                                   "Filter": PName("DCTDecode")}, blob)
                dpi = 96
            else:
                image = read_image(entry)
                name, xobj, width, height, dpi = entry.name, None, image.width, image.height, image.dpi or 96
                entry = (entry.name, image)
        if isinstance(entry, tuple):
            name, image = entry
            rgb = flatten_to_rgb(image)
            width, height = rgb.width, rgb.height
            dpi = rgb.dpi or 96
            xobj = writer.add({"Type": PName("XObject"), "Subtype": PName("Image"),
                               "Width": width, "Height": height,
                               "ColorSpace": PName("DeviceRGB"), "BitsPerComponent": 8,
                               "Filter": PName("FlateDecode")}, zlib.compress(bytes(rgb.data), 6))
        w_pt = width * 72.0 / dpi
        h_pt = height * 72.0 / dpi
        content = writer.add({}, f"q {w_pt:.2f} 0 0 {h_pt:.2f} 0 0 cm /Im0 Do Q".encode("ascii"))
        page = writer.add({"Type": PName("Page"),
                           "MediaBox": [0, 0, round(w_pt, 2), round(h_pt, 2)],
                           "Resources": {"XObject": {"Im0": xobj}},
                           "Contents": content})
        page_refs.append(page)
        if log:
            log(f"SEITE: {name} ({width}×{height})")
    writer.save(target_path, page_refs)
    return len(page_refs)


def _decoded_stream(doc, obj):
    filters = doc.resolve(obj.value.get("Filter"))
    if isinstance(filters, PName) or isinstance(filters, str):
        filters = [filters]
    filters = [str(f) for f in (filters or [])]
    stream = obj.stream or b""
    if not filters:
        return stream
    if filters == ["FlateDecode"]:
        try:
            return zlib.decompress(stream)
        except zlib.error:
            return None
    return None


def pdf_extract_text(path, target_path, log=None):
    """Text aus Content-Streams ziehen (Tj/TJ/'/\"). Encoding ist eine Annäherung -
    bei exotischen Schrift-Encodings kann Zeichensalat entstehen."""
    doc = PdfDocument(path)
    pages = doc.pages()
    lines = []
    for number, (page_ref, _inherited) in enumerate(pages, start=1):
        page = doc.resolve(page_ref)
        contents = page.get("Contents")
        refs = contents if isinstance(contents, list) else [contents]
        stream = b""
        for ref in refs:
            if ref is None:
                continue
            obj = doc.objects.get(int(ref)) if isinstance(ref, PRef) else None
            if obj is not None:
                decoded = _decoded_stream(doc, obj)
                if decoded:
                    stream += decoded + b"\n"
        text = _content_stream_text(stream)
        lines.append(f"----- Seite {number} -----")
        lines.append(text.strip())
        lines.append("")
    result = "\n".join(lines).strip() + "\n"
    Path(target_path).write_text(result, encoding="utf-8")
    if log:
        log(f"TEXT: {len(pages)} Seite(n) → {Path(target_path).name}")
    return len(result)


def _content_stream_text(stream):
    parts = []
    lexer = _Lexer(stream)
    pos = 0
    operands = []
    length = len(stream)
    while pos < length:
        ch = stream[pos:pos + 1]
        if ch in _WHITESPACE:
            pos += 1
            continue
        if ch in (b"(", b"<", b"[", b"/") or ch.isdigit() or ch in (b"-", b"+", b"."):
            lexer.pos = pos
            try:
                value = lexer.parse()
            except PdfError:
                pos += 1
                continue
            pos = lexer.pos
            operands.append(value)
            continue
        match = re.match(rb"[A-Za-z'\"*]+", stream[pos:])
        if not match:
            pos += 1
            continue
        op = match.group(0)
        pos += match.end()
        if op in (b"Tj", b"'", b'"'):
            for value in operands:
                if isinstance(value, bytes):
                    parts.append(value.decode("cp1252", errors="replace"))
            if op in (b"'", b'"'):
                parts.append("\n")
        elif op == b"TJ":
            for value in operands:
                if isinstance(value, list):
                    for item in value:
                        if isinstance(item, bytes):
                            parts.append(item.decode("cp1252", errors="replace"))
        elif op in (b"Td", b"TD", b"T*", b"ET"):
            if parts and not parts[-1].endswith("\n"):
                parts.append("\n")
        operands = []
    return "".join(parts)


def pdf_extract_images(path, target_dir, stem="", log=None):
    """Eingebettete Bilder (DCTDecode -> .jpg, FlateDecode-RGB/Gray -> .png) speichern."""
    doc = PdfDocument(path)
    target_dir = Path(target_dir)
    saved = []
    skipped = 0
    for num, obj in sorted(doc.objects.items()):
        value = obj.value
        if not (isinstance(value, dict) and value.get("Subtype") == "Image" and obj.stream):
            continue
        filters = doc.resolve(value.get("Filter"))
        filters = [str(f) for f in (filters if isinstance(filters, list) else [filters] if filters else [])]
        width = doc.resolve(value.get("Width"))
        height = doc.resolve(value.get("Height"))
        name = f"{stem}_bild{len(saved) + 1:02d}" if stem else f"bild{len(saved) + 1:02d}"
        if "DCTDecode" in filters:
            target = unique_path(target_dir / f"{name}.jpg")
            target.write_bytes(obj.stream)
            saved.append(target)
        elif filters in ([], ["FlateDecode"]) and doc.resolve(value.get("BitsPerComponent", 8)) == 8:
            raw = _decoded_stream(doc, obj)
            colorspace = doc.resolve(value.get("ColorSpace"))
            if raw is None or not isinstance(width, int) or not isinstance(height, int):
                skipped += 1
                continue
            if colorspace == "DeviceRGB" and len(raw) >= width * height * 3:
                image = Image(width, height, "RGB", bytearray(raw[:width * height * 3]))
            elif colorspace == "DeviceGray" and len(raw) >= width * height:
                rgb = bytearray(width * height * 3)
                for i in range(width * height):
                    rgb[i * 3] = rgb[i * 3 + 1] = rgb[i * 3 + 2] = raw[i]
                image = Image(width, height, "RGB", rgb)
            else:
                skipped += 1
                continue
            target = unique_path(target_dir / f"{name}.png")
            write_png(image, target)
            saved.append(target)
        else:
            skipped += 1
    if log:
        note = f", {skipped} übersprungen (Format)" if skipped else ""
        log(f"BILDER: {len(saved)} extrahiert{note}")
    return saved


# ===========================================================================
# Selbsttest (--self-test)
# ===========================================================================

"""Schneller Selbsttest ohne GUI: python data_converter_gui.py --self-test"""




def run_self_test():
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)

        # XLSX-Roundtrip
        xlsx = root / "probe.xlsx"
        write_xlsx(xlsx, ["SKU", "Preis"], [["A-100", "12.4"], ["007", "8"]])
        rows = read_xlsx(xlsx)
        assert rows == [["SKU", "Preis"], ["A-100", "12.4"], ["007", "8"]], rows

        # CSV -> Tabelle -> JSON
        csv_file = root / "probe.csv"
        csv_file.write_text("sku;preis\nA-100;12,40\n", encoding="utf-8")
        table = read_table(csv_file)
        assert table.headers == ["sku", "preis"], table.headers
        transformed, _ = apply_transform(table, TransformOptions(decimal="comma_to_dot"))
        assert transformed.rows[0][1] == "12.40", transformed.rows

        # Umbenennen-Plan
        sample = root / "ALT_liste.csv"
        sample.write_text("x", encoding="utf-8")
        plan = plan_rename([sample], RenameOptions(search="ALT", replace="NEU"))
        assert plan and plan[0][1] == "NEU_liste.csv", plan

    print("Selbsttest OK")


# ===========================================================================
# Oberflaeche, Ketten, Presets, Headless-CLI und Watch-Modus
# ===========================================================================

"""Einkauf Data Converter - GUI.

Arbeitsfluss in jedem Werkzeug-Tab:
    Quelle wählen -> Vorschau (Dry-Run-Plan) -> Ausführen (Live-Log, Fortschritt)

Reine Standardbibliothek (tkinter/ttk), keine Abhängigkeiten.
"""





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
        files = iter_files(
            self.sources,
            recursive=self.recursive_var.get(),
            exts=parse_ext_filter(self.ext_var.get()),
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
        self.plan_consumed = True
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
        self.tree.bind("<Double-1>", lambda _e: self._context_show())
        self.context_menu = tk.Menu(self, tearoff=False)
        self.context_menu.add_command(label="Datei öffnen", command=self._context_open)
        self.context_menu.add_command(label="Im Explorer zeigen", command=self._context_show)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Zeile aus dem Plan nehmen", command=self._context_remove)
        self.tree.bind("<Button-3>", self._context_popup)

    def _selected_index(self):
        selection = self.tree.selection()
        if selection and selection[0] in self.tree_items:
            return self.tree_items.index(selection[0])
        return None

    def _selected_path(self):
        index = self._selected_index()
        if index is None:
            return None
        payload = self.plan[index].payload
        path = payload if isinstance(payload, Path) else \
            payload[0] if isinstance(payload, tuple) and isinstance(payload[0], Path) else None
        return path if path and path.exists() else None

    def _context_popup(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def _context_open(self):
        path = self._selected_path()
        if path:
            os.startfile(str(path))

    def _context_show(self):
        path = self._selected_path()
        if path:
            subprocess.Popen(["explorer", "/select,", str(path)])

    def _context_remove(self):
        index = self._selected_index()
        if index is None:
            return
        old_len = len(self.plan)
        if isinstance(self.run_cfg, dict):          # parallele Listen synchron halten
            for key in ("plan", "dirs", "files"):
                seq = self.run_cfg.get(key)
                if isinstance(seq, list) and len(seq) == old_len:
                    del seq[index]
        self.tree.delete(self.tree_items.pop(index))
        del self.plan[index]
        self.count_var.set(f"{len(self.plan)} Aktion(en) geplant (Zeile entfernt).")

    # -- Vorschau ----------------------------------------------------------

    def refresh_plan(self):
        """Plan neu berechnen und im Treeview anzeigen. Wirft ValueError."""
        self.plan = self.make_plan()
        self.plan_consumed = False
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
        for ext in TABLE_WRITE_EXTS:
            var = tk.BooleanVar(value=(ext == ".xlsx"))
            self.format_vars[ext] = var
            ttk.Checkbutton(row1, text=ext.lstrip(".").upper(), variable=var,
                            style="Card.TCheckbutton").pack(side="left", padx=(8, 0))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        self.delim_combo = _labeled(row2, "CSV-Trennzeichen",
                                    lambda p: ttk.Combobox(p, values=list(DELIMITER_CHOICES), width=5, state="readonly"))
        self.delim_combo.set(";")
        self.enc_combo = _labeled(row2, "Encoding",
                                  lambda p: ttk.Combobox(p, values=ENCODING_CHOICES, width=10, state="readonly"))
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
            "delimiter": DELIMITER_CHOICES[self.delim_combo.get()],
            "encoding": self.enc_combo.get(),
            "options": TransformOptions(
                trim=self.trim_var.get(),
                drop_empty_rows=self.drop_empty_var.get(),
                dedupe=self.dedupe_var.get(),
                decimal=decimal,
                column_spec=self.columns_var.get().strip(),
            ),
            "merge": self.merge_var.get(),
            "merge_name": safe_filename(self.merge_name_var.get().strip() or "zusammengefuehrt"),
        }
        self.log_dir = target_dir
        plan = []
        if self.run_cfg["merge"]:
            source_label = f"{len(files)} Datei(en) zusammengeführt"
            for ext in formats:
                plan.append(PlanItem(source_label, "", str(target_dir / (self.run_cfg["merge_name"] + ext))))
        else:
            for path in files:
                size = format_size(path.stat().st_size)
                for ext in formats:
                    plan.append(PlanItem(path.name, size, str(target_dir / (path.stem + ext)), payload=(path, ext)))
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        done = 0
        total = len(self.plan)

        def load(path):
            table = read_table(path)
            transformed, warnings = apply_transform(table, cfg["options"])
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
            merged = merge_tables(tables)
            ctx.log(f"Zusammengeführt: {len(merged.rows)} Zeilen, {len(merged.headers)} Spalten.")
            for idx, ext in enumerate(cfg["formats"]):
                if ctx.stopped():
                    return
                target = unique_path(cfg["target_dir"] / (cfg["merge_name"] + ext))
                try:
                    write_table(merged, target, cfg["delimiter"], cfg["encoding"])
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
                target = unique_path(cfg["target_dir"] / (path.stem + ext))
                write_table(cache[path], target, cfg["delimiter"], cfg["encoding"])
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
    hint = "STL · OBJ · PLY · 3MF · STEP · IGES → STL · OBJ · PLY · 3MF · GLB · HTML-3D"
    default_exts = "stl, obj, ply, 3mf, step, stp, iges, igs"
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
        ttk.Radiobutton(row1, text="DXF-Zeichnung → SVG", value="dxf",
                        variable=self.mode_var, style="Card.TRadiobutton",
                        command=self._mode_changed).pack(side="left", padx=(16, 0))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Zielformate", style="CardSub.TLabel").pack(side="left")
        self.format_vars = {}
        labels = {".stl": "STL", ".obj": "OBJ", ".ply": "PLY", ".3mf": "3MF",
                  ".glb": "GLB", ".html": "HTML-3D-Ansicht"}
        for ext in MESH_WRITE_EXTS:
            var = tk.BooleanVar(value=(ext in (".stl", ".html")))
            self.format_vars[ext] = var
            ttk.Checkbutton(row2, text=labels[ext], variable=var,
                            style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        self.quality_combo = _labeled(row2, "STEP/IGES-Qualität",
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
        exts = {"konvertieren": "stl, obj, ply, 3mf, step, stp, iges, igs",
                "bericht": "step, stp, iges, igs", "dxf": "dxf"}
        self.source.ext_var.set(exts[self.mode_var.get()])

    def make_plan(self):
        mode = self.mode_var.get()
        files = self.source.collect()
        target_dir = self.target.resolve(self.source.base_dir())
        self.log_dir = target_dir
        if mode == "dxf":
            self.run_cfg = {"mode": mode, "files": files, "target_dir": target_dir}
            return [PlanItem(path.name, format_size(path.stat().st_size),
                             str(target_dir / (path.stem + ".svg")), payload=path) for path in files]
        if mode == "konvertieren":
            formats = [ext for ext, var in self.format_vars.items() if var.get()]
            if not formats:
                raise ValueError("Bitte mindestens ein Zielformat anhaken.")
            self.run_cfg = {"mode": mode, "files": files, "formats": formats,
                            "target_dir": target_dir, "quality": self.quality_combo.get()}
            plan = []
            for path in files:
                size = format_size(path.stat().st_size)
                for ext in formats:
                    plan.append(PlanItem(path.name, size, str(target_dir / (path.stem + ext)),
                                         payload=(path, ext)))
            return plan
        formats = [ext for ext, var in ((".csv", self.report_csv_var), (".html", self.report_html_var)) if var.get()]
        if not formats:
            raise ValueError("Bitte mindestens ein Berichtsformat anhaken (CSV oder HTML).")
        name = safe_filename(self.report_name_var.get().strip() or "cad_bericht")
        self.run_cfg = {"mode": mode, "files": files, "formats": formats,
                        "target_dir": target_dir, "name": name}
        return [PlanItem(path.name, format_size(path.stat().st_size),
                         "wird geprüft", payload=path) for path in files]

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        if cfg["mode"] == "dxf":
            for idx, item in enumerate(self.plan):
                if ctx.stopped():
                    ctx.log("Gestoppt.")
                    return
                path = item.payload
                try:
                    target = unique_path(cfg["target_dir"] / (path.stem + ".svg"))
                    count = dxf_to_svg(path, target)
                    ctx.item(idx, "OK")
                    ctx.log(f"OK: {path.name} → {target.name} ({count} Elemente)", "ok")
                except Exception as exc:
                    ctx.item(idx, "FEHLER")
                    ctx.log(f"FEHLER: {path.name}: {exc}", "err")
                ctx.progress(idx + 1, len(self.plan))
            return
        if cfg["mode"] == "konvertieren":
            cache = {}
            for idx, item in enumerate(self.plan):
                if ctx.stopped():
                    ctx.log("Gestoppt.")
                    return
                path, ext = item.payload
                try:
                    if path not in cache:
                        cache[path] = read_mesh(path, quality=cfg["quality"], log=ctx.log)
                        ctx.log(f"GELESEN: {path.name} ({cache[path].stats()})")
                    target = unique_path(cfg["target_dir"] / (path.stem + ext))
                    write_mesh(cache[path], target)
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
                info = brep_info(path)
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
        table = Table(headers, [[row.get(h, "") for h in headers] for row in rows])
        for ext in cfg["formats"]:
            target = unique_path(cfg["target_dir"] / (cfg["name"] + ext))
            write_table(table, target)
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
# Tab: PDF
# ---------------------------------------------------------------------------

PDF_MODES = {"mergen": "Zusammenführen",
             "splitten": "Splitten (einzeln)",
             "auszug": "Auszug",
             "drehen": "Seiten drehen",
             "umsortieren": "Umsortieren",
             "abdecken": "Zone abdecken",
             "text": "→ Text (TXT)",
             "bilderraus": "Bilder extrahieren"}


class PdfTab(ToolTab):
    key = "pdf"
    title = "PDF"
    hint = "Mergen · Splitten · Zonen abdecken – eigener PDF-Parser, keine Zusatzsoftware"
    default_exts = "pdf"
    target_subfolder = "_pdf"

    def build_options(self, parent):
        row1 = ttk.Frame(parent, style="Card.TFrame")
        row1.pack(fill="x")
        self.mode_var = tk.StringVar(value="mergen")
        row1b = ttk.Frame(parent, style="Card.TFrame")
        for index, (value, label) in enumerate(PDF_MODES.items()):
            parent_row = row1 if index < 6 else row1b
            ttk.Radiobutton(parent_row, text=label, value=value, variable=self.mode_var,
                            style="Card.TRadiobutton").pack(side="left", padx=(0, 12))
        row1b.pack(fill="x", pady=(4, 0))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Name (mergen)", style="CardSub.TLabel").pack(side="left")
        self.bundle_var = tk.StringVar(value="zusammengefuehrt")
        ttk.Entry(row2, textvariable=self.bundle_var, width=20).pack(side="left", padx=(6, 14))
        ttk.Label(row2, text="Seiten (z. B. 1-3,7; leer = alle)", style="CardSub.TLabel").pack(side="left")
        self.range_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.range_var, width=12).pack(side="left", padx=(6, 14))
        self.angle_combo = _labeled(row2, "Drehen um",
                                    lambda p: ttk.Combobox(p, values=["90", "180", "270"], width=5,
                                                           state="readonly"))
        self.angle_combo.set("90")
        ttk.Label(row2, text="Reihenfolge (z. B. 3,1-2)", style="CardSub.TLabel").pack(side="left")
        self.order_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.order_var, width=10).pack(side="left", padx=(6, 14))
        ttk.Label(row2, text="Zone % (x, y, Breite, Höhe von links oben)", style="CardSub.TLabel").pack(side="left")
        self.zone_vars = [tk.StringVar(value=v) for v in ("5", "5", "40", "10")]
        for var in self.zone_vars:
            ttk.Entry(row2, textvariable=var, width=5).pack(side="left", padx=(4, 0))
        self.color_var = tk.StringVar(value="weiss")
        ttk.Radiobutton(row2, text="weiß", value="weiss", variable=self.color_var,
                        style="Card.TRadiobutton").pack(side="left", padx=(12, 4))
        ttk.Radiobutton(row2, text="schwarz", value="schwarz", variable=self.color_var,
                        style="Card.TRadiobutton").pack(side="left")

        row3 = ttk.Frame(parent, style="Card.TFrame")
        row3.pack(fill="x", pady=(6, 0))
        ttk.Label(row3, text="Hinweis: „Abdecken“ legt eine Fläche über den Inhalt – der Text darunter "
                            "bleibt in der Datei extrahierbar (keine echte Schwärzung). Verschlüsselte "
                            "PDFs werden abgelehnt.", style="CardSub.TLabel").pack(side="left")

    def make_plan(self):
        mode = self.mode_var.get()
        files = self.source.collect()
        target_dir = self.target.resolve(self.source.base_dir())
        self.log_dir = target_dir
        zone = []
        for var in self.zone_vars:
            try:
                zone.append(float(var.get().replace(",", ".")))
            except ValueError:
                raise ValueError(f"Zone: „{var.get()}“ ist keine Zahl.")
        self.run_cfg = {"mode": mode, "files": files, "target_dir": target_dir,
                        "bundle": safe_filename(self.bundle_var.get().strip() or "zusammengefuehrt"),
                        "ranges": self.range_var.get().strip(), "zone": zone,
                        "color": self.color_var.get(), "angle": int(self.angle_combo.get()),
                        "order": self.order_var.get().strip()}
        if mode == "mergen":
            if len(files) < 2:
                raise ValueError("Zum Mergen mindestens zwei PDFs wählen.")
            total = sum(p.stat().st_size for p in files)
            return [PlanItem(f"{len(files)} PDFs", format_size(total),
                             str(target_dir / (self.run_cfg["bundle"] + ".pdf")))]
        if mode == "umsortieren" and not self.run_cfg["order"]:
            raise ValueError("Bitte eine Reihenfolge angeben (z. B. 3,1-2).")
        labels = {"splitten": "jede Seite einzeln", "auszug": f"Auszug {self.run_cfg['ranges'] or 'alle'}",
                  "drehen": f"um {self.run_cfg['angle']}° drehen",
                  "umsortieren": f"Reihenfolge {self.run_cfg['order']}",
                  "abdecken": f"Zone abdecken ({self.run_cfg['color']})",
                  "text": "Text extrahieren (TXT)",
                  "bilderraus": "eingebettete Bilder extrahieren"}
        return [PlanItem(path.name, format_size(path.stat().st_size),
                         labels[mode], payload=path) for path in files]

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        if cfg["mode"] == "mergen":
            target = unique_path(cfg["target_dir"] / (cfg["bundle"] + ".pdf"))
            pages = merge_pdfs(cfg["files"], target)
            ctx.item(0, "OK")
            ctx.log(f"MERGE: {target.name} ({pages} Seiten aus {len(cfg['files'])} Dateien)", "ok")
            ctx.progress(1, 1)
            return
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            path = item.payload
            try:
                if cfg["mode"] == "splitten":
                    written = split_pdf(path, cfg["target_dir"], cfg["ranges"],
                                               single_pages=True, log=ctx.log)
                    ctx.item(idx, f"OK ({len(written)})")
                elif cfg["mode"] == "auszug":
                    written = split_pdf(path, cfg["target_dir"], cfg["ranges"],
                                               single_pages=False, log=ctx.log)
                    ctx.item(idx, "OK")
                elif cfg["mode"] == "drehen":
                    target = unique_path(cfg["target_dir"] / (path.stem + "_gedreht.pdf"))
                    rotated = rotate_pdf(path, target, cfg["angle"], cfg["ranges"], log=ctx.log)
                    ctx.item(idx, f"OK ({rotated})")
                    ctx.log(f"OK: {target.name} ({rotated} Seite(n) gedreht)", "ok")
                elif cfg["mode"] == "umsortieren":
                    target = unique_path(cfg["target_dir"] / (path.stem + "_sortiert.pdf"))
                    count = reorder_pdf(path, target, cfg["order"], log=ctx.log)
                    ctx.item(idx, f"OK ({count})")
                elif cfg["mode"] == "text":
                    target = unique_path(cfg["target_dir"] / (path.stem + ".txt"))
                    pdf_extract_text(path, target, log=ctx.log)
                    ctx.item(idx, "OK")
                    ctx.log(f"OK: {path.name} → {target.name}", "ok")
                elif cfg["mode"] == "bilderraus":
                    saved = pdf_extract_images(path, cfg["target_dir"], stem=path.stem, log=ctx.log)
                    ctx.item(idx, f"OK ({len(saved)})")
                else:
                    target = unique_path(cfg["target_dir"] / (path.stem + "_abgedeckt.pdf"))
                    covered = redact_pdf(path, target, cfg["zone"], cfg["color"],
                                                cfg["ranges"], log=ctx.log)
                    ctx.item(idx, f"OK ({covered})")
                    ctx.log(f"OK: {target.name} ({covered} Seite(n) abgedeckt)", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))

    def option_state(self):
        return {"mode": self.mode_var.get(), "bundle": self.bundle_var.get(),
                "ranges": self.range_var.get(), "zone": [v.get() for v in self.zone_vars],
                "color": self.color_var.get(), "angle": self.angle_combo.get(),
                "order": self.order_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "mergen"))
        self.bundle_var.set(state.get("bundle", "zusammengefuehrt"))
        self.range_var.set(state.get("ranges", ""))
        for var, value in zip(self.zone_vars, state.get("zone", ["5", "5", "40", "10"])):
            var.set(value)
        self.color_var.set(state.get("color", "weiss"))
        self.angle_combo.set(state.get("angle", "90"))
        self.order_var.set(state.get("order", ""))


# ---------------------------------------------------------------------------
# Tab: Bilder
# ---------------------------------------------------------------------------

class BilderTab(ToolTab):
    key = "bilder"
    title = "Bilder"
    hint = "PNG · BMP · JPG · GIF · TIFF → PNG · BMP · JPG · PDF – eigene Codecs"
    default_exts = "png, bmp, jpg, jpeg, gif, tif, tiff"
    target_subfolder = "_bilder"

    def build_options(self, parent):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Label(row, text="Zielformate", style="CardSub.TLabel").pack(side="left")
        self.png_var = tk.BooleanVar(value=True)
        self.bmp_var = tk.BooleanVar(value=False)
        self.jpg_var = tk.BooleanVar(value=False)
        self.pdf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="PNG", variable=self.png_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Checkbutton(row, text="BMP", variable=self.bmp_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Checkbutton(row, text="JPG", variable=self.jpg_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 0))
        ttk.Checkbutton(row, text="PDF", variable=self.pdf_var, style="Card.TCheckbutton").pack(side="left", padx=(8, 14))
        ttk.Label(row, text="max. Breite px", style="CardSub.TLabel").pack(side="left")
        self.width_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.width_var, width=7).pack(side="left", padx=(6, 14))
        ttk.Label(row, text="DPI", style="CardSub.TLabel").pack(side="left")
        self.dpi_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.dpi_var, width=6).pack(side="left", padx=(6, 14))
        ttk.Label(row, text="Wasserzeichen", style="CardSub.TLabel").pack(side="left")
        self.mark_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.mark_var, width=16).pack(side="left", padx=(6, 0))

        row2 = ttk.Frame(parent, style="Card.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Komprimierung: JPEG-Qualität", style="CardSub.TLabel").pack(side="left")
        self.quality_var = tk.StringVar(value="80")
        ttk.Spinbox(row2, from_=1, to=95, textvariable=self.quality_var, width=5).pack(side="left", padx=(6, 14))
        self.optimize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="PNG optimieren (Palette, wenn ≤ 256 Farben)",
                        variable=self.optimize_var, style="Card.TCheckbutton").pack(side="left", padx=(0, 14))
        self.pdf_combine_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="PDF: alle Bilder in eine Datei",
                        variable=self.pdf_combine_var, style="Card.TCheckbutton").pack(side="left")

        row3 = ttk.Frame(parent, style="Card.TFrame")
        row3.pack(fill="x", pady=(8, 0))
        ttk.Label(row3, text="Zuschneiden % (x, y, Breite, Höhe von links oben; leer = aus)",
                  style="CardSub.TLabel").pack(side="left")
        self.crop_vars = [tk.StringVar() for _ in range(4)]
        for var in self.crop_vars:
            ttk.Entry(row3, textvariable=var, width=5).pack(side="left", padx=(4, 0))
        ttk.Label(row3, text="Hinweis: progressive JPEGs werden (noch) nicht gelesen; "
                            "große Fotos brauchen etwas Zeit (reiner Python-Codec).",
                  style="CardSub.TLabel").pack(side="left", padx=(14, 0))

    def make_plan(self):
        files = self.source.collect()
        formats = [ext for ext, var in ((".png", self.png_var), (".bmp", self.bmp_var),
                                        (".jpg", self.jpg_var), (".pdf", self.pdf_var)) if var.get()]
        if not formats:
            raise ValueError("Bitte mindestens ein Zielformat anhaken (PNG, BMP, JPG oder PDF).")
        try:
            quality = max(1, min(95, int(self.quality_var.get())))
        except ValueError:
            raise ValueError("JPEG-Qualität muss eine Zahl von 1 bis 95 sein.")
        width = dpi = None
        if self.width_var.get().strip():
            try:
                width = max(1, int(self.width_var.get()))
            except ValueError:
                raise ValueError("Max. Breite muss eine ganze Zahl sein.")
        if self.dpi_var.get().strip():
            try:
                dpi = max(1, int(self.dpi_var.get()))
            except ValueError:
                raise ValueError("DPI muss eine ganze Zahl sein.")
        crop = None
        crop_values = [v.get().strip() for v in self.crop_vars]
        if any(crop_values):
            try:
                crop = [float(v.replace(",", ".")) for v in crop_values]
            except ValueError:
                raise ValueError("Zuschneiden: alle vier Werte als Zahlen angeben (oder alle leer).")
        target_dir = self.target.resolve(self.source.base_dir())
        self.log_dir = target_dir
        combine = self.pdf_combine_var.get()
        self.run_cfg = {"files": files, "formats": formats, "target_dir": target_dir,
                        "width": width, "dpi": dpi, "mark": self.mark_var.get().strip(),
                        "crop": crop, "quality": quality,
                        "optimize": self.optimize_var.get(), "pdf_combine": combine}
        plan = []
        for path in files:
            size = format_size(path.stat().st_size)
            for ext in formats:
                if ext == ".pdf" and combine:
                    continue
                plan.append(PlanItem(path.name, size, str(target_dir / (path.stem + ext)),
                                     payload=(path, ext)))
        if ".pdf" in formats and combine:
            total = sum(p.stat().st_size for p in files)
            plan.append(PlanItem(f"{len(files)} Bild(er)", format_size(total),
                                 str(target_dir / "bilder.pdf")))
        return plan

    def _process(self, path, cfg):
        image = read_image(path)
        if cfg["crop"]:
            image = crop(image, cfg["crop"])
        if cfg["width"] and image.width > cfg["width"]:
            image = resize(image, cfg["width"])
        if cfg["dpi"]:
            image.dpi = cfg["dpi"]
        if cfg["mark"]:
            image = watermark(image, cfg["mark"])
        return image

    def _ops_active(self, cfg):
        return bool(cfg["crop"] or cfg["width"] or cfg["dpi"] or cfg["mark"])

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        cache = {}
        for idx, item in enumerate(self.plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                return
            if item.payload is None:                      # Sammel-PDF am Ende
                try:
                    entries = []
                    for path in cfg["files"]:
                        if not self._ops_active(cfg) and path.suffix.lower() in (".jpg", ".jpeg"):
                            entries.append(path)          # JPEG unveraendert einbetten
                        else:
                            if path not in cache:
                                cache[path] = self._process(path, cfg)
                            entries.append((path.name, cache[path]))
                    target = unique_path(cfg["target_dir"] / "bilder.pdf")
                    pages = images_to_pdf(entries, target, log=ctx.log)
                    ctx.item(idx, f"OK ({pages} Seiten)")
                    ctx.log(f"OK: {target.name} ({pages} Seiten)", "ok")
                except Exception as exc:
                    ctx.item(idx, "FEHLER")
                    ctx.log(f"FEHLER Sammel-PDF: {exc}", "err")
                ctx.progress(idx + 1, len(self.plan))
                continue
            path, ext = item.payload
            try:
                if path not in cache:
                    cache[path] = self._process(path, cfg)
                    ctx.log(f"GELESEN: {path.name} ({cache[path].stats()})")
                target = unique_path(cfg["target_dir"] / (path.stem + ext))
                if ext == ".pdf":
                    images_to_pdf([(path.name, cache[path])], target)
                else:
                    write_image(cache[path], target, jpeg_quality=cfg["quality"],
                                optimize_png=cfg["optimize"])
                ctx.item(idx, "OK")
                ctx.log(f"OK: {path.name} → {target.name}", "ok")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))

    def option_state(self):
        return {"png": self.png_var.get(), "bmp": self.bmp_var.get(),
                "jpg": self.jpg_var.get(), "pdf": self.pdf_var.get(),
                "quality": self.quality_var.get(), "optimize": self.optimize_var.get(),
                "pdf_combine": self.pdf_combine_var.get(),
                "width": self.width_var.get(), "dpi": self.dpi_var.get(),
                "mark": self.mark_var.get(), "crop": [v.get() for v in self.crop_vars]}

    def apply_option_state(self, state):
        self.png_var.set(state.get("png", True))
        self.bmp_var.set(state.get("bmp", False))
        self.jpg_var.set(state.get("jpg", False))
        self.pdf_var.set(state.get("pdf", False))
        self.quality_var.set(state.get("quality", "80"))
        self.optimize_var.set(state.get("optimize", False))
        self.pdf_combine_var.set(state.get("pdf_combine", True))
        self.width_var.set(state.get("width", ""))
        self.dpi_var.set(state.get("dpi", ""))
        self.mark_var.set(state.get("mark", ""))
        for var, value in zip(self.crop_vars, state.get("crop", ["", "", "", ""])):
            var.set(value)


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
        return RenameOptions(
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
        rename_plan = plan_rename(files, self._rename_options())
        if not rename_plan:
            raise ValueError("Die Regeln ändern keinen einzigen Dateinamen – bitte anpassen.")
        self.run_cfg = {"plan": rename_plan}
        self.log_dir = self.source.base_dir()
        return [PlanItem(src.name, format_size(src.stat().st_size), new, payload=(src, new))
                for src, new in rename_plan]

    def execute(self, ctx):
        plan = self.run_cfg["plan"]
        journal = []
        for idx, (source, new_name) in enumerate(plan):
            if ctx.stopped():
                ctx.log("Gestoppt.")
                break
            try:
                target = unique_path(source.with_name(new_name))
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
            journal_path = unique_path(
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
        restored = undo_rename(self.last_journal, log=self.app.log_line)
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
    "zip": "Ein ZIP-Paket",
    "targz": "Ein TAR.GZ-Paket",
    "splitzip": "Mehrere ZIPs mit max. Größe",
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
        ttk.Label(row2, text="max. MB je Teil", style="CardSub.TLabel").pack(side="left", padx=(16, 6))
        self.split_mb_var = tk.StringVar(value="20")
        ttk.Entry(row2, textvariable=self.split_mb_var, width=6).pack(side="left")

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
        bundle_name = safe_filename(self.bundle_var.get().strip() or "Lieferantenpaket")
        try:
            split_mb = max(1, int(self.split_mb_var.get()))
        except ValueError:
            raise ValueError("„max. MB je Teil“ muss eine ganze Zahl sein.")
        self.run_cfg = {"mode": mode, "files": files, "base": base, "target_dir": target_dir,
                        "bundle": bundle_name, "flatten": self.flatten_var.get(),
                        "split_mb": split_mb}
        self.log_dir = target_dir
        plan = []
        if mode == "splitzip":
            total = sum(p.stat().st_size for p in files)
            parts = max(1, -(-total // (split_mb * 1024 * 1024)))
            plan.append(PlanItem(f"{len(files)} Datei(en)", format_size(total),
                                 f"≈ {parts} ZIP-Teil(e) à max. {split_mb} MB → {target_dir}"))
            return plan
        if mode == "einzeln":
            for path in files:
                plan.append(PlanItem(path.name, format_size(path.stat().st_size),
                                     str(target_dir / f"{path.stem}.zip"), payload=path))
        elif mode in ("zip", "targz"):
            ext = ".zip" if mode == "zip" else ".tar.gz"
            total = sum(p.stat().st_size for p in files)
            plan.append(PlanItem(f"{len(files)} Datei(en)", format_size(total),
                                 str(target_dir / (bundle_name + ext))))
        else:
            for path in files:
                plan.append(PlanItem(path.name, format_size(path.stat().st_size),
                                     f"entpacken nach {target_dir}", payload=path))
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        mode = cfg["mode"]
        if mode == "splitzip":
            written = split_zip_bundles(cfg["files"], cfg["target_dir"], cfg["bundle"],
                                                  cfg["split_mb"], log=ctx.log,
                                                  stopped=ctx.stopped, progress=ctx.progress)
            ctx.item(0, f"OK ({len(written)} Teile)")
            return
        if mode in ("zip", "targz"):
            ext = ".zip" if mode == "zip" else ".tar.gz"
            target = cfg["target_dir"] / (cfg["bundle"] + ext)
            bundle_archive(cfg["files"], target, base_dir=cfg["base"], flatten=cfg["flatten"],
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
                    zip_single(path, cfg["target_dir"], log=ctx.log)
                else:
                    extract_archive(path, cfg["target_dir"], flatten=cfg["flatten"], log=ctx.log)
                ctx.item(idx, "OK")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {path.name}: {exc}", "err")
            ctx.progress(idx + 1, len(self.plan))

    def option_state(self):
        return {"mode": self.mode_var.get(), "bundle": self.bundle_var.get(),
                "flatten": self.flatten_var.get(), "split_mb": self.split_mb_var.get()}

    def apply_option_state(self, state):
        self.mode_var.set(state.get("mode", "zip"))
        self.bundle_var.set(state.get("bundle", "Lieferantenpaket"))
        self.flatten_var.set(state.get("flatten", False))
        self.split_mb_var.set(state.get("split_mb", "20"))


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
            empty = find_empty_dirs(base)
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
            move_plan = plan_organize(files, base, mode)
        if not move_plan:
            raise ValueError("Alles liegt schon am richtigen Platz – nichts zu tun.")
        self.run_cfg = {"mode": mode, "plan": move_plan, "copy": self.copy_var.get()}
        verb = "kopieren nach" if self.copy_var.get() else "verschieben nach"
        return [PlanItem(src.name, format_size(src.stat().st_size),
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
                entries = apply_moves([(source, target)], copy=cfg["copy"], log=ctx.log)
                journal.extend(entries)
                ctx.item(idx, "OK")
            except Exception as exc:
                ctx.item(idx, "FEHLER")
                ctx.log(f"FEHLER: {source.name}: {exc}", "err")
            ctx.progress(idx + 1, len(cfg["plan"]))
        self.last_journal = journal
        if journal:
            journal_path = unique_path(
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
        restored = undo_moves(self.last_journal, log=self.app.log_line)
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
        name = safe_filename(self.report_var.get().strip() or "inventar")
        self.run_cfg = {"files": files, "formats": formats, "target_dir": target_dir,
                        "name": name, "hash": self.hash_var.get(), "base": self.source.base_dir(),
                        "sums": self.sums_var.get()}
        self.log_dir = target_dir
        plan = [PlanItem(path.name, format_size(path.stat().st_size), "wird erfasst", payload=path)
                for path in files]
        return plan

    def execute(self, ctx):
        cfg = self.run_cfg

        def progress(done, total):
            ctx.item(done - 1, "OK")
            ctx.progress(done, total)

        manifest = build_manifest(cfg["files"], base_dir=cfg["base"], with_hash=cfg["hash"],
                                            stopped=ctx.stopped, progress=progress, log=ctx.log)
        if ctx.stopped():
            ctx.log("Gestoppt.")
            return
        cfg["target_dir"].mkdir(parents=True, exist_ok=True)
        for ext in cfg["formats"]:
            target = unique_path(cfg["target_dir"] / (cfg["name"] + ext))
            write_table(manifest, target)
            ctx.log(f"BERICHT: {target}", "ok")
        if cfg["sums"]:
            sums_path = unique_path(cfg["target_dir"] / "SHA256SUMS.txt")
            count = write_checksum_file(cfg["files"], cfg["base"], sums_path, stopped=ctx.stopped)
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
            plan.append(PlanItem(path.name, format_size(path.stat().st_size), label, payload=path))
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
                    count = extract_eml_attachments(path, cfg["target_dir"],
                                                              per_mail_subfolder=cfg["subfolder"], log=ctx.log)
                    total_attachments += count
                    ctx.item(idx, f"OK ({count})")
                else:
                    ext = ".html" if cfg["mode"] == "html" else ".txt"
                    target = unique_path(cfg["target_dir"] / (path.stem + ext))
                    if cfg["mode"] == "html":
                        eml_to_html(path, target)
                    else:
                        eml_to_text(path, target)
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
        return [PlanItem(path.name, format_size(path.stat().st_size),
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
                target = unique_path(cfg["target_dir"] / path.name)
                detected = convert_text_file(path, target, cfg["encoding"], cfg["newline"])
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


TAB_CLASSES = [TabellenTab, CadTab, PdfTab, BilderTab, RenameTab, PackTab, OrdnenTab,
               InventarTab, EmlTab, TextTab]


# ---------------------------------------------------------------------------
# Hauptfenster
# ---------------------------------------------------------------------------

SETTINGS_PATH = Path.home() / ".einkauf_data_converter.json"


class DataConverterApp:
    def __init__(self, root, restore_session=True):
        self.root = root
        self.root.title(f"{APP_TITLE} {__version__}")
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.status_var = tk.StringVar(value="Bereit")
        self.recent_presets = []
        self.chain_keys = []

        self._configure_style()
        self._build_menu()
        self._build_ui()
        if restore_session:
            self._load_settings()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(80, self._drain_queue)

    # -- Sitzung merken ----------------------------------------------------

    def _load_settings(self):
        try:
            state = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        for tab in self.tabs:
            if tab.key in state.get("tabs", {}):
                try:
                    tab.set_state(state["tabs"][tab.key])
                except Exception:
                    continue
        index = state.get("active_tab", 0)
        if 0 <= index < len(self.tabs):
            self.notebook.select(index)
        self.recent_presets = [p for p in state.get("recent_presets", []) if Path(p).exists()]
        self.chain_keys = state.get("chain", [])
        self._refresh_preset_menu()

    def _save_settings(self):
        state = {
            "tabs": {tab.key: tab.get_state() for tab in self.tabs},
            "active_tab": self.notebook.index(self.notebook.select()),
            "recent_presets": self.recent_presets[:8],
            "chain": self.chain_keys,
        }
        try:
            SETTINGS_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _on_close(self):
        self._save_settings()
        self.root.destroy()

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
        style.configure("TNotebook.Tab", padding=(11, 8), font=FONT)
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
        file_menu.add_command(label="Kette ausführen…", command=self.chain_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self._on_close)
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
        self.preset_button = ttk.Menubutton(header, text="Presets ▾")
        self.preset_button.pack(side="right", pady=10)
        self.preset_menu = tk.Menu(self.preset_button, tearoff=False)
        self.preset_button.configure(menu=self.preset_menu)
        self._refresh_preset_menu()

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
            # Vorhandener (ggf. per Kontextmenü editierter) Plan zählt; nach einem
            # Lauf oder ohne Vorschau wird neu berechnet.
            if not tab.plan or getattr(tab, "plan_consumed", True):
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
            log_path = unique_path(Path(tab.log_dir) / f"lauf_log_{datetime.now():%Y%m%d_%H%M%S}.txt")
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
            elif kind == "tab":
                self.active_tab = next(t for t in self.tabs if t.key == message[1])
                self.notebook.select(self.tabs.index(self.active_tab))
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
                for tab in self.tabs:
                    tab.plan_consumed = True
                    tab.after_run()
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
            "chain": self.chain_keys,
        }
        Path(path).write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        self.log_line(f"Preset gespeichert: {path}", "ok")
        self._remember_preset(path)

    def load_preset(self):
        path = filedialog.askopenfilename(title="Preset laden", filetypes=[("Preset (JSON)", "*.json")])
        if path:
            self.load_preset_file(path)

    def load_preset_file(self, path):
        try:
            state = json.loads(Path(path).read_text(encoding="utf-8"))
            for tab in self.tabs:
                if tab.key in state.get("tabs", {}):
                    tab.set_state(state["tabs"][tab.key])
            index = state.get("active_tab", 0)
            if 0 <= index < len(self.tabs):
                self.notebook.select(index)
            self.chain_keys = state.get("chain", self.chain_keys)
            self.log_line(f"Preset geladen: {path}", "ok")
            self._remember_preset(path)
            return state
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Preset konnte nicht geladen werden:\n{exc}")
            return None

    def _remember_preset(self, path):
        path = str(Path(path))
        self.recent_presets = [path] + [p for p in self.recent_presets if p != path]
        self.recent_presets = self.recent_presets[:8]
        self._refresh_preset_menu()
        self._save_settings()

    def _refresh_preset_menu(self):
        if not hasattr(self, "preset_menu"):
            return
        self.preset_menu.delete(0, "end")
        if not self.recent_presets:
            self.preset_menu.add_command(label="(noch keine Presets verwendet)", state="disabled")
        for path in self.recent_presets:
            self.preset_menu.add_command(label=Path(path).name,
                                         command=lambda p=path: self.load_preset_file(p))

    # -- Werkzeug-Kette ----------------------------------------------------

    def chain_dialog(self):
        window = tk.Toplevel(self.root)
        window.title("Werkzeug-Kette ausführen")
        window.configure(padx=16, pady=14, background=BG)
        window.resizable(False, False)
        ttk.Label(window, text="Diese Werkzeuge nacheinander ausführen (in Tab-Reihenfolge, "
                               "jedes mit seinen aktuellen Einstellungen):").pack(anchor="w")
        vars_by_key = {}
        for tab in self.tabs:
            var = tk.BooleanVar(value=tab.key in self.chain_keys)
            vars_by_key[tab.key] = var
            ttk.Checkbutton(window, text=tab.title, variable=var).pack(anchor="w", pady=2)

        def start():
            keys = [tab.key for tab in self.tabs if vars_by_key[tab.key].get()]
            window.destroy()
            if keys:
                self.run_chain(keys)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="Abbrechen", command=window.destroy).pack(side="left")
        ttk.Button(buttons, text="Kette ausführen  ▶", style="Accent.TButton",
                   command=start).pack(side="right")

    def run_chain(self, keys):
        if self.worker and self.worker.is_alive():
            return
        self.chain_keys = keys
        tabs_to_run = []
        for key in keys:
            tab = next((t for t in self.tabs if t.key == key), None)
            if tab is None:
                continue
            try:
                tab.refresh_plan()
                tabs_to_run.append(tab)
            except ValueError as exc:
                self.log_line(f"[{tab.title}] übersprungen: {exc}")
        if not tabs_to_run:
            messagebox.showwarning(APP_TITLE, "Kein Werkzeug der Kette hat einen ausführbaren Plan.")
            return

        self.stop_event.clear()
        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_var.set(f"Kette läuft ({len(tabs_to_run)} Werkzeuge) …")
        self.log_line(f"KETTE: {' → '.join(t.title for t in tabs_to_run)}", "ok")

        def worker():
            started = datetime.now()
            try:
                for tab in tabs_to_run:
                    if self.stop_event.is_set():
                        break
                    self.queue.put(("tab", tab.key))
                    log_path = None
                    if tab.log_dir is not None:
                        log_path = unique_path(
                            Path(tab.log_dir) / f"lauf_log_{datetime.now():%Y%m%d_%H%M%S}.txt")
                    ctx = RunContext(self.queue, self.stop_event, log_path)
                    ctx.log(f"--- Kette: {tab.title} ---")
                    try:
                        tab.execute(ctx)
                    except Exception as exc:
                        ctx.log(f"ABBRUCH in {tab.title}: {exc}", "err")
                seconds = (datetime.now() - started).total_seconds()
                self.queue.put(("done", f"Kette fertig in {seconds:.1f} s."))
            finally:
                self.queue.put(("state", "idle"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    # -- Hilfe -------------------------------------------------------------

    def show_capabilities(self):
        messagebox.showinfo(
            "Funktionsumfang",
            "Alles läuft mit reiner Python-Standardbibliothek – ohne Installationen:\n\n"
            "• Tabellen: CSV/TSV/TXT/XLSX/JSON/XML lesen → CSV/TSV/XLSX/JSON/XML/HTML/MD schreiben,\n"
            "  inkl. Dezimalzeichen, Spaltenauswahl, Duplikate, Zusammenführen.\n"
            "• CAD: STL/OBJ/PLY/3MF konvertieren, STEP mit eigenem B-Rep-Kern tessellieren\n"
            "  (→ STL/OBJ/PLY/3MF/GLB/HTML-3D-Ansicht), STEP/IGES-Prüfbericht, DXF → SVG.\n"
            "• PDF: mergen, splitten, drehen, umsortieren, Zonen abdecken, Text und\n"
            "  eingebettete Bilder extrahieren (eigener PDF-Parser).\n"
            "• Bilder: PNG/BMP/JPG/GIF/TIFF lesen → PNG/BMP/JPG/PDF schreiben, eigener\n"
            "  JPEG-Codec mit Qualitätsregler, PNG-Palette-Optimierung, zuschneiden,\n"
            "  skalieren, Wasserzeichen, DPI.\n"
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


class HeadlessContext:
    """RunContext-Ersatz für Läufe ohne Fenster: Konsole + Logdatei."""

    def __init__(self, log_path=None):
        self.errors = 0
        self.log_path = log_path
        if log_path is not None:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.write_text(f"{APP_TITLE} {__version__} (headless)\n"
                                f"Start: {datetime.now():%Y-%m-%d %H:%M:%S}\n\n", encoding="utf-8")

    def stopped(self):
        return False

    def log(self, message, level="info"):
        if level == "err":
            self.errors += 1
        print(f"  {message}")
        if self.log_path is not None:
            with self.log_path.open("a", encoding="utf-8") as handle:
                handle.write(f"{datetime.now():%H:%M:%S}  {message}\n")

    def item(self, index, status):
        pass

    def progress(self, done, total):
        pass


def _headless_app(preset_path):
    root = tk.Tk()
    root.withdraw()
    app = DataConverterApp(root, restore_session=False)
    state = json.loads(Path(preset_path).read_text(encoding="utf-8"))
    for tab in app.tabs:
        if tab.key in state.get("tabs", {}):
            tab.set_state(state["tabs"][tab.key])
    keys = state.get("chain") or []
    if not keys:
        index = state.get("active_tab", 0)
        keys = [app.tabs[index].key if 0 <= index < len(app.tabs) else app.tabs[0].key]
    return root, app, keys


def _headless_run(app, keys):
    errors = 0
    for key in keys:
        tab = next((t for t in app.tabs if t.key == key), None)
        if tab is None:
            print(f"[?] Unbekanntes Werkzeug: {key}")
            continue
        try:
            plan = tab.refresh_plan()
        except ValueError as exc:
            print(f"[{tab.title}] übersprungen: {exc}")
            continue
        print(f"[{tab.title}] {len(plan)} Aktion(en) …")
        log_path = None
        if tab.log_dir is not None:
            log_path = unique_path(
                Path(tab.log_dir) / f"lauf_log_{datetime.now():%Y%m%d_%H%M%S}.txt")
        ctx = HeadlessContext(log_path)
        try:
            tab.execute(ctx)
        except Exception as exc:
            print(f"[{tab.title}] ABBRUCH: {exc}")
            errors += 1
            continue
        errors += 1 if ctx.errors else 0
    return errors


def run_headless(preset_path, watch_interval=None):
    """Preset ohne Fenster ausführen; optional als Watch-Ordner-Schleife."""
    root, app, keys = _headless_app(preset_path)
    try:
        if watch_interval is None:
            errors = _headless_run(app, keys)
            print("Fertig." if not errors else f"Fertig mit {errors} Fehler-Werkzeug(en).")
            return 1 if errors else 0

        import time
        first_tab = next((t for t in app.tabs if t.key == keys[0]), app.tabs[0])
        print(f"Watch-Modus: prüfe alle {watch_interval} s auf neue Dateien "
              f"({first_tab.title}). Beenden mit Strg+C.")
        seen = {}
        pending = None
        while True:
            try:
                files = first_tab.source.collect()
            except ValueError:
                files = []
            snapshot = {}
            for path in files:
                stat = path.stat()
                snapshot[str(path)] = (stat.st_size, int(stat.st_mtime))
            changed = {k: v for k, v in snapshot.items() if seen.get(k) != v}
            if changed and changed == pending:      # zwei Polls stabil -> loslegen
                print(f"{datetime.now():%H:%M:%S} {len(changed)} neue/geänderte Datei(en) "
                      "- Kette startet.")
                _headless_run(app, keys)
                seen = snapshot
                pending = None
            elif changed:
                pending = changed
            else:
                pending = None
            time.sleep(watch_interval)
    except KeyboardInterrupt:
        print("\nWatch-Modus beendet.")
        return 0
    finally:
        root.destroy()


def main(argv=None):
    argv = list(argv or [])
    if "--self-test" in argv:
        run_self_test()
        return
    preset = None
    if "--preset" in argv:
        index = argv.index("--preset")
        if index + 1 < len(argv):
            preset = argv[index + 1]
    if preset and ("--run" in argv or "--watch" in argv):
        watch_interval = None
        if "--watch" in argv:
            index = argv.index("--watch")
            watch_interval = 30
            if index + 1 < len(argv) and argv[index + 1].isdigit():
                watch_interval = max(5, int(argv[index + 1]))
        sys.exit(run_headless(preset, watch_interval))
    _enable_windows_dpi()
    root = tk.Tk()
    root.geometry("1280x780")
    root.minsize(1040, 640)
    DataConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main(sys.argv[1:])
