"""Tabellen lesen, aufbereiten und schreiben - reine Standardbibliothek.

Zentrales Modell ist `Table`: eine Kopfzeile plus Datenzeilen (Strings).
Damit lassen sich Preislisten, Stuecklisten und Lieferantendaten zwischen
CSV/TSV/TXT, JSON, XML und XLSX wandeln und dabei normalisieren.
"""

import csv
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from . import xlsx_io

READ_EXTS = [".csv", ".tsv", ".txt", ".xlsx", ".json", ".xml"]
WRITE_EXTS = [".csv", ".tsv", ".xlsx", ".json", ".xml", ".html", ".md"]
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
        rows = xlsx_io.read_xlsx(path)
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
        xlsx_io.write_xlsx(path, table.headers, table.rows, sheet_name=path.stem)
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
