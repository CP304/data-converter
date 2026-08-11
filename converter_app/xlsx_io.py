"""Minimaler XLSX-Reader und -Writer auf Basis der Standardbibliothek.

XLSX ist ein ZIP-Container mit XML-Dateien. Fuer einfache Datentabellen
(eine Arbeitsmappe, ein Blatt, Text und Zahlen) reicht zipfile + ElementTree.
"""

import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

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
