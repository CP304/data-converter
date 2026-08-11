"""PDF-Werkzeuge auf Basis der Standardbibliothek: mergen, splitten, Zonen abdecken.

Eigener Mini-Parser/-Writer für das PDF-Objektmodell (COS): Objekte werden per
Scan eingesammelt (robust auch bei kaputten xref-Tabellen), komprimierte
Objekt-Streams (/ObjStm, Flate) werden entpackt. Beim Schreiben wird der
Objektgraph der benötigten Seiten kopiert und neu nummeriert.

Grenzen: verschlüsselte PDFs werden abgelehnt; „Abdecken" legt eine deckende
Fläche ÜBER den Inhalt (der Text darunter bleibt technisch extrahierbar -
für echte Schwärzung Dokument neu erzeugen lassen).
"""

import re
import zlib
from pathlib import Path


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
