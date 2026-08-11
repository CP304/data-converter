"""Datei-Werkzeuge: Scannen, Umbenennen (mit Undo), Archive, Inventar,
E-Mail-Anhaenge und Text-Encoding - reine Standardbibliothek."""

import base64
import bz2
import gzip
import hashlib
import json
import lzma
import re
import shutil
import tarfile
import zipfile
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.parser import BytesParser
from html import unescape as html_unescape
from pathlib import Path

from .tabular import Table, decode_bytes

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
