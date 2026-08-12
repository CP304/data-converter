"""Backend-Tests ohne GUI: python test_converter_backend.py"""

import struct
import tempfile
import zipfile
from pathlib import Path

# Alles lebt in einer Datei - die alten Modulnamen zeigen auf dasselbe Modul.
import data_converter_gui as _app

cad_io = filetools = iges_mesh = img_io = pdf_io = step_mesh = tabular = xlsx_io = _app


def test_xlsx_roundtrip(root):
    path = root / "preise.xlsx"
    headers = ["SKU", "Lieferant", "Preis", "Bemerkung"]
    rows = [
        ["A-100", "Müller GmbH", "12.4", "mit Umlaut äöü"],
        ["007", "Schmidt AG", "8", "führende Null bleibt Text"],
        ["C-300", "Demo & Co <>", "-44.9", 'Sonderzeichen "&<>"'],
    ]
    xlsx_io.write_xlsx(path, headers, rows)
    result = xlsx_io.read_xlsx(path)
    assert result[0] == headers, result[0]
    assert result[1:] == rows, result[1:]


def test_csv_json_xlsx_chain(root):
    csv_file = root / "liste.csv"
    csv_file.write_text("sku;supplier;preis\nA-100;Müller GmbH;1.234,50\nB-200;Schmidt AG;8,10\n", encoding="cp1252")
    table = tabular.read_table(csv_file)
    assert table.headers == ["sku", "supplier", "preis"]
    assert table.rows[0][1] == "Müller GmbH"

    transformed, warnings = tabular.apply_transform(
        table, tabular.TransformOptions(decimal="comma_to_dot", column_spec="sku; preis>Netto"))
    assert transformed.headers == ["sku", "Netto"], transformed.headers
    assert transformed.rows[0] == ["A-100", "1234.50"], transformed.rows[0]
    assert not warnings, warnings

    for ext in (".json", ".xlsx", ".tsv", ".xml", ".html", ".md"):
        target = root / ("kette" + ext)
        tabular.write_table(transformed, target)
        assert target.exists() and target.stat().st_size > 0, target

    back = tabular.read_table(root / "kette.json")
    assert back.headers == ["sku", "Netto"]
    back_xlsx = tabular.read_table(root / "kette.xlsx")
    assert back_xlsx.rows[0] == ["A-100", "1234.50"], back_xlsx.rows


def test_dedupe_and_merge(root):
    a = tabular.Table(["sku", "preis"], [["A", "1"], ["A", "1"], ["B", "2"]])
    deduped, warnings = tabular.apply_transform(a, tabular.TransformOptions(dedupe=True))
    assert len(deduped.rows) == 2 and any("Duplikat" in w for w in warnings)

    b = tabular.Table(["sku", "lieferzeit"], [["C", "5"]])
    merged = tabular.merge_tables([("a.csv", a), ("b.csv", b)])
    assert merged.headers == ["Quelle", "sku", "preis", "lieferzeit"]
    assert merged.rows[0] == ["a.csv", "A", "1", ""]
    assert merged.rows[-1] == ["b.csv", "C", "", "5"]


def test_rename_plan_apply_undo(root):
    folder = root / "rename"
    folder.mkdir()
    files = []
    for name in ("ALT_angebot.pdf", "ALT_zeichnung.pdf", "sonstiges.txt"):
        path = folder / name
        path.write_text("x", encoding="utf-8")
        files.append(path)

    options = filetools.RenameOptions(search="ALT", replace="NEU", prefix="PO_", numbering=True)
    plan = filetools.plan_rename(files, options)
    names = [new for _src, new in plan]
    assert names == ["001_PO_NEU_angebot.pdf", "002_PO_NEU_zeichnung.pdf", "003_PO_sonstiges.txt"], names

    journal, journal_path = filetools.apply_rename(plan, log=lambda m: None)
    assert len(journal) == 3 and journal_path.exists()
    assert (folder / "001_PO_NEU_angebot.pdf").exists()

    restored = filetools.undo_rename(journal, log=lambda m: None)
    assert restored == 3
    assert (folder / "ALT_angebot.pdf").exists() and (folder / "sonstiges.txt").exists()


def test_archive_roundtrip(root):
    src = root / "paket"
    src.mkdir()
    (src / "a.txt").write_text("A", encoding="utf-8")
    sub = src / "unter"
    sub.mkdir()
    (sub / "b.txt").write_text("B", encoding="utf-8")

    files = filetools.iter_files([src], recursive=True)
    bundle = filetools.bundle_archive(files, root / "Lieferantenpaket.zip", base_dir=src, log=lambda m: None)
    with zipfile.ZipFile(bundle) as archive:
        names = set(archive.namelist())
    assert "a.txt" in names and "unter/b.txt" in names, names

    out = root / "entpackt"
    count = filetools.extract_archive(bundle, out, log=lambda m: None)
    assert count == 2 and (out / "unter" / "b.txt").read_text(encoding="utf-8") == "B"

    flat = root / "flach"
    filetools.extract_archive(bundle, flat, flatten=True, log=lambda m: None)
    assert (flat / "b.txt").exists()


def test_manifest(root):
    folder = root / "inventar"
    folder.mkdir()
    (folder / "eins.txt").write_text("inhalt", encoding="utf-8")
    (folder / "zwei.txt").write_text("inhalt", encoding="utf-8")  # identisch -> Duplikat
    (folder / "drei.txt").write_text("anders", encoding="utf-8")

    files = filetools.iter_files([folder], recursive=True)
    manifest = filetools.build_manifest(files, base_dir=folder, log=lambda m: None)
    assert manifest.headers[:2] == ["Datei", "Typ"]
    dup_index = manifest.headers.index("Duplikat")
    dup_values = sorted(row[dup_index] for row in manifest.rows)
    assert dup_values == ["", "JA", "JA"], dup_values

    tabular.write_table(manifest, root / "inventar.html")
    assert "SHA-256" in (root / "inventar.html").read_text(encoding="utf-8")


def test_eml_extraction(root):
    eml = root / "angebot.eml"
    eml.write_bytes(
        b"From: lieferant@example.com\r\n"
        b"To: einkauf@example.com\r\n"
        b"Subject: Angebot\r\n"
        b"MIME-Version: 1.0\r\n"
        b"Content-Type: multipart/mixed; boundary=GRENZE\r\n\r\n"
        b"--GRENZE\r\n"
        b"Content-Type: text/plain\r\n\r\n"
        b"Siehe Anhang.\r\n"
        b"--GRENZE\r\n"
        b"Content-Type: application/octet-stream; name=preisliste.csv\r\n"
        b"Content-Disposition: attachment; filename=preisliste.csv\r\n"
        b"Content-Transfer-Encoding: base64\r\n\r\n"
        b"c2t1O3ByZWlzCkEtMTAwOzEyLDQwCg==\r\n"
        b"--GRENZE--\r\n"
    )
    out = root / "anhaenge"
    count = filetools.extract_eml_attachments(eml, out, per_mail_subfolder=True, log=lambda m: None)
    saved = out / "angebot" / "preisliste.csv"
    assert count == 1 and saved.exists()
    assert "A-100" in saved.read_text(encoding="utf-8")


def test_text_conversion(root):
    src = root / "cp1252.txt"
    src.write_bytes("Bestellung über 5 Stück\r\n".encode("cp1252"))
    target = root / "utf8.txt"
    detected = filetools.convert_text_file(src, target, "utf-8", "lf")
    assert detected == "cp1252"
    assert target.read_bytes() == "Bestellung über 5 Stück\n".encode("utf-8")


_STEP_FIXTURE = """ISO-10303-21;
HEADER;
FILE_DESCRIPTION(('Test'),'2;1');
FILE_NAME('test.step','2026-08-11',('Einkauf'),('Firma'),'','','');
FILE_SCHEMA(('AUTOMOTIVE_DESIGN { 1 0 10303 214 1 1 1 1 }'));
ENDSEC;
DATA;
#1=CARTESIAN_POINT('',(0.,0.,0.));
#2=DIRECTION('',(0.,0.,1.));
#3=DIRECTION('',(1.,0.,0.));
#4=AXIS2_PLACEMENT_3D('',#1,#2,#3);
#5=PLANE('',#4);
#10=CARTESIAN_POINT('',(0.,0.,0.));
#11=CARTESIAN_POINT('',(10.,0.,0.));
#12=CARTESIAN_POINT('',(10.,20.,0.));
#13=CARTESIAN_POINT('',(0.,20.,0.));
#14=VERTEX_POINT('',#10);
#15=VERTEX_POINT('',#11);
#16=VERTEX_POINT('',#12);
#17=VERTEX_POINT('',#13);
#20=DIRECTION('',(1.,0.,0.));
#21=VECTOR('',#20,1.);
#22=LINE('',#10,#21);
#30=EDGE_CURVE('',#14,#15,#22,.T.);
#31=EDGE_CURVE('',#15,#16,#22,.T.);
#32=EDGE_CURVE('',#16,#17,#22,.T.);
#33=EDGE_CURVE('',#17,#14,#22,.T.);
#40=ORIENTED_EDGE('',*,*,#30,.T.);
#41=ORIENTED_EDGE('',*,*,#31,.T.);
#42=ORIENTED_EDGE('',*,*,#32,.T.);
#43=ORIENTED_EDGE('',*,*,#33,.T.);
#50=EDGE_LOOP('',(#40,#41,#42,#43));
#51=FACE_OUTER_BOUND('',#50,.T.);
#52=ADVANCED_FACE('',(#51),#5,.T.);
#100=CARTESIAN_POINT('',(30.,0.,0.));
#104=AXIS2_PLACEMENT_3D('',#100,#2,#3);
#105=CYLINDRICAL_SURFACE('',#104,5.);
#110=CARTESIAN_POINT('',(35.,0.,0.));
#111=VERTEX_POINT('',#110);
#113=CIRCLE('',#104,5.);
#114=EDGE_CURVE('',#111,#111,#113,.T.);
#115=ORIENTED_EDGE('',*,*,#114,.T.);
#116=EDGE_LOOP('',(#115));
#117=FACE_BOUND('',#116,.T.);
#120=CARTESIAN_POINT('',(35.,0.,10.));
#121=VERTEX_POINT('',#120);
#122=CARTESIAN_POINT('',(30.,0.,10.));
#123=AXIS2_PLACEMENT_3D('',#122,#2,#3);
#124=CIRCLE('',#123,5.);
#125=EDGE_CURVE('',#121,#121,#124,.T.);
#126=ORIENTED_EDGE('',*,*,#125,.F.);
#127=EDGE_LOOP('',(#126));
#128=FACE_BOUND('',#127,.T.);
#130=ADVANCED_FACE('',(#117,#128),#105,.T.);
#200=PRODUCT('Testteil','Testteil','',());
ENDSEC;
END-ISO-10303-21;
"""


def test_step_kernel(root):
    import math
    step = root / "teil.step"
    step.write_text(_STEP_FIXTURE, encoding="ascii")
    mesh = step_mesh.read_step_mesh(step, quality="mittel")
    assert len(mesh.faces) > 50, len(mesh.faces)
    zs = [v[2] for v in mesh.vertices]
    assert min(zs) == 0 and max(zs) == 10
    radii = [math.hypot(v[0] - 30, v[1]) for v in mesh.vertices if v[0] > 20]
    assert radii and all(abs(r - 5) < 0.01 for r in radii), (min(radii), max(radii))
    assert any(abs(v[0] - 10) < 1e-6 and abs(v[1] - 20) < 1e-6 for v in mesh.vertices)

    info = cad_io.step_info(step)
    assert info["Schema"].startswith("AUTOMOTIVE_DESIGN")
    assert "Testteil" in info["Produkte"]
    assert int(info["Entitäten"]) > 20


def test_mesh_roundtrip(root):
    step = root / "teil.step"
    step.write_text(_STEP_FIXTURE, encoding="ascii")
    mesh = cad_io.read_mesh(step, quality="grob")
    for ext in (".stl", ".obj", ".ply", ".3mf", ".glb", ".html"):
        target = root / ("mesh" + ext)
        cad_io.write_mesh(mesh, target)
        assert target.exists() and target.stat().st_size > 0, target
    for ext in (".stl", ".obj", ".ply", ".3mf"):
        back = cad_io.read_mesh(root / ("mesh" + ext))
        assert len(back.faces) == len(mesh.faces), (ext, len(back.faces), len(mesh.faces))
    glb = (root / "mesh.glb").read_bytes()
    assert glb[:4] == b"glTF"
    html = (root / "mesh.html").read_text(encoding="utf-8")
    assert "webgl2" in html and "Dreiecke" in html


def test_organize_and_undo(root):
    folder = root / "chaos"
    folder.mkdir()
    (folder / "preise.csv").write_text("x", encoding="utf-8")
    (folder / "teil.step").write_text("x", encoding="utf-8")
    (folder / "foto.png").write_bytes(b"x")
    files = filetools.iter_files([folder], recursive=False)
    plan = filetools.plan_organize(files, folder, "typ")
    assert len(plan) == 3
    journal = filetools.apply_moves(plan, log=lambda m: None)
    assert (folder / "Tabellen" / "preise.csv").exists()
    assert (folder / "CAD" / "teil.step").exists()
    assert (folder / "Bilder" / "foto.png").exists()
    restored = filetools.undo_moves(journal, log=lambda m: None)
    assert restored == 3 and (folder / "preise.csv").exists()
    empty = filetools.find_empty_dirs(folder)
    assert {e.name for e in empty} == {"Tabellen", "CAD", "Bilder"}
    for e in empty:
        e.rmdir()


def test_checksum_file(root):
    folder = root / "paket"
    folder.mkdir()
    (folder / "a.txt").write_text("A", encoding="utf-8")
    (folder / "b.txt").write_text("B", encoding="utf-8")
    files = filetools.iter_files([folder], recursive=True)
    sums = root / "SHA256SUMS.txt"
    count = filetools.write_checksum_file(files, folder, sums)
    text = sums.read_text(encoding="utf-8")
    assert count == 2 and "a.txt" in text and len(text.splitlines()[0].split()[0]) == 64


def test_eml_conversion(root):
    eml = root / "angebot.eml"
    eml.write_bytes(
        b"From: lieferant@example.com\r\n"
        b"To: einkauf@example.com\r\n"
        b"Subject: Angebot 4711\r\n"
        b"MIME-Version: 1.0\r\n"
        b"Content-Type: multipart/mixed; boundary=GRENZE\r\n\r\n"
        b"--GRENZE\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n\r\n"
        b"Preis: 12,40 EUR\r\n"
        b"--GRENZE\r\n"
        b"Content-Type: application/pdf; name=angebot.pdf\r\n"
        b"Content-Disposition: attachment; filename=angebot.pdf\r\n"
        b"Content-Transfer-Encoding: base64\r\n\r\n"
        b"JVBERi0=\r\n"
        b"--GRENZE--\r\n"
    )
    filetools.eml_to_html(eml, root / "angebot.html")
    html = (root / "angebot.html").read_text(encoding="utf-8")
    assert "Angebot 4711" in html and "angebot.pdf" in html and "12,40" in html
    filetools.eml_to_text(eml, root / "angebot.txt")
    text = (root / "angebot.txt").read_text(encoding="utf-8")
    assert "Betreff: Angebot 4711" in text and "12,40" in text and "angebot.pdf" in text


_MINI_PDF = (b"%PDF-1.4\n"
             b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
             b"2 0 obj\n<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 /MediaBox [0 0 595 842] >>\nendobj\n"
             b"3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>\nendobj\n"
             b"4 0 obj\n<< /Length 46 >>\nstream\n"
             b"BT /F1 24 Tf 72 770 Td (Angebot 4711) Tj ET\n"
             b"endstream\nendobj\n"
             b"5 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>\nendobj\n"
             b"6 0 obj\n<< /Length 44 >>\nstream\n"
             b"BT /F1 24 Tf 72 770 Td (Seite ZWEI) Tj ET\n"
             b"endstream\nendobj\n"
             b"trailer\n<< /Size 7 /Root 1 0 R >>\n%%EOF\n")


def test_pdf_tools(root):
    src = root / "quelle.pdf"
    src.write_bytes(_MINI_PDF)
    assert pdf_io.pdf_page_count(src) == 2

    merged = root / "merged.pdf"
    assert pdf_io.merge_pdfs([src, src], merged) == 4
    assert pdf_io.pdf_page_count(merged) == 4

    files = pdf_io.split_pdf(merged, root, ranges_text="2-3", single_pages=True)
    assert len(files) == 2 and all(pdf_io.pdf_page_count(f) == 1 for f in files)
    auszug = pdf_io.split_pdf(merged, root, ranges_text="1,4", single_pages=False)
    assert pdf_io.pdf_page_count(auszug[0]) == 2

    red = root / "abgedeckt.pdf"
    covered = pdf_io.redact_pdf(src, red, zone=(10, 5, 50, 10), color="schwarz", ranges_text="1")
    assert covered == 1 and pdf_io.pdf_page_count(red) == 2
    assert b"0 0 0 rg" in red.read_bytes()


def test_image_tools(root):
    image = img_io.Image(64, 40, "RGB", bytearray([220, 80, 40] * (64 * 40)), dpi=300)
    img_io.write_png(image, root / "probe.png")
    back = img_io.read_png(root / "probe.png")
    assert (back.width, back.height, back.mode, back.dpi) == (64, 40, "RGB", 300)
    assert bytes(back.data[:3]) == bytes([220, 80, 40])

    small = img_io.resize(back, 32)
    assert (small.width, small.height) == (32, 20)      # proportional: 40*32/64
    before = bytes(small.data)
    img_io.watermark(small, "GEPRÜFT 2026")
    assert bytes(small.data) != before

    img_io.write_bmp(small, root / "probe.bmp")
    bmp = img_io.read_bmp(root / "probe.bmp")
    assert (bmp.width, bmp.height, bmp.mode) == (32, 20, "RGB")
    img_io.write_png(bmp, root / "rueck.png")
    assert img_io.read_png(root / "rueck.png").width == 32


def test_dxf_to_svg(root):
    dxf_lines = ["0", "SECTION", "2", "ENTITIES",
                 "0", "LINE", "10", "0", "20", "0", "11", "100", "21", "0",
                 "0", "CIRCLE", "10", "50", "20", "30", "40", "10",
                 "0", "ARC", "10", "0", "20", "0", "40", "20", "50", "0", "51", "90",
                 "0", "LWPOLYLINE", "70", "1", "10", "0", "20", "50",
                 "10", "20", "20", "70", "10", "40", "20", "50",
                 "0", "TEXT", "10", "5", "20", "60", "40", "5", "1", "Bauteil A-100",
                 "0", "ENDSEC", "0", "EOF"]
    (root / "zeichnung.dxf").write_text("\n".join(dxf_lines) + "\n", encoding="ascii")
    count = cad_io.dxf_to_svg(root / "zeichnung.dxf", root / "zeichnung.svg")
    svg = (root / "zeichnung.svg").read_text(encoding="utf-8")
    assert count == 5
    assert "<line" in svg and "<circle" in svg and "<path" in svg
    assert "polygon" in svg and "A-100" in svg


def _build_iges(entities):
    def line(body, section, seq):
        return f"{body:<72}{section}{seq:7d}"

    def fmt(value):
        return f"{value:.6g}" if isinstance(value, float) else str(value)

    s_lines = [line("Test", "S", 1)]
    g_body = "1H,,1H;,4HTest;"
    g_lines = [line(g_body, "G", 1)]
    d_lines, p_lines = [], []
    for etype, params, form in entities:
        pptr = len(p_lines) + 1
        body = ",".join(fmt(v) for v in [etype] + params) + ";"
        chunks = [body[i:i + 64] for i in range(0, len(body), 64)]
        de_seq = len(d_lines) + 1
        for chunk in chunks:
            p_lines.append(f"{chunk:<64}{de_seq:8d}")
        d1 = f"{etype:8d}{pptr:8d}{0:8d}{0:8d}{0:8d}{0:8d}{0:8d}{0:8d}00000000"
        d2 = f"{etype:8d}{0:8d}{0:8d}{len(chunks):8d}{form:8d}{'':24}{0:8d}"
        d_lines.append(line(d1[:72], "D", de_seq))
        d_lines.append(line(d2[:72], "D", de_seq + 1))
    p_out = [line(pl[:72], "P", i + 1) for i, pl in enumerate(p_lines)]
    t_body = f"S{len(s_lines):7d}G{len(g_lines):7d}D{len(d_lines):7d}P{len(p_out):7d}"
    return "\n".join(s_lines + g_lines + d_lines + p_out + [line(t_body, "T", 1)]) + "\n"


def test_iges_kernel(root):
    import math
    plate = (128, [1, 1, 1, 1, 0, 0, 1, 0, 0,
                   0., 0., 1., 1., 0., 0., 1., 1., 1., 1., 1., 1.,
                   0., 0., 0., 10., 0., 0., 0., 20., 0., 10., 20., 0.,
                   0., 1., 0., 1.], 0)
    axis = (110, [30., 0., 0., 30., 0., 1.], 0)          # DE 3
    generatrix = (110, [35., 0., 0., 35., 0., 10.], 0)   # DE 5
    revolution = (120, [3, 5, 0., 2 * math.pi], 0)
    path = root / "teil.igs"
    path.write_text(_build_iges([plate, axis, generatrix, revolution]), encoding="ascii")

    mesh = iges_mesh.read_iges_mesh(path, quality="mittel")
    assert len(mesh.faces) > 20, len(mesh.faces)
    assert any(abs(v[0] - 10) < 1e-6 and abs(v[1] - 20) < 1e-6 for v in mesh.vertices), "Platte fehlt"
    cylinder = [v for v in mesh.vertices if v[0] > 20]
    radii = [math.hypot(v[0] - 30, v[1]) for v in cylinder]
    assert radii and all(abs(r - 5) < 0.05 for r in radii), (min(radii), max(radii))
    assert abs(max(v[2] for v in mesh.vertices) - 10) < 1e-6

    via_cad = cad_io.read_mesh(path, quality="grob")
    assert len(via_cad.faces) > 10


def test_pdf_rotate_reorder(root):
    src = root / "quelle.pdf"
    src.write_bytes(_MINI_PDF)
    rotated = root / "gedreht.pdf"
    count = pdf_io.rotate_pdf(src, rotated, 90, ranges_text="1")
    assert count == 1 and pdf_io.pdf_page_count(rotated) == 2
    assert b"/Rotate 90" in rotated.read_bytes()

    reordered = root / "sortiert.pdf"
    count = pdf_io.reorder_pdf(src, reordered, "2,1,2")
    assert count == 3 and pdf_io.pdf_page_count(reordered) == 3


def test_image_crop(root):
    data = bytearray()
    for y in range(20):
        for x in range(40):
            data += bytes((255, 0, 0) if x < 20 else (0, 255, 0))
    image = img_io.Image(40, 20, "RGB", data)
    right_half = img_io.crop(image, (50, 0, 50, 100))
    assert (right_half.width, right_half.height) == (20, 20)
    assert bytes(right_half.data[:3]) == bytes((0, 255, 0))


def test_split_zip_bundles(root):
    src = root / "quelle"
    src.mkdir()
    for i in range(5):
        (src / f"datei{i}.bin").write_bytes(b"x" * 400_000)   # 5 x ~0,4 MB
    files = filetools.iter_files([src], recursive=True)
    written = filetools.split_zip_bundles(files, root, "paket", 1, log=lambda m: None)
    assert len(written) == 3, [w.name for w in written]       # 2+2+1 bei 1-MB-Limit
    total = 0
    for part in written:
        with zipfile.ZipFile(part) as archive:
            total += len(archive.namelist())
    assert total == 5


def test_jpeg_roundtrip(root):
    width, height = 64, 48
    data = bytearray()
    for y in range(height):
        for x in range(width):
            data += bytes((min(255, x * 4), min(255, y * 5), 120))
    image = img_io.Image(width, height, "RGB", data, dpi=150)
    img_io.write_jpeg(image, root / "probe.jpg", quality=90)
    back = img_io.read_jpeg(root / "probe.jpg")
    assert (back.width, back.height, back.dpi) == (width, height, 150)
    diff_total = 0
    for i in range(0, len(data), 3):
        diff_total += abs(back.data[i] - data[i])
    mean_diff = diff_total / (width * height)
    assert mean_diff < 10, mean_diff

    # Encoder-Ausgabe durch write_image mit Qualitätsregler
    img_io.write_image(image, root / "klein.jpg", jpeg_quality=30)
    assert (root / "klein.jpg").stat().st_size < (root / "probe.jpg").stat().st_size


def _gif_bytes(width, height, indices, palette):
    """Minimal-GIF mit Clear-Code vor jedem Pixel (gültiges LZW ohne Tabellenaufbau)."""
    min_code = 2
    clear, end = 4, 5
    codes = []
    for idx in indices:
        codes.extend((clear, idx))
    codes.append(end)
    acc = 0
    bits = 0
    stream = bytearray()
    for code in codes:
        acc |= code << bits
        bits += 3                                   # Breite = min_code + 1
        while bits >= 8:
            stream.append(acc & 0xFF)
            acc >>= 8
            bits -= 8
    if bits:
        stream.append(acc & 0xFF)
    out = bytearray(b"GIF89a")
    out += struct.pack("<HHBBB", width, height, 0x80 | 0x01, 0, 0)   # globale Palette, 4 Farben
    for color in palette:
        out += bytes(color)
    out += b"\x2c" + struct.pack("<HHHHB", 0, 0, width, height, 0)
    out += bytes([min_code])
    out += bytes([len(stream)]) + bytes(stream) + b"\x00"
    out += b"\x3b"
    return bytes(out)


def test_gif_read(root):
    palette = [(255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0)]
    indices = [0, 1, 2, 3, 3, 2, 1, 0]
    (root / "probe.gif").write_bytes(_gif_bytes(4, 2, indices, palette))
    image = img_io.read_gif(root / "probe.gif")
    assert (image.width, image.height, image.mode) == (4, 2, "RGB")
    assert bytes(image.data[0:3]) == bytes(palette[0])
    assert bytes(image.data[3:6]) == bytes(palette[1])
    assert bytes(image.data[-3:]) == bytes(palette[0])


def test_tiff_read(root):
    width, height = 3, 2
    pixels = bytes((255, 0, 0, 0, 255, 0, 0, 0, 255,
                    10, 20, 30, 40, 50, 60, 70, 80, 90))
    entries = [
        (256, 3, 1, width), (257, 3, 1, height), (258, 3, 1, 8),
        (259, 3, 1, 1), (262, 3, 1, 2), (273, 4, 1, 0),   # Offset wird gepatcht
        (277, 3, 1, 3), (278, 3, 1, height), (279, 4, 1, len(pixels)),
    ]
    header = struct.pack("<2sHI", b"II", 42, 8)
    ifd = struct.pack("<H", len(entries))
    for tag, ftype, num, value in entries:
        ifd += struct.pack("<HHII", tag, ftype, num, value)
    ifd += struct.pack("<I", 0)
    data_offset = len(header) + len(ifd)
    blob = bytearray(header + ifd + pixels)
    for i, (tag, _t, _n, _v) in enumerate(entries):     # StripOffsets patchen
        if tag == 273:
            struct.pack_into("<I", blob, len(header) + 2 + i * 12 + 8, data_offset)
    (root / "probe.tif").write_bytes(bytes(blob))
    image = img_io.read_tiff(root / "probe.tif")
    assert (image.width, image.height, image.mode) == (width, height, "RGB")
    assert bytes(image.data) == pixels


def test_png_palette(root):
    data = bytearray()
    for i in range(40 * 30):
        data += bytes((200, 40, 40) if i % 3 else (40, 40, 200))
    image = img_io.Image(40, 30, "RGB", data)
    img_io.write_png(image, root / "gross.png")
    assert img_io.write_png_palette(image, root / "klein.png")
    assert (root / "klein.png").stat().st_size < (root / "gross.png").stat().st_size
    back = img_io.read_png(root / "klein.png")
    assert bytes(back.data) == bytes(data)


def test_images_to_pdf_and_extract(root):
    image = img_io.Image(20, 10, "RGB", bytearray([50, 100, 150] * 200), dpi=72)
    img_io.write_png(image, root / "bild.png")
    pages = pdf_io.images_to_pdf([("bild.png", image)], root / "bilder.pdf")
    assert pages == 1 and pdf_io.pdf_page_count(root / "bilder.pdf") == 1

    out = root / "raus"
    out.mkdir()
    saved = pdf_io.pdf_extract_images(root / "bilder.pdf", out, stem="bilder")
    assert len(saved) == 1 and saved[0].suffix == ".png"
    back = img_io.read_png(saved[0])
    assert (back.width, back.height) == (20, 10)
    assert bytes(back.data) == bytes(image.data)

    # JPEG wird 1:1 eingebettet und 1:1 extrahiert
    img_io.write_jpeg(image, root / "bild.jpg", quality=85)
    pdf_io.images_to_pdf([root / "bild.jpg"], root / "jpg.pdf")
    saved = pdf_io.pdf_extract_images(root / "jpg.pdf", out, stem="jpg")
    assert saved[0].suffix == ".jpg"
    assert saved[0].read_bytes() == (root / "bild.jpg").read_bytes()


def test_pdf_text(root):
    src = root / "quelle.pdf"
    src.write_bytes(_MINI_PDF)
    pdf_io.pdf_extract_text(src, root / "quelle.txt")
    text = (root / "quelle.txt").read_text(encoding="utf-8")
    assert "Angebot 4711" in text and "Seite ZWEI" in text
    assert "Seite 1" in text and "Seite 2" in text


def test_text_to_pdf(root):
    (root / "notiz.md").write_text(
        "# Bestellung 4711\n\nLieferant: Müller GmbH\n\n- Position 1: Welle\n- Position 2: Flansch\n",
        encoding="utf-8")
    pages = pdf_io.text_to_pdf(root / "notiz.md", root / "notiz.pdf", title="notiz")
    assert pages == 1 and pdf_io.pdf_page_count(root / "notiz.pdf") == 1
    pdf_io.pdf_extract_text(root / "notiz.pdf", root / "rueck.txt")
    text = (root / "rueck.txt").read_text(encoding="utf-8")
    assert "Bestellung 4711" in text and "Müller" in text and "Flansch" in text


def test_table_to_pdf_and_extract(root):
    table = tabular.Table(["SKU", "Lieferant", "Preis"],
                          [["A-100", "Müller GmbH", "12,40"],
                           ["B-200", "Schmidt AG", "8,10"],
                           ["C-300", "Demo Parts", "44,90"]])
    tabular.write_table(table, root / "preise.pdf")          # über den Tabellen-Dispatch
    assert pdf_io.pdf_page_count(root / "preise.pdf") == 1

    back = pdf_io.pdf_extract_table(root / "preise.pdf")
    flat = [cell for row in back.rows for cell in row] + back.headers
    assert any("A-100" in cell for cell in flat), flat
    assert any("Schmidt AG" in cell for cell in flat), flat
    # SKU und Preis muessen in unterschiedlichen Spalten gelandet sein
    sku_row = next(row for row in back.rows if any("B-200" in c for c in row))
    assert any("8,10" in c for c in sku_row) and len([c for c in sku_row if c]) >= 3, sku_row


def test_eml_to_pdf(root):
    eml = root / "angebot.eml"
    eml.write_bytes(
        b"From: lieferant@example.com\r\nTo: einkauf@example.com\r\n"
        b"Subject: Angebot 4711\r\nMIME-Version: 1.0\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n\r\n"
        b"Preis: 12,40 EUR\r\nLieferzeit: 3 Wochen\r\n")
    pages = filetools.eml_to_pdf(eml, root / "angebot.pdf")
    assert pages >= 1
    pdf_io.pdf_extract_text(root / "angebot.pdf", root / "angebot.txt")
    text = (root / "angebot.txt").read_text(encoding="utf-8")
    assert "Angebot 4711" in text and "12,40" in text


def test_pdf_to_html(root):
    src = root / "quelle.pdf"
    src.write_bytes(_MINI_PDF)
    pdf_io.pdf_to_html(src, root / "quelle.html")
    html = (root / "quelle.html").read_text(encoding="utf-8")
    assert "Angebot 4711" in html and "Seite 2" in html


def test_dxf_to_pdf(root):
    dxf_lines = ["0", "SECTION", "2", "ENTITIES",
                 "0", "LINE", "10", "0", "20", "0", "11", "100", "21", "0",
                 "0", "CIRCLE", "10", "50", "20", "30", "40", "10",
                 "0", "TEXT", "10", "5", "20", "60", "40", "5", "1", "Platte A-100",
                 "0", "ENDSEC", "0", "EOF"]
    (root / "zeichnung.dxf").write_text("\n".join(dxf_lines) + "\n", encoding="ascii")
    count = cad_io.dxf_to_pdf(root / "zeichnung.dxf", root / "zeichnung.pdf")
    assert count == 3
    assert pdf_io.pdf_page_count(root / "zeichnung.pdf") == 1
    raw = (root / "zeichnung.pdf").read_bytes()
    assert b" c\n" in raw and b" l S" in raw            # Bezier-Kreis + Linie
    pdf_io.pdf_extract_text(root / "zeichnung.pdf", root / "z.txt")
    assert "Platte A-100" in (root / "z.txt").read_text(encoding="utf-8")


def main():
    tests = [value for name, value in globals().items() if name.startswith("test_") and callable(value)]
    for test in tests:
        with tempfile.TemporaryDirectory() as tmp:
            test(Path(tmp))
        print(f"OK  {test.__name__}")
    print(f"\n{len(tests)} Tests bestanden.")


if __name__ == "__main__":
    main()
