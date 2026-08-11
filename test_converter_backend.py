"""Backend-Tests ohne GUI: python test_converter_backend.py"""

import tempfile
import zipfile
from pathlib import Path

from converter_app import cad_io, filetools, step_mesh, tabular, xlsx_io


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


def main():
    tests = [value for name, value in globals().items() if name.startswith("test_") and callable(value)]
    for test in tests:
        with tempfile.TemporaryDirectory() as tmp:
            test(Path(tmp))
        print(f"OK  {test.__name__}")
    print(f"\n{len(tests)} Tests bestanden.")


if __name__ == "__main__":
    main()
