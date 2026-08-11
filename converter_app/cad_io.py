"""CAD-Werkzeuge auf Basis der Standardbibliothek.

1. Mesh-Konvertierung zwischen Neutralformaten:
   STL (ASCII/binär), OBJ, PLY (ASCII/binär LE), 3MF  ->  STL/OBJ/PLY/3MF/GLB
   plus eigenständige HTML-3D-Ansicht (eigener WebGL-Viewer, offline).
2. STEP -> Mesh über den eigenen B-Rep-Tessellierungs-Kern (step_mesh.py).
3. STEP/IGES-Prüfbericht: Header, Schema, Produkte, Einheiten, Entitäten.
"""

import base64
import json
import re
import struct
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

MESH_READ_EXTS = [".stl", ".obj", ".ply", ".3mf", ".step", ".stp"]
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
        from .step_mesh import read_step_mesh
        return read_step_mesh(path, quality=quality, log=log)
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
