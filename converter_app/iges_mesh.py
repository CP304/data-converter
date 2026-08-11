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

import math
from collections import Counter
from pathlib import Path

from .cad_io import Mesh, _VertexPool
from .step_mesh import (
    QUALITY_SEGMENTS, TWO_PI, BSplineCurve, BSplineSurface, Surface,
    _add, _cross, _dot, _generic_uv_of, _mul, _norm, _sub, _unwrap_loop,
    face_triangles_from_uv, grid_triangles,
)


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
