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

import math
import re
from collections import Counter
from pathlib import Path

from .cad_io import Mesh, _VertexPool

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
