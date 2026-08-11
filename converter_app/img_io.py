"""Bild-Werkzeuge auf Basis der Standardbibliothek: PNG- und BMP-Codec,
Skalieren (bilinear), Wasserzeichen (eigener 5x7-Pixelfont), DPI setzen.

JPEG/TIFF/HEIC brauchen echte Codecs und bleiben bewusst außen vor.
"""

import struct
import zlib
from dataclasses import dataclass, field
from pathlib import Path

READ_EXTS = [".png", ".bmp"]
WRITE_EXTS = [".png", ".bmp"]


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
    raise ValueError(f"Kein unterstütztes Bildformat: {ext} (nur PNG/BMP).")


def write_image(image, path):
    ext = Path(path).suffix.lower()
    if ext == ".png":
        write_png(image, path)
    elif ext == ".bmp":
        write_bmp(image, path)
    else:
        raise ValueError(f"Kein unterstütztes Bild-Zielformat: {ext} (nur PNG/BMP).")


# ---------------------------------------------------------------------------
# Operationen
# ---------------------------------------------------------------------------

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
