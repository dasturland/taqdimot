#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simsiz aloqa va elektromagnit to'lqinlar: Wi-Fi, Bluetooth va 5G
texnologiyalarining jismoniy asoslari — 10 slaydli taqdimot generatori.

- Sof Python standart kutubxonasi (tashqi bog'liqliklar shart emas).
- Har bir slayd uchun mavzuga mos fon rasmi dasturiy chiziladi (PNG).
- Agar tarmoq ochiq bo'lsa, Unsplash rasmlari avtomatik yuklab olinadi;
  aks holda dasturiy chizilgan ekvivalent fon ishlatiladi.
- Shrift: faqat "Times New Roman", eng kichik o'lcham 12 pt.
"""

import math
import os
import random
import struct
import urllib.request
import zipfile
import zlib
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------------------
# Umumiy sozlamalar
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
OUT_PPTX = os.path.join(BASE_DIR, "Simsiz_aloqa_WiFi_Bluetooth_5G.pptx")

IMG_W, IMG_H = 1280, 720
EMU_IN = 914400
SLIDE_CX, SLIDE_CY = 12192000, 6858000  # 16:9

FONT = "Times New Roman"

# Ranglar
C_WHITE = "FFFFFF"
C_BODY = "DCE3F5"
C_CYAN = "6FE3FF"
C_CYAN_DEEP = "22D3EE"
C_MUTED = "8A93B8"
C_PANEL = "0A1330"

UNSPLASH_URLS = [
    "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1508614589041-895b88991e3e?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1574607383476-f517f260d30b?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1614064641913-6b70f0556e4c?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1519999482648-25049ddd37b1?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1480694313141-fce5e697ee25?auto=format&fit=crop&w=1000&q=80",
    "https://images.unsplash.com/photo-1517245386807-bb43f82c33c4?auto=format&fit=crop&w=1000&q=80",
]

# ---------------------------------------------------------------------------
# 1-qism: PNG chizish dvijokashta (sof Python)
# ---------------------------------------------------------------------------


def hsv(h, s, v):
    i = int(h * 6)
    f = h * 6 - i
    p, q, t = v * (1 - s), v * (1 - f * s), v * (1 - (1 - f) * s)
    i %= 6
    r, g, b = [(v, t, p), (q, v, p), (p, v, t), (p, q, v), (t, p, v), (v, p, q)][i]
    return (int(r * 255), int(g * 255), int(b * 255))


class Canvas:
    def __init__(self, w=IMG_W, h=IMG_H):
        self.w, self.h = w, h
        self.buf = bytearray(w * h * 3)

    def add_px(self, x, y, col, inten):
        if 0 <= x < self.w and 0 <= y < self.h:
            i = (int(y) * self.w + int(x)) * 3
            b = self.buf
            b[i] = min(255, b[i] + col[0] * inten // 255)
            b[i + 1] = min(255, b[i + 1] + col[1] * inten // 255)
            b[i + 2] = min(255, b[i + 2] + col[2] * inten // 255)

    def gdot(self, cx, cy, rad, col, inten):
        x0, x1 = int(cx - rad), int(cx + rad) + 1
        y0, y1 = int(cy - rad), int(cy + rad) + 1
        r2 = rad * rad
        for yy in range(max(0, y0), min(self.h, y1)):
            dy = yy - cy
            for xx in range(max(0, x0), min(self.w, x1)):
                d2 = (xx - cx) ** 2 + dy * dy
                if d2 < r2:
                    f = inten * (1.0 - math.sqrt(d2) / rad)
                    if f > 0:
                        self.add_px(xx, yy, col, int(f))

    def line(self, x0, y0, x1, y1, col, inten, thick):
        dist = math.hypot(x1 - x0, y1 - y0)
        steps = max(1, int(dist * 1.6))
        for s in range(steps + 1):
            t = s / steps
            self.gdot(x0 + (x1 - x0) * t, y0 + (y1 - y0) * t, thick, col, inten)

    def arc(self, cx, cy, R, thick, col, inten, a0, a1):
        # burchaklar gradusda: 0 = o'ng, 90 = yuqori
        steps = max(12, int(abs(a1 - a0) / 180.0 * math.pi * R))
        for s in range(steps + 1):
            t = math.radians(a0 + (a1 - a0) * s / steps)
            self.gdot(cx + R * math.cos(t), cy - R * math.sin(t), thick, col, inten)

    def fill_rect(self, x0, y0, x1, y1, col):
        x0, x1 = max(0, int(x0)), min(self.w, int(x1))
        y0, y1 = max(0, int(y0)), min(self.h, int(y1))
        seg = bytes(col) * max(0, x1 - x0)
        if not seg:
            return
        for y in range(y0, y1):
            o = (y * self.w + x0) * 3
            self.buf[o:o + len(seg)] = seg

    def base_gradient(self):
        w, h = self.w, self.h
        r0 = bytes(min(255, int(8 + 5.5 * x / w * 100 / 100 * 10 * x / w)) for x in range(w))
        # soddaroq: tx bo'yicha oldindan hisoblash
        rx = [8 + int(10 * (0.55 * x / w)) for x in range(w)]
        gx = [12 + int(18 * (0.55 * x / w)) for x in range(w)]
        bx = [34 + int(46 * (0.55 * x / w)) for x in range(w)]
        for y in range(h):
            ty = 0.45 * y / h
            dr, dg, db = int(10 * ty), int(18 * ty), int(46 * ty)
            row = bytearray()
            for x in range(w):
                row.append(min(255, rx[x] + dr))
                row.append(min(255, gx[x] + dg))
                row.append(min(255, bx[x] + db))
            o = y * w * 3
            self.buf[o:o + w * 3] = row

    def darken_left(self, x_till, x_ramp, kmin):
        """Chap tomoni (matn joyi) to'qnashuvchan qorong'ulashtirish."""
        w, h = self.w, self.h
        xt, xr = x_till * w, x_ramp * w
        for x in range(w):
            if x < xt:
                k = kmin
            elif x < xr:
                k = kmin + (1.0 - kmin) * (x - xt) / (xr - xt)
            else:
                continue
            kk = int(k * 256)
            for y in range(h):
                i = (y * w + x) * 3
                b = self.buf
                b[i] = b[i] * kk >> 8
                b[i + 1] = b[i + 1] * kk >> 8
                b[i + 2] = b[i + 2] * kk >> 8

    def darken_uniform(self, k):
        kk = int(k * 256)
        b = self.buf
        for i in range(len(b)):
            b[i] = b[i] * kk >> 8


def png_bytes(cnv):
    w, h = cnv.w, cnv.h
    raw = bytearray()
    rowlen = w * 3
    for y in range(h):
        raw.append(0)
        raw += cnv.buf[y * rowlen:(y + 1) * rowlen]

    def chunk(typ, data):
        return (struct.pack(">I", len(data)) + typ + data +
                struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) +
            chunk(b"IDAT", zlib.compress(bytes(raw), 6)) + chunk(b"IEND", b""))


# ---------------------------------------------------------------------------
# 2-qism: har bir slayd uchun mavzuga mos fon rasmlari
# ---------------------------------------------------------------------------

CY = (110, 220, 255)   # asosiy cyan
CY2 = (150, 228, 255)  # ochiq cyan
PUR = (185, 125, 255)  # binafsha


def art_title(c):
    cx, cy = 0.80 * IMG_W, 0.46 * IMG_H
    c.gdot(cx, cy, 48, CY, 110)
    c.gdot(cx, cy, 16, (210, 246, 255), 255)
    for idx, R in enumerate(range(70, 421, 55)):
        c.arc(cx, cy, R, 2.2, CY, max(40, 200 - idx * 22), 0, 360)
    random.seed(11)
    for _ in range(120):
        x = random.uniform(0.52 * IMG_W, IMG_W)
        y = random.uniform(0, IMG_H)
        c.gdot(x, y, random.uniform(1.0, 2.4), (180, 230, 255), random.randint(55, 185))
    c.darken_left(0.60, 0.76, 0.30)


def art_plan(c):
    random.seed(7)
    nodes = [(random.uniform(0.56, 0.96) * IMG_W, random.uniform(0.14, 0.90) * IMG_H)
             for _ in range(22)]
    diag = math.hypot(IMG_W, IMG_H)
    for i, (x1, y1) in enumerate(nodes):
        for x2, y2 in nodes[i + 1:]:
            if math.hypot(x2 - x1, y2 - y1) < 0.17 * diag:
                c.line(x1, y1, x2, y2, (70, 170, 255), 40, 1.1)
    for j, (x, y) in enumerate(nodes):
        c.gdot(x, y, 7 if j % 6 == 0 else 4, CY2, 235)
        if j % 6 == 0:
            c.arc(x, y, 15, 1.1, CY, 120, 0, 360)
    c.darken_left(0.56, 0.72, 0.32)


def art_em_waves(c):
    by = 0.52 * IMG_H
    for gy in range(90, IMG_H, 90):
        c.line(0.05 * IMG_W, gy, IMG_W, gy, (60, 90, 160), 15, 0.8)
    c.line(0.04 * IMG_W, by, IMG_W, by, (110, 140, 200), 36, 1.0)
    for x in range(0, IMG_W, 2):
        c.gdot(x, by + 150 * math.sin(x * 0.018), 2.2, CY, 205)
        c.gdot(x, by + 150 * math.sin(x * 0.018 - 1.25), 2.2, PUR, 165)
    c.darken_left(0.54, 0.70, 0.32)


def art_spectrum(c):
    xs = int(0.50 * IMG_W)
    y0, y1 = int(0.34 * IMG_H), int(0.62 * IMG_H)
    for xi in range(xs, IMG_W):
        t = (xi - xs) / (IMG_W - xs)
        col = hsv(0.78 * t, 0.85, 1.0)
        for yy in range(y0, y1):
            d = min(yy - y0, y1 - yy) / ((y1 - y0) / 2)
            c.add_px(xi, yy, col, int(195 * min(1.0, d * 1.6)))
    for k in range(9):
        x = xs + k * (IMG_W - xs) / 8
        c.line(x, y1 + 8, x, y1 + 26, (200, 220, 255), 110, 1.3)
    c.line(xs, y0, IMG_W, y0, (255, 255, 255), 55, 1.0)
    random.seed(17)
    for _ in range(40):
        x = random.uniform(xs, IMG_W)
        h_ = random.uniform(20, 110)
        c.line(x, y0 - 14, x, y0 - 14 - h_, CY, random.randint(30, 80), 1.0)
    c.darken_left(0.50, 0.68, 0.32)


def art_wifi(c):
    cx, cy = 0.76 * IMG_W, 0.60 * IMG_H
    for R in (300, 362):
        c.arc(cx, cy, R, 1.0, (70, 140, 220), 55, 0, 360)
    for R, th in ((70, 7), (120, 8), (170, 9), (220, 10)):
        c.arc(cx, cy, R, th, CY, 232, 48, 132)
    c.gdot(cx, cy, 34, CY, 110)
    c.gdot(cx, cy, 15, (195, 242, 255), 255)
    c.darken_left(0.56, 0.72, 0.32)


def art_bluetooth(c):
    cx, cy = 0.76 * IMG_W, 0.50 * IMG_H
    c.arc(cx, cy, 172, 2.0, CY, 165, 0, 360)
    c.arc(cx, cy, 184, 1.1, CY, 85, 0, 360)
    col, inten, th = (140, 225, 255), 240, 5
    segs = [(0, 110, 0, -110), (0, 110, 75, 73), (75, 73, 0, 36),
            (0, -36, 75, -73), (75, -73, 0, -110),
            (75, 73, -75, -73), (-75, 73, 75, -73)]
    for x1, y1, x2, y2 in segs:
        c.line(cx + x1, cy - y1, cx + x2, cy - y2, col, inten, th)
    random.seed(3)
    for _ in range(28):
        x = random.uniform(0.52, 0.98) * IMG_W
        y = random.uniform(0.08, 0.92) * IMG_H
        c.gdot(x, y, 2.2, CY, random.randint(65, 165))
    c.darken_left(0.56, 0.72, 0.32)


def art_5g(c):
    R = 54
    dx, dy = math.sqrt(3) * R, 1.5 * R
    row, y = 0, R * 0.8
    while y < IMG_H + R:
        x = 0.56 * IMG_W + (dx / 2 if row % 2 else 0)
        while x < IMG_W + R:
            pts = [(x + R * math.cos(math.radians(60 * k - 30)),
                    y + R * math.sin(math.radians(60 * k - 30))) for k in range(6)]
            for k in range(6):
                c.line(*pts[k], *pts[(k + 1) % 6], (80, 180, 255), 32, 0.9)
            x += dx
        y += dy
        row += 1
    ax, ay = 0.58 * IMG_W, 0.26 * IMG_H
    for k in range(42):
        ang = math.radians(-8 + k * 1.15)
        c.line(ax, ay, ax + 430 * math.cos(ang), ay + 430 * math.sin(ang),
               (90, 200, 255), 20, 0.9)
    c.gdot(ax, ay, 10, (205, 246, 255), 255)
    for x in range(int(0.50 * IMG_W), IMG_W, 2):
        c.gdot(x, 0.74 * IMG_H + 50 * math.sin(x * 0.13), 2.0, CY, 205)
    c.darken_left(0.54, 0.70, 0.32)


def _skyline(c, seed):
    random.seed(seed)
    x, tops = 0.50 * IMG_W, []
    while x < IMG_W:
        bw = random.randint(56, 116)
        bh = random.randint(90, 250)
        c.fill_rect(x, IMG_H - bh, min(x + bw, IMG_W), IMG_H, (10, 16, 42))
        c.line(x, IMG_H - bh, min(x + bw, IMG_W), IMG_H - bh, (110, 200, 255), 65, 1.0)
        for wy in range(IMG_H - bh + 16, IMG_H - 10, 20):
            for wx in range(int(x) + 10, int(x + bw) - 8, 20):
                if random.random() < 0.42:
                    inten = random.randint(40, 105)
                    for ox, oy in ((0, 0), (1, 0), (0, 1), (1, 1)):
                        c.add_px(wx + ox, wy + oy, (150, 215, 255), inten)
        if random.random() < 0.7:
            ax, ay = x + bw / 2, IMG_H - bh
            c.gdot(ax, ay, 4, (205, 246, 255), 240)
            c.arc(ax, ay, 16, 1.0, CY, 125, 20, 160)
            c.arc(ax, ay, 30, 1.0, CY, 75, 20, 160)
            tops.append((ax, ay))
        x += bw + random.randint(6, 14)
    return tops


def art_smallcells(c):
    tops = _skyline(c, 21)
    for i in range(len(tops) - 1):
        if i % 2 == 0:
            c.line(*tops[i], *tops[i + 1], (100, 200, 255), 45, 0.9)
    c.darken_left(0.56, 0.72, 0.32)


def art_smartcity(c):
    _skyline(c, 33)
    random.seed(41)
    nodes = [(random.uniform(0.52, 0.98) * IMG_W, random.uniform(0.12, 0.45) * IMG_H)
             for _ in range(14)]
    for i, (x1, y1) in enumerate(nodes):
        for x2, y2 in nodes[i + 1:]:
            if math.hypot(x2 - x1, y2 - y1) < 205:
                c.line(x1, y1, x2, y2, (80, 185, 255), 38, 1.0)
    for nx, ny in nodes:
        c.gdot(nx, ny, 4, (165, 230, 255), 230)
    c.arc(0.75 * IMG_W, 0.27 * IMG_H, 120, 1.4, CY, 85, 0, 360)
    c.darken_left(0.54, 0.70, 0.32)


def art_thanks(c):
    cx, cy = 0.5 * IMG_W, 0.33 * IMG_H
    c.gdot(cx, cy, 330, (60, 140, 240), 58)
    c.gdot(cx, cy, 150, (90, 180, 255), 66)
    for R in (195, 265, 335):
        c.arc(cx, cy, R, 1.2, CY, 66, 0, 360)
    random.seed(5)
    for _ in range(95):
        x = random.uniform(0.06, 0.94) * IMG_W
        y = random.uniform(0.25, 0.98) * IMG_H
        inten = random.randint(50, 145)
        c.gdot(x, y, random.uniform(1.2, 2.6), (170, 225, 255), inten)
        c.gdot(x, y + random.uniform(6, 18), 1.2, (170, 225, 255), inten // 2)
    c.darken_uniform(0.82)


ART_FUNCS = [art_title, art_plan, art_em_waves, art_spectrum, art_wifi,
             art_bluetooth, art_5g, art_smallcells, art_smartcity, art_thanks]


def try_download(url, timeout=8):
    """Unsplash rasmini yuklab olishga urinadi; muvaffaqiyatsiz bo'lsa None."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        if len(data) > 20000:
            return data, "jpg"
    except Exception:
        pass
    return None, None


def build_media():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    media = []  # (filename_in_zip, bytes, local_path)
    for i, func in enumerate(ART_FUNCS, start=1):
        data, ext = try_download(UNSPLASH_URLS[i - 1])
        if data is None:
            cnv = Canvas()
            cnv.base_gradient()
            func(cnv)
            data, ext = png_bytes(cnv), "png"
            src = "dasturiy chizilgan fon"
        else:
            src = "Unsplash (yuklab olindi)"
        local = os.path.join(ASSETS_DIR, f"slide{i}.{ext}")
        with open(local, "wb") as f:
            f.write(data)
        media.append((f"image{i}.{ext}", data, local))
        print(f"  Slayd {i:2d}: {src} -> assets/slide{i}.{ext} ({len(data)//1024} KB)")
    return media


# ---------------------------------------------------------------------------
# 3-qism: OOXML (PPTX) yaratish
# ---------------------------------------------------------------------------

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def IN(v):
    return int(round(v * EMU_IN))


class SlideBuilder:
    def __init__(self):
        self.shapes = []
        self._id = 1

    def nid(self):
        self._id += 1
        return self._id

    # --- rasm ---
    def pic(self, rid, x, y, w, h, name="Background"):
        self.shapes.append(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{self.nid()}" name="{name}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')

    # --- to'g'ri to'rtburchak ---
    def rect(self, x, y, w, h, fill=None, alpha=None, line=None, line_w=12700,
             name="Shape", round_=False):
        if fill is None:
            fill_xml = "<a:noFill/>"
        else:
            a = f'<a:alpha val="{alpha}"/>' if alpha else ""
            fill_xml = f'<a:solidFill><a:srgbClr val="{fill}">{a}</a:srgbClr></a:solidFill>'
        ln_xml = ("<a:ln><a:noFill/></a:ln>" if line is None else
                  f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>')
        geom = "roundRect" if round_ else "rect"
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{name}"/>'
            f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="uz-UZ"/></a:p></p:txBody></p:sp>')

    # --- matn qutisi ---
    def text(self, x, y, w, h, paras, anchor="t", name="TextBox"):
        ps = "".join(self._para(p) for p in paras)
        self.shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self.nid()}" name="{name}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
            f'<a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" anchor="{anchor}" lIns="0" tIns="0" '
            f'rIns="0" bIns="0"/><a:lstStyle/>{ps}</p:txBody></p:sp>')

    @staticmethod
    def _run(t, sz, b=False, i=False, color=C_WHITE, spc=None):
        attrs = f'lang="uz-UZ" sz="{sz}"'
        if b:
            attrs += ' b="1"'
        if i:
            attrs += ' i="1"'
        if spc:
            attrs += f' spc="{spc}"'
        return (f'<a:r><a:rPr {attrs} dirty="0">'
                f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/>'
                f'<a:cs typeface="{FONT}"/></a:rPr>'
                f'<a:t xml:space="preserve">{esc(t)}</a:t></a:r>')

    @classmethod
    def _para(cls, p):
        runs = "".join(cls._run(**r) for r in p["runs"])
        inner = ""
        if p.get("ls"):
            inner += f'<a:lnSpc><a:spcPct val="{p["ls"]}"/></a:lnSpc>'
        if p.get("sb"):
            inner += f'<a:spcBef><a:spcPts val="{p["sb"]}"/></a:spcBef>'
        if p.get("bullet"):
            inner += (f'<a:buClr><a:srgbClr val="{C_CYAN}"/></a:buClr>'
                      f'<a:buFont typeface="{FONT}"/><a:buChar char="•"/>')
        else:
            inner += "<a:buNone/>"
        attrs = ""
        if p.get("bullet"):
            attrs += ' marL="274320" indent="-274320"'
        if p.get("align"):
            attrs += f' algn="{p["align"]}"'
        return f"<a:p><a:pPr{attrs}>{inner}</a:pPr>{runs}</a:p>"

    def xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
            "<p:cSld><p:spTree>"
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(self.shapes) +
            "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>")


# --- matn qismlari uchun qisqartmalar ---------------------------------------

def R(t, sz=1600, b=False, i=False, color=C_WHITE, spc=None):
    return {"t": t, "sz": sz, "b": b, "i": i, "color": color, "spc": spc}


def footer(sb, num):
    sb.text(IN(0.55), IN(7.04), IN(6.5), IN(0.34),
            [{"runs": [R("Jizzax davlat pedagogika universiteti • Axborot tizimlari",
                         sz=1200, color=C_MUTED)]}], name="Footer")
    sb.text(IN(12.10), IN(7.04), IN(1.0), IN(0.34),
            [{"runs": [R(f"{num} / 10", sz=1200, color=C_MUTED)], "align": "r"}],
            name="PageNum")


def header(sb, chip, title):
    sb.rect(IN(0.55), IN(0.47), IN(0.16), IN(0.16), fill=C_CYAN_DEEP, name="ChipSq")
    sb.text(IN(0.82), IN(0.40), IN(5.0), IN(0.32),
            [{"runs": [R(chip, sz=1300, b=True, color=C_CYAN, spc=300)]}], name="Chip")
    sb.text(IN(0.53), IN(0.74), IN(11.2), IN(0.70),
            [{"runs": [R(title, sz=2800, b=True)]}], name="Title")
    sb.rect(IN(0.56), IN(1.46), IN(1.60), IN(0.035), fill=C_CYAN_DEEP, name="TitleBar")


def body_panel(sb):
    sb.rect(IN(0.50), IN(1.66), IN(7.45), IN(5.05), fill=C_PANEL, alpha=52000,
            name="Panel", round_=True)
    sb.rect(IN(0.50), IN(1.86), IN(0.035), IN(4.65), fill=C_CYAN_DEEP, name="PanelBar")


def make_bg(sb, num):
    sb.pic("rId2", 0, 0, SLIDE_CX, SLIDE_CY, name=f"BgImage{num}")


# --- slaydlar ----------------------------------------------------------------

def slide1():
    sb = SlideBuilder()
    make_bg(sb, 1)
    sb.text(IN(0.65), IN(0.52), IN(9.5), IN(0.4),
            [{"runs": [R("JIZZAX DAVLAT PEDAGOGIKA UNIVERSITETI",
                         sz=1400, b=True, color=C_CYAN, spc=300)]}], name="UniChip")
    sb.rect(IN(0.65), IN(1.78), IN(0.07), IN(2.30), fill=C_CYAN_DEEP, name="AccentBar")
    sb.text(IN(0.95), IN(1.80), IN(7.4), IN(1.85),
            [{"runs": [R("Simsiz aloqa va", sz=3600, b=True)], "ls": 104000},
             {"runs": [R("elektromagnit to'lqinlar", sz=3600, b=True)], "ls": 104000}],
            name="MainTitle")
    sb.text(IN(0.95), IN(3.62), IN(7.2), IN(1.0),
            [{"runs": [R("Wi-Fi, Bluetooth va 5G texnologiyalarining jismoniy asoslari",
                         sz=2000, color=C_CYAN)], "ls": 110000}], name="SubTitle")
    sb.rect(IN(0.95), IN(4.62), IN(2.2), IN(0.03), fill=C_CYAN_DEEP, name="Divider")
    info = [("Universitet: ",
             "Jizzax davlat pedagogika universiteti tabiiy fanlar oliy maktabi"),
            ("Fan: ", "Axborot tizimlari"),
            ("Guruh: ", "242_26"),
            ("Talaba: ", "Toshboʻriyeva Marjona")]
    paras = [{"runs": [R(lbl, sz=1600, b=True, color=C_CYAN),
                       R(val, sz=1600, color=C_BODY)],
              "sb": 0 if k == 0 else 700, "ls": 108000}
             for k, (lbl, val) in enumerate(info)]
    sb.text(IN(0.95), IN(4.82), IN(7.4), IN(2.2), paras, name="Info")
    return sb.xml()


def slide2():
    sb = SlideBuilder()
    make_bg(sb, 2)
    header(sb, "MUNDARIJA", "Taqdimot rejasi")
    body_panel(sb)
    items = [("1-reja. ", "Elektromagnit to'lqinlar va chastota spektri fizikasi."),
             ("2-reja. ", "Wi-Fi va Bluetooth texnologiyalarining ishlash mexanizmlari."),
             ("3-reja. ", "5G inqilobi: Millimetrli to'lqinlar va tarmoq tezligi."),
             ("Xulosa. ", "Simsiz axborot tizimlarining jamiyatdagi o'rni.")]
    paras = [{"runs": [R(a, sz=1800, b=True, color=C_CYAN),
                       R(b_, sz=1800, color=C_WHITE)],
              "sb": 0 if k == 0 else 1500, "ls": 112000}
             for k, (a, b_) in enumerate(items)]
    sb.text(IN(0.85), IN(2.00), IN(6.9), IN(4.5), paras, name="Agenda")
    footer(sb, 2)
    return sb.xml()


def content_slide(num, chip, title, paras):
    sb = SlideBuilder()
    make_bg(sb, num)
    header(sb, chip, title)
    body_panel(sb)
    body = []
    for kind, label, txt in paras:
        runs = []
        if label:
            runs.append(R(label, sz=1600, b=True, color=C_CYAN))
        runs.append(R(txt, sz=1600, color=C_BODY))
        body.append({"runs": runs, "bullet": kind == "b",
                     "sb": 0 if not body else 1000, "ls": 115000})
    sb.text(IN(0.85), IN(1.95), IN(6.9), IN(4.6), body, name="Body")
    footer(sb, num)
    return sb.xml()


def slide3():
    return content_slide(
        3, "1-REJA", "1. Elektromagnit to'lqinlar fizikasi",
        [("p", None, "Simsiz aloqa ko'zga ko'rinmas elektromagnit to'lqinlar yordamida amalga oshiriladi."),
         ("b", "Asosiy tushuncha: ",
          "Axborot tizimlaridagi raqamli ma'lumotlar (0 va 1) elektr va magnit maydonlarining "
          "tebranishlari orqali fazoda yorug'lik tezligida (300,000 km/s) tarqaladi."),
         ("b", "To'lqin uzunligi va chastota: ",
          "Chastota (Gerts) bir soniyadagi to'lqinlar tebranishi sonini bildiradi. Chastota "
          "qanchalik yuqori bo'lsa, to'lqin shunchalik ko'p ma'lumot tashiydi, lekin qisqa masofaga boradi.")])


def slide4():
    return content_slide(
        4, "1-REJA", "Radioto'lqinlar spektri",
        [("p", None, "Turli texnologiyalar elektromagnit spektrning turli qismlaridan foydalanadi:"),
         ("b", "Radio va TV: ", "Past chastotali, juda uzoq masofalarga tarqaluvchi to'lqinlar."),
         ("b", "Mikroto'lqinlar: ",
          "Wi-Fi, Bluetooth va mobil aloqa (4G, 5G) aynan shu diapazonda (Gigagerts - GHz) ishlaydi."),
         ("p", None,
          "Axborot tizimlarida signallar bir-biriga xalaqit bermasligi (interferensiya bo'lmasligi) "
          "uchun har bir qurilma o'ziga ajratilgan qat'iy chastota yo'lagida harakatlanadi.")])


def slide5():
    return content_slide(
        5, "2-REJA", "2. Wi-Fi texnologiyasi qanday ishlaydi?",
        [("p", None, "Wi-Fi (Wireless Fidelity) lokal axborot tarmoqlarining asosidir."),
         ("b", "Ishlash chastotasi: ", "Wi-Fi asosan ikkita chastotada — 2.4 GHz va 5 GHz da ishlaydi."),
         ("b", "Fizik farqi: ",
          "2.4 GHz chastota devorlardan oson o'tadi va uzoqroqqa boradi, lekin tezligi past. "
          "5 GHz chastota esa juda yuqori tezlikda axborot uzatadi, biroq to'siqlarga tez urilib "
          "so'nadi (qisqa masofali). Router ma'lumotlarni paketlarga bo'lib, shu to'lqinlar orqali tarqatadi.")])


def slide6():
    return content_slide(
        6, "2-REJA", "Bluetooth va qisqa masofali aloqa",
        [("p", None,
          "Bluetooth texnologiyasi ham Wi-Fi kabi 2.4 GHz chastotada ishlaydi, biroq u uzoq masofaga "
          "emas, balki shaxsiy qurilmalarni (smartfon, quloqchin, soat) bog'lashga mo'ljallangan."),
         ("b", "Frequency Hopping: ",
          "Signallar aralashib ketmasligi uchun Bluetooth soniyasiga 1600 marta o'z chastotasini "
          "almashtirib turadi."),
         ("b", "Energiya tejamkorligi: ",
          "Ushbu texnologiyaning radioto'lqinlari juda past quvvatda tarqaladi, shu sababli "
          "batareya quvvatini kam sarflaydi.")])


def slide7():
    return content_slide(
        7, "3-REJA", "3. 5G inqilobi va texnologiyasi",
        [("p", None,
          "Beshinchi avlod (5G) mobil tarmoqlari telekommunikatsiya va axborot tizimlarida "
          "inqilobiy tezlikni (10 Gbit/s gacha) taqdim etadi."),
         ("b", "Millimetrli to'lqinlar (mmWave): ",
          "5G juda yuqori chastotalardan (24 GHz dan 100 GHz gacha) foydalanadi. Bu to'lqinlar "
          "o'ta kalta bo'lib, ma'lumotlarni misli ko'rilmagan tezlikda tashiydi."),
         ("b", "Kam kechikish (Low Latency): ",
          "Signalning borib qaytish vaqti 1 millisoniyagacha qisqartirilgan, bu kelajakda avtopilot "
          "mashinalar va masofaviy jarrohlik amaliyotlari uchun muhimdir.")])


def slide8():
    return content_slide(
        8, "3-REJA", "5G arxitekturasi va to'siqlar fizikasi",
        [("p", None,
          "Millimetrli to'lqinlarning asosiy kamchiligi ularning fizik tabiatidir — ular yomg'ir, "
          "daraxt barglari yoki hatto odam tanasidan ham o'ta olmaydi, tez so'nadi."),
         ("b", "Kichik stansiyalar (Small Cells): ",
          "Shu sababli 5G tarmog'ini yaratishda bitta katta minora emas, balki har bir ko'cha va "
          "binoga o'nlab mitti antennalar o'rnatiladi."),
         ("b", "Beamforming (Nurni yo'naltirish): ",
          "5G signallarni barcha tomonga sochamaydi, balki to'g'ridan-to'g'ri foydalanuvchi "
          "smartfoniga lazer kabi aniq yo'naltiradi.")])


def slide9():
    sb = SlideBuilder()
    make_bg(sb, 9)
    header(sb, "XULOSA", "Xulosa")
    sb.text(IN(0.55), IN(1.80), IN(7.4), IN(1.1),
            [{"runs": [R("Xulosa qilib aytganda, simsiz aloqa texnologiyalari to'liq "
                         "elektromagnit fizika qonuniyatlariga tayanadi.",
                         sz=1700, color=C_WHITE)], "ls": 115000}], name="Lead")
    sb.rect(IN(0.50), IN(3.00), IN(7.45), IN(3.55), fill=C_PANEL, alpha=58000,
            line=C_CYAN_DEEP, line_w=15875, name="Box", round_=True)
    sb.text(IN(0.85), IN(3.30), IN(6.85), IN(3.0),
            [{"runs": [R("Wi-Fi va Bluetooth kundalik lokal ehtiyojlarimizni qondirsa, ",
                         sz=1600, color=C_BODY),
                       R("5G to'lqinlari", sz=1600, b=True, color=C_CYAN),
                       R(" kelajakdagi ", sz=1600, color=C_BODY),
                       R("“Aqlli shaharlar” (Smart Cities)", sz=1600, b=True, color=C_CYAN),
                       R(", ", sz=1600, color=C_BODY),
                       R("buyumlar interneti (IoT)", sz=1600, b=True, color=C_CYAN),
                       R(" va global axborot tizimlari (masalan, JDPU ta'lim platformalari) "
                         "integratsiyasi uchun mustahkam infratuzilma poydevori bo'lib xizmat qiladi.",
                         sz=1600, color=C_BODY)], "ls": 118000}], name="BoxText")
    footer(sb, 9)
    return sb.xml()


def slide10():
    sb = SlideBuilder()
    make_bg(sb, 10)
    sb.text(IN(0), IN(2.30), IN(13.333), IN(1.1),
            [{"runs": [R("E'tiboringiz uchun rahmat!", sz=4000, b=True)],
              "align": "ctr"}], name="Thanks")
    sb.rect(IN(5.867), IN(3.52), IN(1.6), IN(0.04), fill=C_CYAN_DEEP, name="Divider")
    details = [("Mavzu: ", "Simsiz aloqa: Wi-Fi, Bluetooth va 5G fizikasi"),
               ("Taqdimotchi: ", "Toshboʻriyeva Marjona (242_26 - guruh)"),
               (None, "Jizzax davlat pedagogika universiteti - Axborot tizimlari fani.")]
    paras = []
    for lbl, val in details:
        runs = []
        if lbl:
            runs.append(R(lbl, sz=1600, b=True, color=C_CYAN))
        runs.append(R(val, sz=1600, color=C_BODY))
        paras.append({"runs": runs, "align": "ctr", "sb": 0 if not paras else 900,
                      "ls": 110000})
    sb.text(IN(1.667), IN(3.95), IN(10.0), IN(2.1), paras, name="Details")
    return sb.xml()


# --- statik XML qismlar -------------------------------------------------------

def content_types(nslides, exts):
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
         '<Default Extension="xml" ContentType="application/xml"/>']
    for e in sorted(set(exts)):
        ct = "image/png" if e == "png" else "image/jpeg"
        o.append(f'<Default Extension="{e}" ContentType="{ct}"/>')
    base = "application/vnd.openxmlformats-officedocument.presentationml"
    o.append(f'<Override PartName="/ppt/presentation.xml" ContentType="{base}.presentation.main+xml"/>')
    o.append(f'<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="{base}.slideMaster+xml"/>')
    o.append(f'<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="{base}.slideLayout+xml"/>')
    for i in range(1, nslides + 1):
        o.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="{base}.slide+xml"/>')
    o.append(f'<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
    o.append(f'<Override PartName="/ppt/presProps.xml" ContentType="{base}.presProps+xml"/>')
    o.append(f'<Override PartName="/ppt/viewProps.xml" ContentType="{base}.viewProps+xml"/>')
    o.append(f'<Override PartName="/ppt/tableStyles.xml" ContentType="{base}.tableStyles+xml"/>')
    o.append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    o.append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    o.append("</Types>")
    return "".join(o)


def rels_xml(rels):
    items = "".join(
        f'<Relationship Id="{rid}" Type="{typ}" Target="{tgt}"/>' for rid, typ, tgt in rels)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{items}</Relationships>")


RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ROOT_RELS = rels_xml([
    ("rId1", f"{RT}/officeDocument", "ppt/presentation.xml"),
    ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
     "docProps/core.xml"),
    ("rId3", f"{RT}/extended-properties", "docProps/app.xml"),
])

PRES_RELS = rels_xml(
    [("rId1", f"{RT}/slideMaster", "slideMasters/slideMaster1.xml")] +
    [(f"rId{i + 1}", f"{RT}/slide", f"slides/slide{i}.xml") for i in range(1, 11)] +
    [("rId12", f"{RT}/presProps", "presProps.xml"),
     ("rId13", f"{RT}/viewProps", "viewProps.xml"),
     ("rId14", f"{RT}/tableStyles", "tableStyles.xml")])

MASTER_RELS = rels_xml([
    ("rId1", f"{RT}/slideLayout", "../slideLayouts/slideLayout1.xml"),
    ("rId2", f"{RT}/theme", "../theme/theme1.xml"),
])

LAYOUT_RELS = rels_xml([("rId1", f"{RT}/slideMaster", "../slideMasters/slideMaster1.xml")])

PRESENTATION_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    "<p:sldIdLst>" + "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, 11)) +
    "</p:sldIdLst>"
    f'<p:sldSz cx="{SLIDE_CX}" cy="{SLIDE_CY}" type="screen16x9"/>'
    '<p:notesSz cx="6858000" cy="9144000"/>'
    '<p:defaultTextStyle><a:defPPr><a:defRPr lang="uz-UZ">'
    f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/>'
    "</a:defRPr></a:defPPr></p:defaultTextStyle></p:presentation>")

MASTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p:sldMaster xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="070C1D"/></a:solidFill>'
    "<a:effectLst/></p:bgPr></p:bg><p:spTree>"
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    "</p:spTree></p:cSld>"
    '<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" '
    'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
    'hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    "<p:txStyles>"
    '<p:titleStyle><a:lvl1pPr><a:defRPr sz="3200" b="1">'
    '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>'
    f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1600">'
    '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>'
    f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1600">'
    f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
    "</p:txStyles></p:sldMaster>")

LAYOUT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p:sldLayout xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}" type="blank" preserve="1">'
    '<p:cSld name="Blank"><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>")

THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<a:theme xmlns:a="{NS_A}" name="JDPU Cyber"><a:themeElements>'
    '<a:clrScheme name="Cyber">'
    '<a:dk1><a:srgbClr val="0B1026"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="141B3D"/></a:dk2><a:lt2><a:srgbClr val="DCE3F5"/></a:lt2>'
    '<a:accent1><a:srgbClr val="22D3EE"/></a:accent1><a:accent2><a:srgbClr val="7C6CFF"/></a:accent2>'
    '<a:accent3><a:srgbClr val="FFD166"/></a:accent3><a:accent4><a:srgbClr val="4ADE80"/></a:accent4>'
    '<a:accent5><a:srgbClr val="F472B6"/></a:accent5><a:accent6><a:srgbClr val="60A5FA"/></a:accent6>'
    '<a:hlink><a:srgbClr val="6FE3FF"/></a:hlink><a:folHlink><a:srgbClr val="A78BFA"/></a:folHlink>'
    "</a:clrScheme>"
    f'<a:fontScheme name="TNR"><a:majorFont><a:latin typeface="{FONT}"/><a:ea typeface=""/>'
    f'<a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="{FONT}"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/>'
    '<a:tint val="67000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/>'
    '<a:tint val="73000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/>'
    '<a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/>'
    '<a:tint val="94000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/>'
    '<a:shade val="100000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/>'
    '<a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    "</a:fillStyleLst>"
    '<a:lnStyleLst>'
    '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>'
    '</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>'
    '</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>'
    '</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    "</a:lnStyleLst>"
    "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"
    "<a:effectStyle><a:effectLst/></a:effectStyle>"
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" '
    'algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>'
    "</a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>"
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/>'
    "</a:schemeClr></a:solidFill>"
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/>'
    '<a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/>'
    '<a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/>'
    "</a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
    "</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>")

PRES_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<p:presentationPr xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"/>')

VIEW_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<p:viewPr xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}">'
              '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/>'
              "</p:normalViewPr><p:slideViewPr/><p:notesTextViewPr/>"
              '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>')

TABLE_STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<a:tblStyleLst xmlns:a="{NS_A}" '
                'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>')

CORE_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>Simsiz aloqa va elektromagnit to'lqinlar: Wi-Fi, Bluetooth va 5G "
            "texnologiyalarining jismoniy asoslari</dc:title>"
            "<dc:creator>Toshboʻriyeva Marjona</dc:creator>"
            "<cp:lastModifiedBy>Kiro</cp:lastModifiedBy>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-10-18T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-18T00:00:00Z</dcterms:modified>'
            "</cp:coreProperties>")

APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           "<TotalTime>0</TotalTime><Words>0</Words>"
           "<Application>Kiro Presentation Generator</Application>"
           "<PresentationFormat>16:9</PresentationFormat>"
           "<Paragraphs>0</Paragraphs><Slides>10</Slides><Notes>0</Notes>"
           "<HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop>"
           "<Company>JDPU</Company><LinksUpToDate>false</LinksUpToDate>"
           "<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>"
           "<AppVersion>1.0</AppVersion></Properties>")


def build_pptx(media, slides_xml):
    exts = [m[0].rsplit(".", 1)[1] for m in media]
    parts = {
        "[Content_Types].xml": content_types(10, exts),
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": CORE_XML,
        "docProps/app.xml": APP_XML,
        "ppt/presentation.xml": PRESENTATION_XML,
        "ppt/_rels/presentation.xml.rels": PRES_RELS,
        "ppt/slideMasters/slideMaster1.xml": MASTER_XML,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": LAYOUT_XML,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": LAYOUT_RELS,
        "ppt/theme/theme1.xml": THEME_XML,
        "ppt/presProps.xml": PRES_PROPS,
        "ppt/viewProps.xml": VIEW_PROPS,
        "ppt/tableStyles.xml": TABLE_STYLES,
    }
    for i, sx in enumerate(slides_xml, start=1):
        parts[f"ppt/slides/slide{i}.xml"] = sx
        fname = media[i - 1][0]
        parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = rels_xml([
            ("rId1", f"{RT}/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", f"{RT}/image", f"../media/{fname}"),
        ])
    with zipfile.ZipFile(OUT_PPTX, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml)
        for fname, data, _ in media:
            z.writestr(f"ppt/media/{fname}", data)


# ---------------------------------------------------------------------------
# 4-qism: tekshirish (validatsiya)
# ---------------------------------------------------------------------------

def validate():
    import re
    ok = True
    with zipfile.ZipFile(OUT_PPTX) as z:
        names = set(z.namelist())
        # 1) barcha XML hujjatlar well-formed
        for n in sorted(names):
            if n.endswith((".xml", ".rels")):
                try:
                    ET.fromstring(z.read(n))
                except ET.ParseError as e:
                    print(f"  XATO: {n}: {e}")
                    ok = False
        # 2) barcha relationship targetlari mavjud
        for n in sorted(names):
            if n.endswith(".rels"):
                base = os.path.dirname(os.path.dirname(n))
                root = ET.fromstring(z.read(n))
                for rel in root:
                    tgt = rel.get("Target")
                    full = os.path.normpath(os.path.join(base, tgt)).replace("\\", "/")
                    if full not in names:
                        print(f"  XATO: {n} -> {tgt} topilmadi")
                        ok = False
        # 3) slaydlar soni
        slides = sorted(n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        print(f"  Slaydlar soni: {len(slides)} (kutilgan: 10)")
        ok &= len(slides) == 10
        # 4) shrift va o'lcham tekshiruvi
        min_sz, tnr_all, pic_count = 10 ** 9, True, 0
        for s in slides:
            xml = z.read(s).decode("utf-8")
            for m in re.finditer(r'sz="(\d+)"', xml):
                min_sz = min(min_sz, int(m.group(1)))
            if f'typeface="{FONT}"' not in xml:
                tnr_all = False
            # boshqa shrift qoldiqmi?
            for m in re.finditer(r'typeface="([^"]+)"', xml):
                if m.group(1) not in (FONT,):
                    tnr_all = False
                    print(f"  XATO: {s} da begona shrift: {m.group(1)}")
            pic_count += xml.count("<p:pic>")
        print(f"  Eng kichik shrift: {min_sz / 100:.0f} pt (chegara: 12 pt)")
        print(f"  Faqat Times New Roman: {'HA' if tnr_all else 'YO‘Q'}")
        print(f"  Slaydlardagi rasmlar soni: {pic_count} (kutilgan: 10)")
        ok &= min_sz >= 1200 and tnr_all and pic_count == 10
        # 5) media fayllar
        media_n = [n for n in names if n.startswith("ppt/media/")]
        print(f"  Media fayllar: {len(media_n)}")
        ok &= len(media_n) == 10
    size = os.path.getsize(OUT_PPTX)
    print(f"  Fayl hajmi: {size / 1024:.0f} KB")
    print("  UMUMIY HOLAT:", "MAVFFAQIYATLI ✅" if ok else "XATO ❌")
    return ok


def main():
    print("1) Rasmlar tayyorlanmoqda...")
    media = build_media()
    print("2) Slaydlar XML yaratilmoqda...")
    slides_xml = [slide1(), slide2(), slide3(), slide4(), slide5(),
                  slide6(), slide7(), slide8(), slide9(), slide10()]
    print("3) PPTX jamlanmoqda...")
    build_pptx(media, slides_xml)
    print(f"   -> {OUT_PPTX}")
    print("4) Validatsiya:")
    if not validate():
        raise SystemExit("Validatsiya xatosi!")
    print("TAYYOR.")


if __name__ == "__main__":
    main()
