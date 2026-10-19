#!/usr/bin/env python3
"""Presentation figure for the SM25-07 U-Pb zircon dataset.

Left: Wetherill concordia (207Pb/235U versus 206Pb/238U), all analyses as
2-sigma error ellipses. Right: accepted 206Pb/238U ages with horizontal
2-sigma error bars and their inverse-variance weighted mean.
"""

from __future__ import annotations

import csv
import math
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "upb" / "SM25-07_upb_audit.csv"
OUT_PNG = ROOT / "assets" / "SM25-07_concordia_WM_horizontal_30cm_600dpi.png"
OUT_PDF = ROOT / "assets" / "SM25-07_concordia_WM_horizontal_30cm_600dpi.pdf"

# Logical drawing canvas. Every primitive and glyph is rendered natively at
# three times this size: 7086 x 3366 px, or 30.0 x 14.25 cm at 600 dpi.
WIDTH, HEIGHT = 2362, 1122
RENDER_SCALE = 3
LAMBDA_238 = 1.55125e-10
LAMBDA_235 = 9.8485e-10

FONT = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"

INK = "#27312c"
MUTED = "#68736e"
GRID = "#dfe5e1"
GREEN = "#126e64"
GREEN_EDGE = "#0c5d55"
GRAY = "#c1c6c3"
GRAY_EDGE = "#7d8581"
MEAN = "#a33e3a"
MEAN_BAND = "#d9ebe7"
DISCORDANCE_LIMIT = 5.0


def load_rows() -> list[dict[str, object]]:
    numeric = {
        "ratio_75", "err_ratio_75", "ratio_68", "err_ratio_68", "rho_3",
        "age_68", "err_age_68", "d1",
    }
    with SOURCE.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    for row in rows:
        for field in numeric:
            row[field] = float(row[field])
        # Recalculate the filter for this presentation figure rather than use
        # the portal audit's earlier 10% acceptance flag.
        row["accepted"] = abs(row["d1"]) <= DISCORDANCE_LIMIT
    return rows


def weighted_mean(rows: list[dict[str, object]]) -> tuple[float, float, float]:
    ages = np.array([r["age_68"] for r in rows], dtype=float)
    errors = np.array([r["err_age_68"] for r in rows], dtype=float)
    weights = 1.0 / errors**2
    mean = float(np.sum(weights * ages) / np.sum(weights))
    uncertainty_2s = float(2.0 / math.sqrt(float(np.sum(weights))))
    mswd = float(np.sum(weights * (ages - mean) ** 2) / (len(rows) - 1))
    return mean, uncertainty_2s, mswd


def concordia(age_ma: float) -> tuple[float, float]:
    years = age_ma * 1e6
    return math.expm1(LAMBDA_235 * years), math.expm1(LAMBDA_238 * years)


class ScaledDraw:
    """ImageDraw facade that scales logical coordinates and stroke widths."""

    def __init__(self, image: Image.Image, scale: int):
        self.raw = ImageDraw.Draw(image)
        self.scale = scale

    def _point(self, point):
        return tuple(float(value) * self.scale for value in point)

    def _points(self, values):
        if len(values) and isinstance(values[0], (tuple, list, np.ndarray)):
            return [self._point(point) for point in values]
        return tuple(float(value) * self.scale for value in values)

    def line(self, xy, fill=None, width=1, joint=None):
        self.raw.line(self._points(xy), fill=fill, width=max(1, round(width * self.scale)), joint=joint)

    def rectangle(self, xy, fill=None, outline=None, width=1):
        self.raw.rectangle(self._points(xy), fill=fill, outline=outline, width=max(1, round(width * self.scale)))

    def rounded_rectangle(self, xy, radius=0, fill=None, outline=None, width=1):
        self.raw.rounded_rectangle(
            self._points(xy), radius=radius * self.scale, fill=fill, outline=outline,
            width=max(1, round(width * self.scale)),
        )

    def ellipse(self, xy, fill=None, outline=None, width=1):
        self.raw.ellipse(self._points(xy), fill=fill, outline=outline, width=max(1, round(width * self.scale)))

    def polygon(self, xy, fill=None, outline=None):
        self.raw.polygon(self._points(xy), fill=fill, outline=outline)

    def text(self, xy, text, font=None, fill=None, anchor=None, stroke_width=0, stroke_fill=None):
        self.raw.text(
            self._point(xy), text, font=font, fill=fill, anchor=anchor,
            stroke_width=round(stroke_width * self.scale), stroke_fill=stroke_fill,
        )


def main() -> None:
    rows = load_rows()
    accepted = sorted((r for r in rows if r["accepted"]), key=lambda r: r["age_68"])
    rejected = [r for r in rows if not r["accepted"]]
    mean, mean_2s, mswd = weighted_mean(accepted)

    image = Image.new("RGB", (WIDTH * RENDER_SCALE, HEIGHT * RENDER_SCALE), "white")
    draw = ScaledDraw(image, RENDER_SCALE)
    regular = {s: ImageFont.truetype(FONT, s * RENDER_SCALE) for s in (17, 18, 19, 20, 22, 24, 25, 26, 28, 30, 32, 34)}
    bold = {s: ImageFont.truetype(FONT_BOLD, s * RENDER_SCALE) for s in (22, 24, 26, 28, 30, 32, 34)}

    def label(x, y, value, size=24, anchor="mm", color=INK, strong=False):
        draw.text((x, y), str(value), font=(bold if strong else regular)[size], fill=color, anchor=anchor)

    def rotated_label(x, y, value, size=30):
        font = regular[size]
        box = font.getbbox(value)
        tile = Image.new("RGBA", (box[2] - box[0] + 20, box[3] - box[1] + 20), (255, 255, 255, 0))
        ImageDraw.Draw(tile).text((10 - box[0], 10 - box[1]), value, font=font, fill=INK)
        tile = tile.rotate(90, expand=True, resample=Image.Resampling.BICUBIC)
        image.paste(tile, (int(x * RENDER_SCALE - tile.width / 2), int(y * RENDER_SCALE - tile.height / 2)), tile)

    # Panel geometry.
    top, bottom = 65, 960
    left_l, left_r = 130, 1110
    right_l, right_r = 1390, 2305

    # ---------------- Wetherill concordia ----------------
    # Same close view used by the portal: the young cluster and the nearby
    # discordant analyses, without an overview inset.
    xmin, xmax = 1.00, 1.62
    ymin, ymax = 0.119, 0.1535

    def xmap(x):
        return left_l + (x - xmin) / (xmax - xmin) * (left_r - left_l)

    def ymap(y):
        return bottom - (y - ymin) / (ymax - ymin) * (bottom - top)

    # Grid, axes and numeric ticks.
    for x in (1.0, 1.2, 1.4, 1.6):
        px = xmap(x)
        draw.line((px, top, px, bottom), fill=GRID, width=2)
        draw.line((px, bottom, px, bottom + 10), fill=INK, width=2)
        label(px, bottom + 30, f"{x:.1f}", 22, anchor="ma", color=MUTED)
    for y in (0.120, 0.130, 0.140, 0.150):
        py = ymap(y)
        draw.line((left_l, py, left_r, py), fill=GRID, width=2)
        draw.line((left_l - 10, py, left_l, py), fill=INK, width=2)
        label(left_l - 18, py, f"{y:.3f}", 22, anchor="rm", color=MUTED)
    draw.line((left_l, top, left_l, bottom), fill=INK, width=3)
    draw.line((left_l, bottom, left_r, bottom), fill=INK, width=3)
    label((left_l + left_r) / 2, 1035, "²⁰⁷Pb/²³⁵U", 30)
    rotated_label(40, (top + bottom) / 2, "²⁰⁶Pb/²³⁸U", 30)

    # Concordia line, clipped naturally by the chosen plotting window.
    curve = []
    for age in np.linspace(700, 950, 500):
        x, y = concordia(float(age))
        if xmin <= x <= xmax and ymin <= y <= ymax:
            curve.append((xmap(x), ymap(y)))
    draw.line(curve, fill=INK, width=4, joint="curve")

    # Store age ticks; draw them again after the translucent ellipses so their
    # labels stay legible in the dense cluster.
    age_ticks = []
    for age in (750, 800, 900):
        x, y = concordia(age)
        if not (xmin <= x <= xmax and ymin <= y <= ymax):
            continue
        x0, y0 = xmap(x), ymap(y)
        xa, ya = concordia(age - 1)
        xb, yb = concordia(age + 1)
        dx, dy = xmap(xb) - xmap(xa), ymap(yb) - ymap(ya)
        length = math.hypot(dx, dy)
        nx, ny = -dy / length, dx / length
        if ny > 0:  # keep labels on the upper-left side of the curve
            nx, ny = -nx, -ny
        age_ticks.append((age, x0, y0, nx, ny))

    # Error ellipses in screen coordinates, including covariance and unequal axes.
    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    sx = (left_r - left_l) / (xmax - xmin)
    sy = -(bottom - top) / (ymax - ymin)

    def ellipse(row, fill, edge, alpha):
        covariance = np.array([
            [row["err_ratio_75"] ** 2, row["rho_3"] * row["err_ratio_75"] * row["err_ratio_68"]],
            [row["rho_3"] * row["err_ratio_75"] * row["err_ratio_68"], row["err_ratio_68"] ** 2],
        ], dtype=float)
        scale = np.diag([sx, sy])
        screen_cov = scale @ covariance @ scale.T
        values, vectors = np.linalg.eigh(screen_cov)
        order = values.argsort()[::-1]
        values, vectors = values[order], vectors[:, order]
        axes = 2.0 * np.sqrt(np.maximum(values, 0.0))  # semi-axes at 2 sigma
        center = np.array([xmap(row["ratio_75"]), ymap(row["ratio_68"])])
        points = []
        for theta in np.linspace(0, 2 * math.pi, 81):
            point = center + vectors @ np.array([axes[0] * math.cos(theta), axes[1] * math.sin(theta)])
            points.append((float(point[0]), float(point[1])))
        # One RGBA layer per ellipse makes overlapping fills accumulate, as
        # they do in the interactive portal, instead of replacing each other.
        ellipse_layer = Image.new("RGBA", image.size, (255, 255, 255, 0))
        ellipse_draw = ScaledDraw(ellipse_layer, RENDER_SCALE)
        ellipse_draw.polygon(points, fill=fill + (alpha,))
        ellipse_draw.line(points, fill=edge + (215,), width=2, joint="curve")
        overlay.alpha_composite(ellipse_layer)

    for row in rejected:
        ellipse(row, (193, 198, 195), (125, 133, 129), 48)
    for row in accepted:
        ellipse(row, (18, 110, 100), (12, 93, 85), 66)
    # Clip ellipse artwork strictly to the plotting frame.
    crop_box = tuple(round(value * RENDER_SCALE) for value in (left_l, top, left_r + 1, bottom + 1))
    cropped = overlay.crop(crop_box)
    image.paste(cropped, (left_l * RENDER_SCALE, top * RENDER_SCALE), cropped)
    draw = ScaledDraw(image, RENDER_SCALE)

    # Concordia and age labels remain on top of the translucent ellipses.
    draw.line(curve, fill=INK, width=4, joint="curve")
    for age, x0, y0, nx, ny in age_ticks:
        draw.line((x0 - nx * 8, y0 - ny * 8, x0 + nx * 8, y0 + ny * 8), fill=INK, width=3)
        draw.text(
            (x0 + nx * 27, y0 + ny * 27), str(age), font=regular[20],
            fill=MUTED, anchor="mm", stroke_width=3, stroke_fill="white",
        )

    # Explicit filter and population counts requested for the slide.
    info_x = left_l + 22
    label(info_x, top + 22, f"Фильтр: |D1| ≤ {DISCORDANCE_LIMIT:.0f}%", 22, anchor="la", strong=True)
    label(info_x, top + 58, f"Конкордантные зёрна: {len(accepted)}", 20, anchor="la", color=INK)
    label(info_x, top + 90, f"Дискордантные зёрна: {len(rejected)}", 20, anchor="la", color=MUTED)

    # ---------------- Horizontal weighted-mean diagram ----------------
    # "Horizontal" here means analyses arranged left-to-right, with age on Y.
    wymin, wymax = 765.0, 880.0

    def wy(age):
        return bottom - (age - wymin) / (wymax - wymin) * (bottom - top)

    x_positions = np.linspace(right_l + 17, right_r - 17, len(accepted))

    # Mean uncertainty band and mean line are horizontal (2 sigma).
    draw.rectangle((right_l, wy(mean + mean_2s), right_r, wy(mean - mean_2s)), fill=MEAN_BAND)
    for age in range(770, 881, 10):
        py = wy(age)
        draw.line((right_l, py, right_r, py), fill=GRID, width=2)
        draw.line((right_l - 10, py, right_l, py), fill=INK, width=2)
        label(right_l - 18, py, str(age), 22, anchor="rm", color=MUTED)
    draw.line((right_l, top, right_l, bottom), fill=INK, width=3)
    draw.line((right_l, bottom, right_r, bottom), fill=INK, width=3)
    draw.line((right_l, wy(mean), right_r, wy(mean)), fill=MEAN, width=3)

    for px, row in zip(x_positions, accepted):
        lo = wy(row["age_68"] - 2 * row["err_age_68"])
        hi = wy(row["age_68"] + 2 * row["err_age_68"])
        center = wy(row["age_68"])
        draw.line((px, lo, px, hi), fill="#438d85", width=3)
        draw.line((px - 5, lo, px + 5, lo), fill="#438d85", width=2)
        draw.line((px - 5, hi, px + 5, hi), fill="#438d85", width=2)
        draw.ellipse((px - 6, center - 6, px + 6, center + 6), fill=GREEN, outline=GREEN_EDGE, width=2)
        short_id = str(row["sid"]).rsplit("_", 1)[-1]
        label(px, bottom + 30, short_id, 19, anchor="ma", color=MUTED)

    rotated_label(1262, (top + bottom) / 2, "Возраст ²⁰⁶Pb/²³⁸U, млн лет", 30)
    label((right_l + right_r) / 2, 1035, "Номер анализа", 26)
    label(right_l, 25, f"{mean:.1f} ± {mean_2s:.1f} млн лет (2σ)", 22, anchor="la", strong=True)
    label(right_r, 25, f"MSWD = {mswd:.2f}; n = {len(accepted)}", 20, anchor="ra", color=MUTED)

    OUT_PNG.parent.mkdir(parents=True, exist_ok=True)
    image.save(OUT_PNG, dpi=(600, 600), optimize=True)

    pdf = canvas.Canvas(str(OUT_PDF), pagesize=(30 * cm, 14.25 * cm))
    pdf.drawImage(str(OUT_PNG), 0, 0, width=30 * cm, height=14.25 * cm, preserveAspectRatio=True, mask="auto")
    pdf.showPage()
    pdf.save()

    print(f"PNG: {OUT_PNG}")
    print(f"PDF: {OUT_PDF}")
    print(f"accepted={len(accepted)} rejected={len(rejected)} mean={mean:.3f} 2s={mean_2s:.3f} mswd={mswd:.3f}")


if __name__ == "__main__":
    main()
