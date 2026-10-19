#!/usr/bin/env python3
"""Presentation KDE comparison for the Gubinskaya Formation zircon datasets."""

from __future__ import annotations

import csv
import math
import os
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parent
OURS = ROOT / "upb" / "SM25-05_upb_audit.csv"
KOCHNEV = ROOT / "upb" / "kochnev2022_41812_audit.csv"

# 30.0 x 14.25 cm at 600 dpi, rendered natively rather than enlarged later.
WIDTH, HEIGHT, SCALE = 2362, 1122, 3
FONT = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"

INK = "#27312c"
MUTED = "#68736e"
GRID = "#dfe5e1"
TEAL = "#126e64"
TEAL_FILL = (18, 110, 100, 45)
BLUE = "#4f729f"
BLUE_FILL = (79, 114, 159, 34)
VIEW = os.environ.get("KDE_VIEW", "full")
BANDWIDTH = 15.0 if VIEW == "young_peaks" else 25.0
if VIEW == "young_peaks":
    XMIN, XMAX = 700.0, 900.0
elif VIEW == "young":
    XMIN, XMAX = 680.0, 1000.0
else:
    XMIN, XMAX = 600.0, 3200.0


class ScaledDraw:
    def __init__(self, image: Image.Image, scale: int):
        self.raw = ImageDraw.Draw(image)
        self.scale = scale

    def _point(self, p):
        return tuple(float(v) * self.scale for v in p)

    def _points(self, values):
        if len(values) and isinstance(values[0], (tuple, list, np.ndarray)):
            return [self._point(p) for p in values]
        return tuple(float(v) * self.scale for v in values)

    def line(self, xy, fill=None, width=1, joint=None):
        self.raw.line(self._points(xy), fill=fill, width=max(1, round(width * self.scale)), joint=joint)

    def rectangle(self, xy, fill=None, outline=None, width=1):
        self.raw.rectangle(self._points(xy), fill=fill, outline=outline, width=max(1, round(width * self.scale)))

    def polygon(self, xy, fill=None):
        self.raw.polygon(self._points(xy), fill=fill)

    def ellipse(self, xy, fill=None, outline=None, width=1):
        self.raw.ellipse(self._points(xy), fill=fill, outline=outline, width=max(1, round(width * self.scale)))

    def text(self, xy, text, font=None, fill=None, anchor=None, stroke_width=0, stroke_fill=None):
        self.raw.text(
            self._point(xy), text, font=font, fill=fill, anchor=anchor,
            stroke_width=round(stroke_width * self.scale), stroke_fill=stroke_fill,
        )


def load_ages(path: Path, age_field: str) -> np.ndarray:
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    return np.array([float(row[age_field]) for row in rows if row["accepted"].lower() == "true"], dtype=float)


def kde(ages: np.ndarray, grid: np.ndarray) -> np.ndarray:
    z = (grid[:, None] - ages[None, :]) / BANDWIDTH
    return np.exp(-0.5 * z * z).sum(axis=1) / (len(ages) * BANDWIDTH * math.sqrt(2 * math.pi))


def peaks(grid: np.ndarray, density: np.ndarray) -> list[tuple[float, float]]:
    found = [
        (float(grid[i]), float(density[i]))
        for i in range(1, len(grid) - 1)
        if density[i] > density[i - 1] and density[i] >= density[i + 1]
    ]
    return [item for item in found if XMIN + 10 <= item[0] <= XMAX - 10 and item[1] >= density.max() * 0.02]


def peak_components(
    ages: np.ndarray,
    grid: np.ndarray,
    density: np.ndarray,
    peak_list: list[tuple[float, float]],
) -> list[dict[str, float | int]]:
    """Partition the selected range at KDE minima and calculate peak weights."""
    in_range = ages[(ages >= XMIN) & (ages <= XMAX)]
    peak_indices = [int(np.argmin(np.abs(grid - age))) for age, _ in peak_list]
    minima = []
    for first, second in zip(peak_indices[:-1], peak_indices[1:]):
        minima.append(first + int(np.argmin(density[first : second + 1])))
    bounds = [XMIN] + [float(grid[index]) for index in minima] + [XMAX]
    components = []
    for index, (age, height) in enumerate(peak_list):
        lo, hi = bounds[index], bounds[index + 1]
        if index < len(peak_list) - 1:
            count = int(np.sum((in_range >= lo) & (in_range < hi)))
        else:
            count = int(np.sum((in_range >= lo) & (in_range <= hi)))
        components.append(
            {
                "age": age,
                "height": height,
                "lo": lo,
                "hi": hi,
                "count": count,
                "probability": 100.0 * count / len(in_range),
                "total": len(in_range),
            }
        )
    return components


def main() -> None:
    show_kochnev = os.environ.get("KDE_SHOW_KOCHNEV", "1") != "0"
    view_label = "young-peaks" if VIEW == "young_peaks" else ("young" if VIEW == "young" else "full")
    stem = (
        f"Gubinskaya_KDE_matched_{view_label}_slide2_SM25-05_plus_1812-2"
        if show_kochnev else
        f"Gubinskaya_KDE_matched_{view_label}_slide1_SM25-05"
    )
    out_png = ROOT / "assets" / f"{stem}_30cm_600dpi.png"
    out_pdf = ROOT / "assets" / f"{stem}_30cm_600dpi.pdf"
    ours = load_ages(OURS, "best_age")
    kochnev = load_ages(KOCHNEV, "age")
    grid = np.linspace(XMIN, XMAX, 2401)
    ours_density = kde(ours, grid)
    kochnev_density = kde(kochnev, grid)
    ours_peaks = peaks(grid, ours_density)
    kochnev_peaks = peaks(grid, kochnev_density)
    ours_components = peak_components(ours, grid, ours_density, ours_peaks)
    kochnev_components = peak_components(kochnev, grid, kochnev_density, kochnev_peaks)

    image = Image.new("RGB", (WIDTH * SCALE, HEIGHT * SCALE), "white")
    draw = ScaledDraw(image, SCALE)
    regular = {s: ImageFont.truetype(FONT, s * SCALE) for s in (18, 20, 21, 22, 24, 26, 28, 30, 32)}
    bold = {s: ImageFont.truetype(FONT_BOLD, s * SCALE) for s in (20, 22, 24, 26, 28, 30, 32)}

    def label(x, y, value, size=24, anchor="mm", color=INK, strong=False, halo=False):
        draw.text(
            (x, y), str(value), font=(bold if strong else regular)[size], fill=color, anchor=anchor,
            stroke_width=3 if halo else 0, stroke_fill="white" if halo else None,
        )

    left, right, top, bottom = 165, 2295, 120, 905
    ymin, ymax = -0.00125, max(ours_density.max(), kochnev_density.max()) * 1.22

    def sx(value):
        return left + (value - XMIN) / (XMAX - XMIN) * (right - left)

    def sy(value):
        return bottom - (value - ymin) / (ymax - ymin) * (bottom - top)

    # Axes and grids.
    if VIEW == "young_peaks":
        x_ticks = range(700, 901, 50)
    elif VIEW == "young":
        x_ticks = range(700, 1001, 50)
    else:
        x_ticks = range(750, 3001, 250)
    for age in x_ticks:
        px = sx(age)
        draw.line((px, top, px, sy(0)), fill=GRID, width=2)
        draw.line((px, sy(0), px, sy(0) + 10), fill=INK, width=2)
        label(px, bottom + 32, str(age), 22, anchor="ma", color=MUTED)
    for density in (0.0, 0.002, 0.004, 0.006):
        py = sy(density)
        draw.line((left, py, right, py), fill=GRID, width=2)
        draw.line((left - 10, py, left, py), fill=INK, width=2)
        label(left - 18, py, f"{density:.3f}", 21, anchor="rm", color=MUTED)
    draw.line((left, top, left, sy(0)), fill=INK, width=3)
    draw.line((left, sy(0), right, sy(0)), fill=INK, width=3)

    # Transparent filled KDEs.
    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    odraw = ScaledDraw(overlay, SCALE)
    # The green area is identical in both slide frames. Kochnev is added as a
    # line only, so the first dataset does not change colour between slides.
    for density, fill in ((ours_density, TEAL_FILL),):
        polygon = [(sx(grid[0]), sy(0))]
        polygon.extend((sx(x), sy(y)) for x, y in zip(grid, density))
        polygon.append((sx(grid[-1]), sy(0)))
        odraw.polygon(polygon, fill=fill)
    image.paste(overlay, (0, 0), overlay)
    draw = ScaledDraw(image, SCALE)

    # Curves on top of the fills.
    if show_kochnev:
        draw.line([(sx(x), sy(y)) for x, y in zip(grid, kochnev_density)], fill=BLUE, width=4, joint="curve")
    draw.line([(sx(x), sy(y)) for x, y in zip(grid, ours_density)], fill=TEAL, width=5, joint="curve")

    # Rug marks in two separate lanes beneath the probability axis.
    ours_in = sorted(age for age in ours if XMIN <= age <= XMAX)
    kochnev_in = sorted(age for age in kochnev if XMIN <= age <= XMAX)
    for age in ours_in:
        draw.line((sx(age), sy(-0.00006), sx(age), sy(-0.00038)), fill=TEAL, width=2)
    if show_kochnev:
        for age in kochnev_in:
            draw.line((sx(age), sy(-0.00060), sx(age), sy(-0.00092)), fill=BLUE, width=2)

    # Label every resolved local maximum; close peaks from the two datasets are
    # offset in opposite directions so both values remain readable.
    for component in ours_components:
        age, density = float(component["age"]), float(component["height"])
        close = any(abs(age - other_age) < 18 for other_age, _ in kochnev_peaks)
        xoff = -18 if close else 0
        anchor = "ra" if close else "ms"
        draw.ellipse((sx(age) - 5, sy(density) - 5, sx(age) + 5, sy(density) + 5), fill=TEAL)
        if VIEW == "young_peaks":
            probability = str(f"{float(component['probability']):.1f}").replace(".", ",")
            label(sx(age) + xoff, sy(density) - 30, f"~{age:.0f} млн лет", 22, anchor=anchor, color=TEAL, strong=True, halo=True)
            label(
                sx(age) + xoff, sy(density) - 4,
                f"n = {component['count']}; p = {probability}%", 20,
                anchor=anchor, color=TEAL, halo=True,
            )
        else:
            label(sx(age) + xoff, sy(density) - 23, f"{age:.0f}", 22, anchor=anchor, color=TEAL, strong=True, halo=True)
    if show_kochnev:
        for component in kochnev_components:
            age, density = float(component["age"]), float(component["height"])
            close = any(abs(age - other_age) < 18 for other_age, _ in ours_peaks)
            xoff = 20 if close else 0
            anchor = "la" if close else "ms"
            draw.ellipse((sx(age) - 5, sy(density) - 5, sx(age) + 5, sy(density) + 5), fill=BLUE)
            if VIEW == "young_peaks":
                probability = str(f"{float(component['probability']):.1f}").replace(".", ",")
                label(sx(age) + xoff, sy(density) + 28, f"~{age:.0f} млн лет", 20, anchor=anchor, color=BLUE, strong=True, halo=True)
                label(
                    sx(age) + xoff, sy(density) + 53,
                    f"n = {component['count']}; p = {probability}%", 20,
                    anchor=anchor, color=BLUE, halo=True,
                )
            else:
                label(sx(age) + xoff, sy(density) + 29, f"{age:.0f}", 20, anchor=anchor, color=BLUE, halo=True)

    # Direct legend and methodological note.
    legend_y = 45
    draw.line((left, legend_y, left + 70, legend_y), fill=TEAL, width=5)
    label(
        left + 86, legend_y,
        f"SM25-05 — Малышев, Пасенко (n = {len(ours)})",
        22, anchor="lm",
    )
    if show_kochnev:
        draw.line((1110, legend_y, 1180, legend_y), fill=BLUE, width=4)
        label(
            1196, legend_y,
            f"1812-2 — Кочнев и др., 2022 (n = {len(kochnev)})",
            22, anchor="lm",
        )
    label(right, 92, f"Фильтр: |D| ≤ 10% · KDE h = {BANDWIDTH:.0f} млн лет", 20, anchor="ra", color=MUTED)

    label((left + right) / 2, 1050, "Возраст, млн лет", 30)
    # Rotated y-axis label.
    value = "Плотность вероятности"
    font = regular[28]
    box = font.getbbox(value)
    tile = Image.new("RGBA", (box[2] - box[0] + 40, box[3] - box[1] + 40), (255, 255, 255, 0))
    ImageDraw.Draw(tile).text((20 - box[0], 20 - box[1]), value, font=font, fill=INK)
    tile = tile.rotate(90, expand=True, resample=Image.Resampling.BICUBIC)
    image.paste(tile, (int(45 * SCALE - tile.width / 2), int((top + sy(0)) / 2 * SCALE - tile.height / 2)), tile)

    out_png.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_png, dpi=(600, 600), optimize=True)
    pdf = canvas.Canvas(str(out_pdf), pagesize=(30 * cm, 14.25 * cm))
    pdf.drawImage(str(out_png), 0, 0, width=30 * cm, height=14.25 * cm, preserveAspectRatio=True, mask="auto")
    pdf.showPage()
    pdf.save()

    print(f"PNG: {out_png}")
    print(f"PDF: {out_pdf}")
    print("SM25-05 peaks:", ", ".join(f"{age:.1f}" for age, _ in ours_peaks))
    print("1812-2 peaks:", ", ".join(f"{age:.1f}" for age, _ in kochnev_peaks))


if __name__ == "__main__":
    main()
