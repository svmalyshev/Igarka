#!/usr/bin/env python3
"""Full-spectrum presentation KDE for SM25-16, excluding ages below 600 Ma."""

from __future__ import annotations

import csv
import math
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

from make_gubinskaya_kde_figure import ScaledDraw


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "upb" / "SM25-16_upb_audit.csv"
OUT_PNG = ROOT / "assets" / "SM25-16_KDE_full_30cm_600dpi.png"
OUT_PDF = ROOT / "assets" / "SM25-16_KDE_full_30cm_600dpi.pdf"

WIDTH, HEIGHT, SCALE = 2362, 1122, 3
XMIN, XMAX = 700.0, 2800.0
BANDWIDTH = 25.0
MIN_AGE = 600.0

FONT = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
INK = "#27312c"
MUTED = "#68736e"
GRID = "#dfe5e1"
TEAL = "#126e64"
TEAL_FILL = (18, 110, 100, 45)


def load_ages() -> tuple[np.ndarray, int]:
    with SOURCE.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    excluded_young = sum(float(row["best_age"]) < MIN_AGE for row in rows)
    ages = np.array(
        [
            float(row["best_age"])
            for row in rows
            # Use every analysis that passes the analytical discordance filter.
            # The portal's broader `accepted` flag also applies an old contextual
            # exclusion to three 753–856 Ma grains, which the present figure must retain.
            if row["analytical_pass"].lower() == "true" and float(row["best_age"]) >= MIN_AGE
        ],
        dtype=float,
    )
    return ages, excluded_young


def kde(ages: np.ndarray, grid: np.ndarray) -> np.ndarray:
    z = (grid[:, None] - ages[None, :]) / BANDWIDTH
    return np.exp(-0.5 * z * z).sum(axis=1) / (len(ages) * BANDWIDTH * math.sqrt(2 * math.pi))


def local_peaks(grid: np.ndarray, density: np.ndarray) -> list[tuple[float, float]]:
    return [
        (float(grid[index]), float(density[index]))
        for index in range(1, len(grid) - 1)
        if density[index] > density[index - 1]
        and density[index] >= density[index + 1]
        and density[index] >= density.max() * 0.01
    ]


def main() -> None:
    ages, excluded_young = load_ages()
    grid = np.linspace(XMIN, XMAX, 8401)
    density = kde(ages, grid)
    peaks = local_peaks(grid, density)

    image = Image.new("RGB", (WIDTH * SCALE, HEIGHT * SCALE), "white")
    draw = ScaledDraw(image, SCALE)
    regular = {size: ImageFont.truetype(FONT, size * SCALE) for size in (18, 20, 21, 22, 24, 26, 28, 30)}
    bold = {size: ImageFont.truetype(FONT_BOLD, size * SCALE) for size in (20, 22, 24, 26, 28, 30)}

    def label(x, y, value, size=24, anchor="mm", color=INK, strong=False, halo=False):
        draw.text(
            (x, y), str(value), font=(bold if strong else regular)[size], fill=color, anchor=anchor,
            stroke_width=3 if halo else 0, stroke_fill="white" if halo else None,
        )

    left, right, top, bottom = 165, 2295, 120, 905
    ymin = -0.00072
    ymax = math.ceil(density.max() / 0.004) * 0.004

    def sx(value):
        return left + (value - XMIN) / (XMAX - XMIN) * (right - left)

    def sy(value):
        return bottom - (value - ymin) / (ymax - ymin) * (bottom - top)

    for age in range(750, 2751, 250):
        px = sx(age)
        draw.line((px, top, px, sy(0)), fill=GRID, width=2)
        draw.line((px, sy(0), px, sy(0) + 10), fill=INK, width=2)
        label(px, bottom + 32, str(age), 22, anchor="ma", color=MUTED)
    y_step = 0.004
    for value in np.arange(0, ymax + y_step / 2, y_step):
        py = sy(float(value))
        draw.line((left, py, right, py), fill=GRID, width=2)
        draw.line((left - 10, py, left, py), fill=INK, width=2)
        label(left - 18, py, f"{value:.3f}", 21, anchor="rm", color=MUTED)
    draw.line((left, top, left, sy(0)), fill=INK, width=3)
    draw.line((left, sy(0), right, sy(0)), fill=INK, width=3)

    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    odraw = ScaledDraw(overlay, SCALE)
    polygon = [(sx(grid[0]), sy(0))]
    polygon.extend((sx(x), sy(y)) for x, y in zip(grid, density))
    polygon.append((sx(grid[-1]), sy(0)))
    odraw.polygon(polygon, fill=TEAL_FILL)
    image.paste(overlay, (0, 0), overlay)
    draw = ScaledDraw(image, SCALE)

    draw.line([(sx(x), sy(y)) for x, y in zip(grid, density)], fill=TEAL, width=5, joint="curve")
    for age in ages:
        if XMIN <= age <= XMAX:
            draw.line((sx(age), sy(-0.00006), sx(age), sy(-0.00038)), fill=TEAL, width=2)

    for age, height in peaks:
        # With only three Neoproterozoic grains, the two local KDE maxima below
        # 1 Ga are numerical features of smoothing, not robust age populations.
        if age < 1000:
            continue
        draw.ellipse((sx(age) - 5, sy(height) - 5, sx(age) + 5, sy(height) + 5), fill=TEAL)
        label(sx(age), sy(height) - 23, f"{age:.0f}", 22, anchor="ms", color=TEAL, strong=True, halo=True)

    young_ages = sorted(age for age in ages if age < 1000)
    if young_ages:
        ages_text = "; ".join(str(f"{age:.1f}").replace(".", ",") for age in young_ages)
        label(
            left + 15, sy(0.00062),
            f"n = {len(young_ages)}: {ages_text} млн лет",
            20, anchor="ls", color=TEAL, halo=True,
        )

    legend_y = 45
    draw.line((left, legend_y, left + 70, legend_y), fill=TEAL, width=5)
    label(left + 86, legend_y, f"SM25-16 — Малышев, Пасенко (n = {len(ages)})", 22, anchor="lm")
    label(
        right, 92,
        f"Фильтр: |D| ≤ 10% · KDE h = {BANDWIDTH:.0f} млн лет",
        20, anchor="ra", color=MUTED,
    )
    label((left + right) / 2, 1050, "Возраст, млн лет", 30)

    value = "Плотность вероятности"
    font = regular[28]
    box = font.getbbox(value)
    tile = Image.new("RGBA", (box[2] - box[0] + 40, box[3] - box[1] + 40), (255, 255, 255, 0))
    ImageDraw.Draw(tile).text((20 - box[0], 20 - box[1]), value, font=font, fill=INK)
    tile = tile.rotate(90, expand=True, resample=Image.Resampling.BICUBIC)
    image.paste(tile, (int(45 * SCALE - tile.width / 2), int((top + sy(0)) / 2 * SCALE - tile.height / 2)), tile)

    OUT_PNG.parent.mkdir(parents=True, exist_ok=True)
    image.save(OUT_PNG, dpi=(600, 600), optimize=True)
    pdf = canvas.Canvas(str(OUT_PDF), pagesize=(30 * cm, 14.25 * cm))
    pdf.drawImage(str(OUT_PNG), 0, 0, width=30 * cm, height=14.25 * cm, preserveAspectRatio=True, mask="auto")
    pdf.showPage()
    pdf.save()

    print(f"PNG: {OUT_PNG}")
    print(f"PDF: {OUT_PDF}")
    print(f"accepted age >= 600 Ma: {len(ages)}; all rows below 600 Ma excluded: {excluded_young}")
    print("peaks:", ", ".join(f"{age:.1f}" for age, _ in peaks))


if __name__ == "__main__":
    main()
