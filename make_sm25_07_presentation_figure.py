#!/usr/bin/env python3
"""Build a 10 cm wide SVG with TAS and PM-normalized diagrams for SM25-07."""

from __future__ import annotations

import csv
import html
import math
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parent
DATA_ROOT = ROOT.parent
OUT_SVG = ROOT / "assets" / "SM25-07_TAS_multielement_10cm.svg"
OUT_PNG = ROOT / "assets" / "SM25-07_TAS_multielement_10cm.png"
OUT_PDF = ROOT / "assets" / "SM25-07_TAS_multielement_10cm.pdf"

# 10.0 × 15.2 cm at 300 dpi. The SVG keeps physical dimensions in centimetres.
WIDTH, HEIGHT = 1181, 1795

TAS_FIELDS = [
    ("Pc", [(41, 3), (45, 5), (45, 0), (41, 0)]),
    ("B", [(45, 5), (52, 5), (52, 0), (45, 0)]),
    ("F", [(35, 9), (37, 14), (52.5, 18), (52.5, 14), (48.4, 11.5), (45, 9.4), (41, 7), (41, 3), (37, 3)]),
    ("O1", [(52, 0), (52, 5), (57, 5.9), (57, 0)]),
    ("O2", [(57, 0), (57, 5.9), (63, 7), (63, 0)]),
    ("O3", [(63, 0), (63, 7), (69, 8), (77.3, 0)]),
    ("Ph", [(52.5, 14), (52.5, 18), (57, 18), (63, 16.2), (61, 13.5), (57.6, 11.7)]),
    ("R", [(69, 8), (69, 13), (85.9, 6.8), (87.5, 4.7), (77.3, 0)]),
    ("S1", [(45, 5), (49.4, 7.3), (52, 5)]),
    ("S2", [(49.4, 7.3), (53, 9.3), (57, 5.9), (52, 5)]),
    ("S3", [(53, 9.3), (57.6, 11.7), (61, 8.6), (63, 7), (57, 5.9)]),
    ("T1", [(57.6, 11.7), (61, 13.5), (63, 16.2), (69, 13), (61, 8.6)]),
    ("T2", [(61, 8.6), (69, 13), (69, 8), (63, 7)]),
    ("U1", [(41, 3), (41, 7), (45, 9.4), (49.4, 7.3), (45, 5), (45, 3)]),
    ("U2", [(45, 9.4), (48.4, 11.5), (53, 9.3), (49.4, 7.3)]),
    ("U3", [(48.4, 11.5), (52.5, 14), (57.6, 11.7), (53, 9.3)]),
]

TAS_LABELS = [
    (42.5, 13.2, ("Фоидит",), 21),
    (43.1, 1.25, ("Пикро-", "базальт"), 19),
    (48.4, 1.35, ("Базальт",), 21),
    (54.3, 1.65, ("Базальтовый", "андезит"), 18),
    (60.2, 4.6, ("Андезит",), 21),
    (66.4, 3.0, ("Дацит",), 21),
    (73.1, 7.1, ("Риолит",), 21),
    (57.0, 14.8, ("Фонолит",), 21),
    (64.2, 12.3, ("Трахит",), 21),
    (61.8, 8.7, ("Трахидацит",), 20),
    (48.0, 6.0, ("Трахи-", "базальт"), 18),
    (53.3, 7.3, ("Базальтовый", "трахиандезит"), 17),
    (57.2, 9.2, ("Трахи-", "андезит"), 18),
    (42.9, 5.6, ("Тефрит /", "базанит"), 18),
    (49.5, 9.2, ("Фоно-", "тефрит"), 18),
    (53.2, 11.7, ("Тефри-", "фонолит"), 18),
]


def key(value: object) -> str:
    return "".join(str(value or "").upper().split())


def read_major() -> dict[str, float]:
    ws = load_workbook(DATA_ROOT / "Пасенко_Малышев_19_проб.xlsx", data_only=True, read_only=True)["Лист1"]
    headers = [str(cell.value or "").replace(" (%)", "").strip() for cell in ws[4]]
    for row in ws.iter_rows(min_row=5, values_only=True):
        if key(row[0]) == "SM25-07":
            return {headers[i]: float(v) for i, v in enumerate(row) if i and v is not None}
    raise RuntimeError("SM25-07 not found in major-element data")


def read_trace() -> dict[str, float]:
    ws = load_workbook(DATA_ROOT / "igarka_trace.xlsx", data_only=True, read_only=True)["Sheet1"]
    headers = [str(cell.value or "").strip() for cell in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if key(row[0]) == "SM25-07":
            return {headers[i]: float(v) for i, v in enumerate(row) if i >= 2 and headers[i] and v is not None}
    raise RuntimeError("SM25-07 not found in trace-element data")


def read_csv(path: Path) -> tuple[list[str], dict[str, dict[str, float]]]:
    with path.open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.reader(stream))
    elements = [value.strip() for value in rows[0][1:]]
    datasets = {}
    for row in rows[1:]:
        if row and row[0].strip():
            datasets[row[0].strip()] = {
                element: float(value)
                for element, value in zip(elements, row[1:])
                if element and str(value).strip()
            }
    return elements, datasets


def text(x, y, value, size=28, anchor="middle", weight=400, fill="#27312c", rotate=None):
    transform = f' transform="rotate({rotate} {x:.2f} {y:.2f})"' if rotate is not None else ""
    return (
        f'<text x="{x:.2f}" y="{y:.2f}" text-anchor="{anchor}" font-size="{size}" '
        f'font-weight="{weight}" fill="{fill}"{transform}>{html.escape(str(value))}</text>'
    )


def multiline_svg(x, y, lines, size, fill="#3f4c46"):
    start = y - (len(lines) - 1) * size * 0.54
    tspans = "".join(
        f'<tspan x="{x:.2f}" y="{start + i * size * 1.08:.2f}">{html.escape(value)}</tspan>'
        for i, value in enumerate(lines)
    )
    return f'<text text-anchor="middle" font-size="{size}" fill="{fill}">{tspans}</text>'


def line(x1, y1, x2, y2, stroke="#53605a", width=2, dash=None):
    dashed = f' stroke-dasharray="{dash}"' if dash else ""
    return f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" stroke="{stroke}" stroke-width="{width}"{dashed}/>'


def path(points, stroke, width, dash=None, markers=False):
    d = " ".join(("M" if i == 0 else "L") + f"{x:.2f},{y:.2f}" for i, (x, y) in enumerate(points))
    dashed = f' stroke-dasharray="{dash}"' if dash else ""
    out = [f'<path d="{d}" fill="none" stroke="{stroke}" stroke-width="{width}" stroke-linejoin="round" stroke-linecap="round"{dashed}/>']
    if markers:
        out.extend(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="5" fill="white" stroke="{stroke}" stroke-width="3"/>' for x, y in points)
    return "".join(out)


FONT = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"


def render_png(major, trace, elements, pm, refs) -> None:
    image = Image.new("RGB", (WIDTH, HEIGHT), "white")
    draw = ImageDraw.Draw(image)
    regular = {size: ImageFont.truetype(FONT, size) for size in (17, 18, 19, 20, 21, 22, 23, 24, 25, 28, 30, 36)}
    bold = {size: ImageFont.truetype(FONT_BOLD, size) for size in (25, 30, 36)}

    def label(x, y, value, size=28, anchor="mm", color="#27312c", is_bold=False):
        draw.text((x, y), str(value), font=(bold if is_bold else regular)[size], fill=color, anchor=anchor)

    def rotated_label(x, y, value, size=28, angle=90, color="#27312c"):
        font = regular[size]
        box = font.getbbox(str(value))
        tile = Image.new("RGBA", (box[2] - box[0] + 12, box[3] - box[1] + 12), (255, 255, 255, 0))
        td = ImageDraw.Draw(tile)
        td.text((6 - box[0], 6 - box[1]), str(value), font=font, fill=color)
        tile = tile.rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
        image.paste(tile, (round(x - tile.width / 2), round(y - tile.height / 2)), tile)

    def multiline_label(x, y, lines, size, color="#3f4c46"):
        draw.multiline_text(
            (x, y),
            "\n".join(lines),
            font=regular[size],
            fill=color,
            anchor="mm",
            align="center",
            spacing=1,
        )

    def dashed(coords, fill, width=4, dash=18, gap=11):
        for (x1, y1), (x2, y2) in zip(coords, coords[1:]):
            dx, dy = x2 - x1, y2 - y1
            length = math.hypot(dx, dy)
            if length == 0:
                continue
            ux, uy = dx / length, dy / length
            pos = 0.0
            while pos < length:
                end = min(length, pos + dash)
                draw.line((x1 + ux * pos, y1 + uy * pos, x1 + ux * end, y1 + uy * end), fill=fill, width=width)
                pos += dash + gap

    # TAS.
    left, right, top, bottom = 142, 1148, 86, 828
    xmap = lambda value: left + (value - 38) / 39 * (right - left)
    ymap = lambda value: bottom - value / 16 * (bottom - top)
    label(left, 39, "а) TAS · Le Bas et al. (1986)", 36, "lm", is_bold=True)
    for tick in (0, 5, 10, 15):
        y = ymap(tick)
        draw.line((left, y, right, y), fill="#dfe5e2", width=2)
        label(left - 17, y, tick, 25, "rm", "#4d5a54")
    for tick in (40, 50, 60, 70):
        x = xmap(tick)
        draw.line((x, top, x, bottom), fill="#edf0ee", width=2)
        label(x, bottom + 30, tick, 25, "mm", "#4d5a54")
    for index, (_, points) in enumerate(TAS_FIELDS):
        coords = [(xmap(max(38, min(77, x))), ymap(max(0, min(16, y)))) for x, y in points]
        draw.polygon(coords, fill="#eef3f0" if index % 2 == 0 else "#f8f9f8", outline="#87938d", width=2)
    for lx, ly, lines, size in TAS_LABELS:
        multiline_label(xmap(lx), ymap(ly), lines, size)
    oxides = ["Na2O", "MgO", "Al2O3", "SiO2", "P2O5", "K2O", "CaO", "TiO2", "MnO", "Fe2O3"]
    factor = 100 / sum(major.get(name, 0) for name in oxides)
    sio2 = major["SiO2"] * factor
    alkalis = (major["Na2O"] + major["K2O"]) * factor
    px, py = xmap(sio2), ymap(alkalis)
    diamond = [(px, py - 15), (px + 15, py), (px, py + 15), (px - 15, py)]
    draw.polygon(diamond, fill="#c74735", outline="white", width=4)
    label(px - 20, py - 22, "SM25-07", 30, "rs", "#66291f", True)
    draw.line((left, top, left, bottom), fill="#39443f", width=3)
    draw.line((left, bottom, right, bottom), fill="#39443f", width=3)
    label((left + right) / 2, bottom + 76, "SiO₂, мас. %", 30)
    rotated_label(42, (top + bottom) / 2, "Na₂O + K₂O, мас. %", 30, 90)

    # Multielement panel.
    mleft, mright, mtop, mbottom = 142, 1148, 1000, 1560
    mx = lambda i: mleft + i / (len(elements) - 1) * (mright - mleft)
    my = lambda value: mbottom - (math.log10(value) + 1) / 4 * (mbottom - mtop)
    label(mleft, 947, "б) Мультиэлементная диаграмма · PM-normalized", 36, "lm", is_bold=True)
    for value in (0.1, 1, 10, 100, 1000):
        y = my(value)
        draw.line((mleft, y, mright, y), fill="#dbe1de", width=2)
        label(mleft - 17, y, f"{value:g}", 25, "rm", "#4d5a54")
    draw.line((mleft, mtop, mleft, mbottom), fill="#39443f", width=3)
    draw.line((mleft, mbottom, mright, mbottom), fill="#39443f", width=3)
    for i, element in enumerate(elements):
        rotated_label(mx(i), mbottom + 51, element, 23, 90, "#39443f")
    series = [
        ("N-MORB", refs["N-MORB (Gale et al., 2013)"], "#536a93", True, 4, False),
        ("OIB", refs["OIB (Sun&McDoniugh, 1989)"], "#b06b2b", True, 4, False),
        ("SM25-07", trace, "#126e64", False, 7, True),
    ]
    for series_name, values, color, is_dashed, width, markers in series:
        coords = [(mx(i), my(values[element] / pm[element])) for i, element in enumerate(elements)]
        if is_dashed:
            dashed(coords, color, width)
        else:
            draw.line(coords, fill=color, width=width, joint="curve")
        if markers:
            for x, y in coords:
                draw.ellipse((x - 5, y - 5, x + 5, y + 5), fill="white", outline=color, width=3)
    legend_x, legend_y = 875, 1025
    for i, (series_name, _, color, is_dashed, width, _) in enumerate(series):
        y = legend_y + i * 38
        coords = [(legend_x, y), (legend_x + 82, y)]
        dashed(coords, color, width) if is_dashed else draw.line((*coords[0], *coords[1]), fill=color, width=width)
        label(legend_x + 96, y, series_name, 25, "lm", is_bold=(series_name == "SM25-07"))
    rotated_label(42, (mtop + mbottom) / 2, "Порода / примитивная мантия", 30, 90)
    label((mleft + mright) / 2, 1700, "Элементы", 30)
    label(mleft, 1755, "PM: McDonough & Sun (1989); N-MORB: Gale et al. (2013); OIB: Sun & McDonough (1989)", 22, "lm", "#53605a")
    image.save(OUT_PNG, dpi=(300, 300), optimize=True)


def render_pdf() -> None:
    document = canvas.Canvas(str(OUT_PDF), pagesize=(10 * cm, 15.2 * cm))
    document.drawImage(str(OUT_PNG), 0, 0, width=10 * cm, height=15.2 * cm, mask="auto")
    document.showPage()
    document.save()


def main() -> None:
    major, trace = read_major(), read_trace()
    elements, pm_rows = read_csv(DATA_ROOT / "PM_1989.csv")
    _, refs = read_csv(DATA_ROOT / "morb.csv")
    pm = pm_rows["Primitive Mantle McDonough & Sun 1989"]

    sample = dict(trace)
    sample.update(K=major["K2O"] * 8301.5, P=major["P2O5"] * 4364.2, Ti=major["TiO2"] * 5993.4)

    svg = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="10cm" height="15.2cm" viewBox="0 0 {WIDTH} {HEIGHT}">',
        '<rect width="100%" height="100%" fill="white"/>',
        '<g font-family="Arial, Helvetica, sans-serif">',
    ]

    # TAS panel.
    left, right, top, bottom = 142, 1148, 86, 828
    xmap = lambda value: left + (value - 38) / 39 * (right - left)
    ymap = lambda value: bottom - value / 16 * (bottom - top)
    svg.append(text(left, 47, "а) TAS · Le Bas et al. (1986)", 36, "start", 700))

    for tick in (0, 5, 10, 15):
        y = ymap(tick)
        svg.append(line(left, y, right, y, "#dfe5e2", 1.5))
        svg.append(text(left - 17, y + 8, tick, 25, "end", fill="#4d5a54"))
    for tick in (40, 50, 60, 70):
        x = xmap(tick)
        svg.append(line(x, top, x, bottom, "#edf0ee", 1.2))
        svg.append(text(x, bottom + 33, tick, 25, fill="#4d5a54"))

    for index, (_, points) in enumerate(TAS_FIELDS):
        coords = " ".join(
            f"{xmap(max(38, min(77, x))):.2f},{ymap(max(0, min(16, y))):.2f}"
            for x, y in points
        )
        fill = "#eef3f0" if index % 2 == 0 else "#f8f9f8"
        svg.append(f'<polygon points="{coords}" fill="{fill}" stroke="#87938d" stroke-width="2"/>')
    for lx, ly, lines, size in TAS_LABELS:
        svg.append(multiline_svg(xmap(lx), ymap(ly), lines, size))

    oxides = ["Na2O", "MgO", "Al2O3", "SiO2", "P2O5", "K2O", "CaO", "TiO2", "MnO", "Fe2O3"]
    factor = 100 / sum(major.get(name, 0) for name in oxides)
    sio2 = major["SiO2"] * factor
    alkalis = (major["Na2O"] + major["K2O"]) * factor
    px, py = xmap(sio2), ymap(alkalis)
    diamond = f"{px:.2f},{py-14:.2f} {px+14:.2f},{py:.2f} {px:.2f},{py+14:.2f} {px-14:.2f},{py:.2f}"
    svg.append(f'<polygon points="{diamond}" fill="#c74735" stroke="white" stroke-width="5"/>')
    svg.append(text(px - 20, py - 17, "SM25-07", 30, "end", 700, "#66291f"))
    svg.append(line(left, top, left, bottom, "#39443f", 2.5))
    svg.append(line(left, bottom, right, bottom, "#39443f", 2.5))
    svg.append(text((left + right) / 2, bottom + 79, "SiO₂, мас. %", 30))
    svg.append(text(40, (top + bottom) / 2, "Na₂O + K₂O, мас. %", 30, rotate=-90))

    # Primitive-mantle-normalized panel.
    mleft, mright, mtop, mbottom = 142, 1148, 1000, 1560
    mx = lambda i: mleft + i / (len(elements) - 1) * (mright - mleft)
    my = lambda value: mbottom - (math.log10(value) + 1) / 4 * (mbottom - mtop)
    svg.append(text(mleft, 955, "б) Мультиэлементная диаграмма · PM-normalized", 36, "start", 700))

    for value in (0.1, 1, 10, 100, 1000):
        y = my(value)
        svg.append(line(mleft, y, mright, y, "#dbe1de", 1.6))
        svg.append(text(mleft - 17, y + 8, f"{value:g}", 25, "end", fill="#4d5a54"))
    svg.append(line(mleft, mtop, mleft, mbottom, "#39443f", 2.5))
    svg.append(line(mleft, mbottom, mright, mbottom, "#39443f", 2.5))

    for i, element in enumerate(elements):
        svg.append(text(mx(i) + 7, mbottom + 30, element, 23, "end", fill="#39443f", rotate=-90))

    series = [
        ("N-MORB", refs["N-MORB (Gale et al., 2013)"], "#536a93", "18 11", 4, False),
        ("OIB", refs["OIB (Sun&McDoniugh, 1989)"], "#b06b2b", "18 11", 4, False),
        ("SM25-07", sample, "#126e64", None, 7, True),
    ]
    for label, values, color, dash, width, markers in series:
        coords = [(mx(i), my(values[element] / pm[element])) for i, element in enumerate(elements)]
        svg.append(path(coords, color, width, dash, markers))

    # Compact direct legend with high contrast.
    legend_x, legend_y = 875, 1025
    for i, (label, _, color, dash, width, _) in enumerate(series):
        y = legend_y + i * 38
        svg.append(line(legend_x, y, legend_x + 82, y, color, width, dash))
        svg.append(text(legend_x + 96, y + 8, label, 25, "start", 600))

    svg.append(text(40, (mtop + mbottom) / 2, "Порода / примитивная мантия", 30, rotate=-90))
    svg.append(text((mleft + mright) / 2, 1707, "Элементы", 30))
    svg.append(text(mleft, 1760, "PM: McDonough & Sun (1989); N-MORB: Gale et al. (2013); OIB: Sun & McDonough (1989)", 22, "start", fill="#53605a"))

    svg.extend(["</g>", "</svg>"])
    OUT_SVG.parent.mkdir(parents=True, exist_ok=True)
    OUT_SVG.write_text("\n".join(svg), encoding="utf-8")
    render_png(major, sample, elements, pm, refs)
    render_pdf()
    print(OUT_SVG)
    print(OUT_PNG)
    print(OUT_PDF)


if __name__ == "__main__":
    main()
