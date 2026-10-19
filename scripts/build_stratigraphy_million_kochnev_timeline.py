from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageOps
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT.parent
OUT = ROOT / "assets"
OUT.mkdir(parents=True, exist_ok=True)

W, H = 1150, 900
rng = np.random.default_rng(2018)
FONT = FontProperties(fname="/System/Library/Fonts/Supplemental/Arial Narrow.ttf")
FONT_BOLD = FontProperties(fname="/System/Library/Fonts/Supplemental/Arial Narrow Bold.ttf")


def make_paper():
    reference = WORKSPACE / "Презентация" / "assets" / "fig3_1_riphean_v11_clean_base.png"
    source = Image.open(reference).convert("RGB").crop((0, 0, 900, 150))
    paper = Image.new("RGB", (W, H), (250, 250, 248))
    tile_w, tile_h = 450, 150
    tile = source.resize((tile_w, tile_h), Image.BICUBIC)
    for y in range(0, H, tile_h):
        for x in range(0, W, tile_w):
            piece = tile
            if ((x // tile_w) + (y // tile_h)) % 2:
                piece = ImageOps.mirror(piece)
            paper.paste(piece, (x, y))
    wash = Image.new("RGBA", (W, H), (255, 255, 255, 45))
    return Image.alpha_composite(paper.convert("RGBA"), wash)


def watercolor(image, polygon, rgb, opacity=0.42):
    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).polygon([(int(x), int(H - y)) for x, y in polygon], fill=255)
    noise = rng.random((45, 90))
    cloud = Image.fromarray(np.uint8(noise * 255)).resize((W, H), Image.BICUBIC)
    cloud = cloud.filter(ImageFilter.GaussianBlur(12))
    arr = np.asarray(cloud, dtype=float)
    arr = (arr - arr.min()) / max(1, arr.max() - arr.min())
    alpha = Image.fromarray(np.uint8((32 + 95 * arr) * opacity / 0.42), "L")
    alpha = Image.composite(alpha, Image.new("L", (W, H), 0), mask)
    layer = Image.new("RGBA", (W, H), (*rgb, 0))
    layer.putalpha(alpha)
    image.alpha_composite(layer)


def rough_horizontal(x0, x1, y, seed, waviness=1.7, waves=None):
    local = np.random.default_rng(seed)
    xs = np.linspace(x0, x1, 140)
    if waves:
        ys = y + waviness * np.sin(np.linspace(0, waves * np.pi, xs.size))
        ys += local.normal(0, 0.35, xs.size)
    else:
        raw = local.normal(0, 1, xs.size)
        ys = y + np.convolve(raw, np.ones(11) / 11, mode="same") * waviness
    return xs, ys


def rough_vertical(x, y0, y1, seed):
    local = np.random.default_rng(seed)
    ys = np.linspace(y0, y1, 140)
    raw = local.normal(0, 1, ys.size)
    xs = x + np.convolve(raw, np.ones(11) / 11, mode="same") * 1.6
    return xs, ys


def stroke(ax, xs, ys, width=1.35, alpha=0.96):
    ax.plot(xs, ys, color="#121313", lw=width, alpha=alpha,
            solid_joinstyle="round", solid_capstyle="round", zorder=20)


def hline(ax, x0, x1, y, seed, wavy=False):
    xs, ys = rough_horizontal(x0, x1, y, seed, waviness=4.0 if wavy else 1.7,
                              waves=12 if wavy else None)
    stroke(ax, xs, ys, width=1.5 if wavy else 1.25)


def vline(ax, x, y0, y1, seed):
    xs, ys = rough_vertical(x, y0, y1, seed)
    stroke(ax, xs, ys, width=1.25)


def break_hatching(ax, x0, x1, y0, y1, seed):
    """Hand-drawn vertical hatching used for a stratigraphic hiatus."""
    local = np.random.default_rng(seed)
    for x in np.arange(x0 + 9, x1 - 5, 14):
        dx = local.normal(0, 0.7)
        ax.plot([x + dx, x - dx], [y0 + 3, y1 - 3], color="#343535",
                lw=0.72, alpha=0.58, solid_capstyle="round", zorder=16)


def text(ax, x, y, value, size=25, rotation=0, color="#171717", ha="center"):
    ax.text(x, y, value, ha=ha, va="center", color=color, rotation=rotation,
            fontsize=size, fontproperties=FONT, linespacing=1.05, zorder=30)


paper = make_paper()

# Three narrow, directly comparable columns on the same timeline.
A = (70, 370)        # 1994, Fig. 3.2, column "Igarka district"
B = (425, 725)       # State Geological Map 1:1,000,000, 2018
C = (780, 1080)      # Kochnev et al., 2022/2025
bottom, top = 235, 755

# 1994: Gubinskaya at the base; the Igarka member intertongues laterally
# with the Ludy unit exactly as in the "Igarka district" column of Fig. 3.2.
zigzag_1994 = [
    (A[1] - 7, 455), (A[0] + 244, 476), (A[0] + 262, 497),
    (A[0] + 232, 519), (A[0] + 252, 542), (A[0] + 222, 565),
    (A[0] + 240, 587), (A[0] + 210, 610),
]
watercolor(paper, [(A[0], 235), (A[1], 235), (A[1], 292), (A[0], 292)],
           (246, 135, 98), 0.43)
watercolor(paper, zigzag_1994 + [(A[1], 610), (A[1], 455)],
           (125, 236, 80), 0.37)

# Million-scale map, 2018: Igarka Formation laterally replaces Ludy + Gubinskaya.
watercolor(paper, [(B[0], 355), (B[0] + 105, 355), (B[0] + 105, 610), (B[0], 610)],
           (125, 236, 80), 0.37)
watercolor(paper, [(B[0] + 105, 505), (B[1], 505), (B[1], 610), (B[0] + 105, 610)],
           (246, 135, 98), 0.45)

# Kochnev, 2025, Fig. 2: the Igarka complex has a jagged lateral boundary
# with both the Ludy and Gubinskaya formations.
zigzag = [
    (C[0] + 82, 360), (C[0] + 122, 380), (C[0] + 78, 400),
    (C[0] + 126, 423), (C[0] + 80, 446), (C[0] + 124, 470),
    (C[0] + 78, 494), (C[0] + 128, 518), (C[0] + 82, 542),
    (C[0] + 126, 566), (C[0] + 80, 588), (C[0] + 112, 610),
]
watercolor(paper, [(C[0], 360)] + zigzag + [(C[0], 610)],
           (125, 236, 80), 0.37)
watercolor(paper, zigzag[7:] + [(C[1], 610), (C[1], 505), (C[0] + 105, 505)],
           (246, 135, 98), 0.45)

fig = plt.figure(figsize=(11.5, 9), dpi=100)
ax = fig.add_axes([0, 0, 1, 1])
ax.imshow(paper, extent=(0, W, 0, H), zorder=0)
ax.set_xlim(0, W)
ax.set_ylim(0, H)
ax.axis("off")

# Shared timeline baseline.
hline(ax, 30, 1120, 205, 1)

# 1994 report, Fig. 3.2 — reproduce precisely the "Igarka district" column.
vline(ax, A[0], bottom, top, 10)
vline(ax, A[1], bottom, top, 11)
hline(ax, A[0], A[1], top, 12)
hline(ax, A[0], A[1], 292, 13)              # Gubinskaya / Malinovo-slate
hline(ax, A[0], A[1], 337, 14)              # Malinovo-slate / Rzhevsky
hline(ax, A[0], A[1], 382, 15, wavy=True)   # Rzhevsky / Medvezhy
hline(ax, A[0], A[1], 455, 16)              # Medvezhy / Ludy
hline(ax, A[0], A[1], 610, 17, wavy=True)   # lower edge of hiatus
hline(ax, A[0], A[1], 650, 18, wavy=True)   # base of Chernorechenskaya
break_hatching(ax, A[0], A[1], 610, 650, 19)
stroke(ax, np.array([p[0] for p in zigzag_1994]),
       np.array([p[1] for p in zigzag_1994]), width=1.4)

text(ax, 220, 263, "Губинская\nтолща", 18)
text(ax, 220, 315, "малиново-сланцевая\nтолща", 14)
text(ax, 220, 359, "свита Ржевского", 16)
text(ax, 220, 418, "Медвежинская\nсвита", 17)
text(ax, 175, 530, "Лудовская\nтолща", 19)
text(ax, 345, 535, "Игарская пачка", 15, rotation=90)
text(ax, 220, 704, "Чернореченская\nсвита", 19)

# State Geological Map, 2018.
vline(ax, B[0], bottom, top, 30)
vline(ax, B[1], bottom, top, 31)
hline(ax, B[0], B[1], top, 32)
hline(ax, B[0], B[1], 355, 33, wavy=True)
hline(ax, B[0] + 105, B[1], 505, 34)
hline(ax, B[0], B[1], 610, 35)
hline(ax, B[0], B[1], 650, 36, wavy=True)
vline(ax, B[0] + 105, 355, 610, 37)

text(ax, 575, 295, "Корабликовская\nсвита", 18)
text(ax, 627, 430, "Лудовская\nсвита", 19)
text(ax, 627, 557, "Губинская\nтолща", 19)
text(ax, 477, 483, "Игарская свита", 16, rotation=90)
text(ax, 575, 630, "Медвежинская свита", 14)
text(ax, 575, 704, "Чернореченская\nсвита", 19)

# Kochnev 2025, Fig. 2.
vline(ax, C[0], bottom, top, 50)
vline(ax, C[1], bottom, top, 51)
hline(ax, C[0], C[1], top, 52)
hline(ax, C[0], C[1], 285, 53)
hline(ax, C[0], C[1], 360, 54)
hline(ax, C[0] + 105, C[1], 505, 55)
hline(ax, C[0], C[1], 610, 56, wavy=True)
hline(ax, C[0], C[1], 650, 57, wavy=True)
break_hatching(ax, C[0], C[1], 610, 650, 58)
stroke(ax, np.array([p[0] for p in zigzag]), np.array([p[1] for p in zigzag]), width=1.45)

text(ax, 930, 260, "Безымянная\nсвита", 16)
text(ax, 930, 323, "Медвежинская\nсвита", 16)
text(ax, 985, 430, "Лудовская\nсвита", 19)
text(ax, 985, 557, "Губинская\nсвита", 19)
text(ax, 821, 485, "Игарский комплекс", 15, rotation=90)
text(ax, 930, 704, "Чернореченская\nсвита", 19)

# Dots and captions repeat the established timeline grammar.
centers = [220, 575, 930]
ax.scatter(centers, [205, 205, 205], s=46, color="#251820", zorder=40)
text(ax, 220, 150, "Отчёт НКГРЭ, 1994", 21)
text(ax, 575, 150, "Госгеолкарта, 2018", 21)
text(ax, 930, 150, "Кочнев и др., 2022, 2025", 20)

png = OUT / "igarka_stratigraphy_1994_million_kochnev_timeline.png"
fig.savefig(png, dpi=150, facecolor="white")
print(png)
