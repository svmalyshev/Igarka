from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets"
OUT.mkdir(parents=True, exist_ok=True)

W, H = 2600, 1500
rng = np.random.default_rng(45)
FONT = FontProperties(fname="/System/Library/Fonts/Noteworthy.ttc")
FONT_BOLD = FontProperties(fname="/System/Library/Fonts/Noteworthy.ttc", weight="bold")


def paper_texture():
    fine = rng.normal(0, 1.8, (H, W))
    coarse = rng.normal(0, 1, (max(2, H // 18), max(2, W // 18)))
    coarse = Image.fromarray(np.uint8(np.clip(128 + coarse * 28, 0, 255))).resize((W, H), Image.BICUBIC)
    coarse = (np.asarray(coarse, dtype=float) - 128) * 0.16
    base = np.zeros((H, W, 3), dtype=float)
    base[:] = np.array([249, 249, 247])
    base += fine[..., None] + coarse[..., None]
    image = Image.fromarray(np.uint8(np.clip(base, 0, 255)), "RGB")
    fibers = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    draw = ImageDraw.Draw(fibers)
    for _ in range(170):
        x = int(rng.integers(-100, W))
        y = int(rng.integers(0, H))
        length = int(rng.integers(20, 170))
        color = (120, 115, 105, int(rng.integers(3, 9)))
        draw.arc((x, y, x + length, y + int(rng.integers(4, 18))), 180, 360, fill=color, width=1)
    return Image.alpha_composite(image.convert("RGBA"), fibers)


def watercolor(image, polygon, color, strength=0.48):
    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).polygon([(int(x), int(H - y)) for x, y in polygon], fill=255)
    small = rng.random((max(2, H // 22), max(2, W // 22)))
    cloud = Image.fromarray(np.uint8(small * 255)).resize((W, H), Image.BICUBIC)
    cloud = cloud.filter(ImageFilter.GaussianBlur(18))
    alpha = np.asarray(cloud, dtype=float)
    alpha = (alpha - alpha.min()) / max(1, alpha.max() - alpha.min())
    alpha = np.uint8((38 + 105 * alpha) * strength / 0.48)
    alpha = Image.fromarray(alpha, "L")
    alpha = Image.composite(alpha, Image.new("L", (W, H), 0), mask)
    wash = Image.new("RGBA", (W, H), (*color, 0))
    wash.putalpha(alpha)
    image.alpha_composite(wash)


def noisy_path(x0, x1, y, amp=5.0, n=100, seed=None):
    local = np.random.default_rng(seed)
    xs = np.linspace(x0, x1, n)
    noise = local.normal(0, 1, n)
    smooth = np.convolve(noise, np.ones(9) / 9, mode="same")
    ys = y + amp * smooth
    return xs, ys


def hand_line(ax, xs, ys, lw=1.65, seed=0):
    local = np.random.default_rng(seed)
    xs = np.asarray(xs, dtype=float)
    ys = np.asarray(ys, dtype=float)
    for i, (alpha, width) in enumerate(((0.78, lw + 0.75), (0.98, lw))):
        jitter_x = local.normal(0, 0.55 if len(xs) > 2 else 0.35, len(xs))
        jitter_y = local.normal(0, 0.55 if len(ys) > 2 else 0.35, len(ys))
        ax.plot(xs + jitter_x, ys + jitter_y, color="#141716", lw=width,
                alpha=alpha, solid_capstyle="round", solid_joinstyle="round", zorder=10 + i)


def boundary(ax, x0, x1, y, wavy=False, seed=0, amp=6):
    if wavy:
        xs = np.linspace(x0, x1, 150)
        ys = y + amp * np.sin(np.linspace(0, 12 * np.pi, 150))
        ys += np.random.default_rng(seed).normal(0, 0.7, 150)
    else:
        xs, ys = noisy_path(x0, x1, y, amp=2.8, seed=seed)
    hand_line(ax, xs, ys, seed=seed)


def vertical(ax, x, y0, y1, seed=0):
    ys = np.linspace(y0, y1, 100)
    local = np.random.default_rng(seed)
    xs = x + np.convolve(local.normal(0, 1, 100), np.ones(9) / 9, mode="same") * 2.3
    hand_line(ax, xs, ys, seed=seed)


def label(ax, x, y, text, size=25, rotation=0, color="#181b19", weight=False, ha="center"):
    ax.text(x, y, text, ha=ha, va="center", fontsize=size,
            fontproperties=FONT_BOLD if weight else FONT, color=color,
            rotation=rotation, linespacing=1.05, zorder=30)


bg = paper_texture()

# Geometry of the two reconstructed columns.
left = (180, 1050)
right = (1450, 2320)
base_y, top_y = 260, 1300

# Watercolour washes: million-scale scheme.
lx0, lx1 = left
watercolor(bg, [(lx0, 555), (lx1, 555), (lx1, 735), (lx0, 735)], (236, 121, 86), 0.52)
watercolor(bg, [(lx0, 555), (lx0 + 255, 555), (lx0 + 255, 890), (lx0, 890)], (122, 226, 80), 0.42)

# Kochnev scheme.
rx0, rx1 = right
watercolor(bg, [(rx0, 575), (rx1, 575), (rx1, 765), (rx0, 765)], (236, 121, 86), 0.52)
watercolor(bg, [(rx0, 420), (rx0 + 250, 420), (rx0 + 250, 765), (rx0, 765)], (122, 226, 80), 0.42)

# A dark green dolerite body in the Gubinskaya Formation.
draw = ImageDraw.Draw(bg)
draw.rectangle((rx0 + 360, H - 680, rx1 - 70, H - 662),
               fill=(55, 116, 70, 150), outline=(25, 64, 36, 190), width=2)

fig = plt.figure(figsize=(13, 7.5), dpi=200)
ax = fig.add_axes([0, 0, 1, 1])
ax.imshow(bg, extent=(0, W, 0, H), zorder=0)
ax.set_xlim(0, W)
ax.set_ylim(0, H)
ax.axis("off")

# Baseline and column outer contours.
boundary(ax, 60, W - 60, base_y - 24, seed=900)
for x0, x1, seed in ((lx0, lx1, 11), (rx0, rx1, 21)):
    vertical(ax, x0, base_y, top_y, seed)
    vertical(ax, x1, base_y, top_y, seed + 1)
    boundary(ax, x0, x1, top_y, seed=seed + 2)

# 1:1,000,000 (2018) - lower to upper.
for i, (y, is_wavy) in enumerate([
    (400, True),     # top of Korablikovskaya / base Ludy
    (555, False),    # Ludy -> Gubinskaya
    (735, False),    # top Gubinskaya
    (890, True),     # base Medvezhinskaya
    (1040, True),    # regional unconformity beneath Chernorechenskaya
]):
    boundary(ax, lx0, lx1, y, wavy=is_wavy, seed=100 + i)
vertical(ax, lx0 + 255, 555, 890, 150)

label(ax, (lx0 + lx1) / 2, 330, "Корабликовская свита", 25)
label(ax, lx0 + 555, 475, "Лудовская свита", 27)
label(ax, lx0 + 555, 645, "Губинская толща", 28)
label(ax, lx0 + 128, 723, "Игарская\nсвита", 25)
label(ax, (lx0 + lx1) / 2, 812, "Медвежинская свита", 27)
label(ax, (lx0 + lx1) / 2, 960, "региональный перерыв", 19, color="#5f6460")
label(ax, (lx0 + lx1) / 2, 1165, "Чернореченская свита", 29)

# Kochnev et al. - lower to upper.
for i, (y, is_wavy) in enumerate([
    (325, False),    # Bezymyannaya -> Medvezhinskaya
    (420, True),     # Medvezhinskaya -> Ludy
    (575, False),    # Ludy -> Gubinskaya
    (765, False),    # top Gubinskaya
    (835, True),     # regional unconformity
]):
    boundary(ax, rx0, rx1, y, wavy=is_wavy, seed=200 + i)
vertical(ax, rx0 + 250, 420, 765, 250)

label(ax, (rx0 + rx1) / 2, 292, "Безымянная (?) свита", 21)
label(ax, (rx0 + rx1) / 2, 372, "Медвежинская свита", 24)
label(ax, rx0 + 565, 492, "Лудовская свита", 27)
label(ax, rx0 + 565, 666, "Губинская свита", 28)
label(ax, rx0 + 125, 590, "Игарский\nкомплекс", 24)
label(ax, rx0 + 600, 683, "1812-1", 17, color="#315f3a")
label(ax, rx1 - 58, 724, "1812-2", 18, ha="right")
label(ax, rx1 - 58, 515, "2007-15", 18, ha="right")
label(ax, (rx0 + rx1) / 2, 801, "региональный перерыв", 19, color="#5f6460")
label(ax, (rx0 + rx1) / 2, 1065, "Чернореченская свита", 29)
label(ax, (rx0 + rx1) / 2, 865, "корабликовская отдельно\nне выделяется", 18, color="#5f6460")

# Handwritten-style captions, like the supplied reference.
ax.scatter([(lx0 + lx1) / 2, (rx0 + rx1) / 2], [base_y - 24, base_y - 24],
           s=75, c="#24191b", zorder=31)
label(ax, (lx0 + lx1) / 2, 162, "2018 — Госгеолкарта 1 : 1 000 000", 25)
label(ax, (lx0 + lx1) / 2, 112, "Варганов, Сосновская, Назаров, Чеканов и др.", 20)
label(ax, (rx0 + rx1) / 2, 162, "2022, 2025 — B. B. Kochnev et al.", 25)
label(ax, (rx0 + rx1) / 2, 112, "сводная колонка с новыми геохронологическими данными", 19)

ax.text(1300, 1415, "Игарский район: две схемы нижней части докембрийского разреза",
        ha="center", va="center", fontsize=31, fontproperties=FONT_BOLD,
        color="#171918", zorder=30)

# Compact, non-intrusive note about the key rearrangement.
ax.annotate("медвежинская перенесена вниз",
            xy=(rx0 + 435, 372), xytext=(1260, 338),
            fontsize=18, fontproperties=FONT, color="#424743",
            ha="center", va="center",
            arrowprops=dict(arrowstyle="->", lw=1.4, color="#424743",
                            connectionstyle="arc3,rad=-0.13"), zorder=35)

png = OUT / "igarka_stratigraphy_million_kochnev.png"
svg = OUT / "igarka_stratigraphy_million_kochnev.svg"
fig.savefig(png, dpi=200, facecolor="white")
fig.savefig(svg, facecolor="white")
print(png)
print(svg)
