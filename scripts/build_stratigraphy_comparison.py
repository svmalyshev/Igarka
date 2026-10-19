from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets"
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.linewidth": 0.8,
    "svg.fonttype": "none",
})

COLORS = {
    "gub": "#d99072",
    "ludy": "#9aa17b",
    "medv": "#e6d8a8",
    "chern": "#a9afb1",
    "izl": "#c46e58",
    "sukh": "#d7bd72",
    "volc": "#548a65",
    "minor": "#c9b8b2",
    "unknown": "#e5e3de",
    "ink": "#242723",
    "muted": "#68706a",
    "paper": "#f7f4ed",
    "accent": "#9a4b37",
}


def rect(ax, x, y0, y1, label, color, hatch=None, width=2.15,
         fontsize=10, label_color=None, zorder=2):
    patch = patches.Rectangle(
        (x, y0), width, y1 - y0,
        facecolor=color, edgecolor=COLORS["ink"], linewidth=0.9,
        hatch=hatch, zorder=zorder,
    )
    ax.add_patch(patch)
    ax.text(
        x + width / 2, (y0 + y1) / 2, label,
        ha="center", va="center", fontsize=fontsize,
        color=label_color or COLORS["ink"], linespacing=1.08,
        zorder=zorder + 1,
    )
    return patch


def wavy(ax, x, y, width=2.15, color=None, amp=0.055, waves=7, lw=2.0):
    xs = np.linspace(x, x + width, 160)
    ys = y + amp * np.sin(np.linspace(0, waves * 2 * np.pi, xs.size))
    ax.plot(xs, ys, color=color or COLORS["accent"], linewidth=lw, zorder=8)


def side_wedge(ax, x, y0, y1, text, color, side="right", width=0.52):
    if side == "right":
        pts = [(x, y0), (x + width, y0 + 0.25), (x + width, y1 - 0.25), (x, y1)]
        tx = x + width * 0.58
    else:
        pts = [(x + width, y0), (x, y0 + 0.25), (x, y1 - 0.25), (x + width, y1)]
        tx = x + width * 0.42
    ax.add_patch(patches.Polygon(pts, closed=True, facecolor=color,
                                 edgecolor=COLORS["ink"], linewidth=0.8, zorder=5))
    ax.text(tx, (y0 + y1) / 2, text, ha="center", va="center",
            rotation=90, fontsize=8.3, color="white", zorder=6)


def sample(ax, x, y, text, color="#7f2439", dx=0.22, ha="left"):
    ax.scatter([x], [y], marker="*", s=155, facecolor=color,
               edgecolor="white", linewidth=0.8, zorder=12)
    ax.text(x + dx, y, text, va="center", ha=ha, fontsize=9,
            color=COLORS["ink"], zorder=12)


fig, ax = plt.subplots(figsize=(16, 10), dpi=190)
fig.patch.set_facecolor(COLORS["paper"])
ax.set_facecolor(COLORS["paper"])
ax.set_xlim(-0.6, 12.8)
ax.set_ylim(-1.62, 9.6)
ax.axis("off")

ax.text(-0.35, 9.35, "Игарский район: как менялась схема докембрийского разреза",
        fontsize=21, fontweight="bold", color=COLORS["ink"], ha="left")
ax.text(-0.35, 9.0,
        "Сопоставление стратиграфических отношений; высота блоков условная и не отражает мощности",
        fontsize=11.5, color=COLORS["muted"], ha="left")

columns = [
    (0.2, "НКГРЭ, 1994", "рис. 3.2, «принятая схема»"),
    (4.45, "Госгеолкарта 1:200 000", "Q-45-IX, X (Полой), 2001"),
    (8.7, "Kochnev et al., 2022/2025", "сводная колонка с новыми данными"),
]
for x, title, sub in columns:
    ax.text(x + 1.075, 8.65, title, ha="center", va="bottom",
            fontsize=13.5, fontweight="bold", color=COLORS["ink"])
    ax.text(x + 1.075, 8.40, sub, ha="center", va="bottom",
            fontsize=9.5, color=COLORS["muted"])

# НКГРЭ, 1994: accepted Igarka-area correlation scheme.
x = 0.2
rect(ax, x, 0.00, 0.75, "губинская\nтолща", COLORS["gub"], hatch="..")
rect(ax, x, 0.75, 1.18, "малиново-сланцевая\nтолща", COLORS["minor"], hatch="---", fontsize=8.6)
rect(ax, x, 1.18, 1.56, "свита Ржевского", "#c6c9a9", hatch="..", fontsize=8.7)
rect(ax, x, 1.56, 2.34, "медвежьинская\nсвита", COLORS["medv"], hatch="-")
rect(ax, x, 2.34, 4.12, "лудовская толща", COLORS["ludy"], hatch="---")
side_wedge(ax, x + 1.63, 2.52, 3.98, "игарская пачка", COLORS["volc"], width=0.52)
wavy(ax, x, 4.18)
rect(ax, x, 4.22, 5.55, "чернореченская\nсвита", COLORS["chern"], hatch="-")
rect(ax, x, 5.55, 6.88, "излучинская свита", COLORS["izl"], hatch="..")
rect(ax, x, 6.88, 7.30, "гравийская +\nлайдашинская", "#d9a477", hatch="..", fontsize=8.4)
rect(ax, x, 7.30, 8.05, "сухарихинская\nсвита", COLORS["sukh"], hatch="-")
ax.text(x + 2.38, 4.18, "региональный\nперерыв", va="center", ha="left",
        fontsize=8.6, color=COLORS["accent"])

# 1:200 000 explanatory note: units actually distinguished on the sheet.
x = 4.45
rect(ax, x, 0.00, 0.62, "ниже: не расчленено\nна данном листе",
     COLORS["unknown"], hatch="//", fontsize=7.4, label_color=COLORS["muted"])
rect(ax, x, 0.62, 4.12, "губинская свита  R₂gb\nкварцевые и полимиктовые песчаники,\nгравелиты, конгломераты; до 1120 м",
     COLORS["gub"], hatch="..", fontsize=9.2)
wavy(ax, x, 4.18)
rect(ax, x, 4.22, 5.58, "чернореченская свита  R₃čr\nкарбонатно-глинистая; около 900 м",
     COLORS["chern"], hatch="-", fontsize=9.0)
rect(ax, x, 5.58, 7.15, "излучинская свита  R₃–V iz\nкрасноцветные терригенные породы;\nоколо 1100 м",
     COLORS["izl"], hatch="..", fontsize=9.0)
rect(ax, x, 7.15, 8.05, "сухарихинская свита\n(продолжение разреза)", COLORS["sukh"], hatch="-", fontsize=9.0)
ax.text(x + 2.38, 4.18, "несогласие", va="center", ha="left",
        fontsize=8.6, color=COLORS["accent"])
ax.text(x + 1.075, 0.28, "лудовская и медвежьинская\nздесь отдельно не показаны",
        ha="center", va="center", fontsize=8.1, color=COLORS["muted"])

# Kochnev composite column.
x = 8.70
rect(ax, x, 0.00, 0.28, "безымянная (?)", "#dad6cd", hatch="..", fontsize=8.2)
rect(ax, x, 0.28, 0.72, "медвежьинская свита", COLORS["medv"], hatch="-", fontsize=8.8)
rect(ax, x, 0.72, 2.42, "лудовская свита\nалевролиты, редкие песчаники\nи известняки", COLORS["ludy"], hatch="---", fontsize=9.0)
side_wedge(ax, x, 0.88, 2.30, "игарский комплекс", COLORS["volc"], side="left", width=0.50)
rect(ax, x, 2.42, 4.12, "губинская свита\nпесчаники → глинистые сланцы", COLORS["gub"], hatch="..", fontsize=8.8)
ax.add_patch(patches.Rectangle((x + 0.23, 2.62), 1.70, 0.12,
                               facecolor="#286044", edgecolor=COLORS["ink"],
                               linewidth=0.8, zorder=9))
wavy(ax, x, 4.18)
rect(ax, x, 4.22, 5.58, "чернореченская свита", COLORS["chern"], hatch="-")
rect(ax, x, 5.58, 7.15, "излучинская свита", COLORS["izl"], hatch="..")
rect(ax, x, 7.15, 8.05, "сухарихинская свита", COLORS["sukh"], hatch="-")
sample(ax, x + 2.03, 1.10, "2007-15")
sample(ax, x + 2.03, 3.75, "1812-2")
sample(ax, x + 0.18, 2.68, "1812-1 (силл)", color="#234d3a", dx=-0.18, ha="right")
ax.text(x + 2.38, 4.18, "региональный\nперерыв", va="center", ha="left",
        fontsize=8.6, color=COLORS["accent"])

# Direction and main interpretation.
ax.annotate("моложе", xy=(-0.20, 8.05), xytext=(-0.20, 0.05),
            arrowprops=dict(arrowstyle="-|>", color=COLORS["muted"], lw=1.2),
            ha="center", va="bottom", rotation=90, fontsize=9, color=COLORS["muted"])

box = patches.FancyBboxPatch((0.2, -0.86), 11.95, 0.58,
                             boxstyle="round,pad=0.02,rounding_size=0.05",
                             facecolor="#f0e2d6", edgecolor="#c08367", linewidth=1.0)
ax.add_patch(box)
ax.text(0.42, -0.57,
        "Главное расхождение: в схеме 1994 г. губинская лежит ниже медвежьинской и лудовской; "
        "у Кочнева порядок обращён — медвежьинская → лудовская → губинская.",
        ha="left", va="center", fontsize=10.2, color=COLORS["ink"])

ax.text(0.2, -1.16, "Основание построения Кочнева:", fontsize=9.3,
        fontweight="bold", color=COLORS["ink"], ha="left")
ax.text(2.35, -1.16,
        "локальная схема карт 1:200 000 (Гройсман, Куликов, 1984; Гайнцева, Невьянцев, 2001), "
        "модифицированная по Козлову и др. (1992а,б); корреляция Мельникова и др. (2005); "
        "наблюдения и датировки Kochnev et al. (2022).\nРис. 2 в статье 2025 г. — схема 2022 г. с дополнениями.",
        fontsize=8.6, color=COLORS["muted"], ha="left", va="center", wrap=True)

ax.text(0.2, -1.47,
        "Источники: Савушкин и др., отчёт Игарской партии НКГРЭ, 1994, кн. 1, с. 184, рис. 3.2; "
        "Госгеолкарта РФ 1:200 000, Q-45-IX, X (Полой), 2001; Kochnev et al., Precambrian Research 382 (2022) 106877; "
        "\nKochnev et al., Geodynamics & Tectonophysics 16 (2025).",
        fontsize=8.1, color=COLORS["muted"], ha="left", va="center")

fig.subplots_adjust(left=0.02, right=0.99, top=0.98, bottom=0.02)
png = OUT / "igarka_stratigraphy_comparison.png"
svg = OUT / "igarka_stratigraphy_comparison.svg"
fig.savefig(png, dpi=190, facecolor=fig.get_facecolor())
fig.savefig(svg, facecolor=fig.get_facecolor())
print(png)
print(svg)
