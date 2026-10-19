from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets"
OUT.mkdir(parents=True, exist_ok=True)

# 16 × 5 cm at 600 dpi: suitable for enlarging on a presentation slide.
W, H = 3780, 1181
SEED = 20261011
rng = np.random.default_rng(SEED)

CORAL = np.array([246, 194, 180], dtype=np.float32)
INK = (20, 23, 22, 245)
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Narrow.ttf"


def smooth_noise(width: int, height: int, scale: int, blur: float) -> np.ndarray:
    sw = max(2, width // scale)
    sh = max(2, height // scale)
    noise = rng.random((sh, sw), dtype=np.float32)
    image = Image.fromarray(np.uint8(noise * 255), "L")
    image = image.resize((width, height), Image.BICUBIC)
    image = image.filter(ImageFilter.GaussianBlur(blur))
    arr = np.asarray(image, dtype=np.float32) / 255.0
    return (arr - arr.min()) / max(1e-6, arr.max() - arr.min())


def irregular_edge(start, end, normal_axis: str, amplitude: float, seed: int,
                   broad_waves: float = 0.0):
    local = np.random.default_rng(seed)
    n = 260
    t = np.linspace(0.0, 1.0, n)
    raw = local.normal(0.0, 1.0, n)
    kernel = np.ones(21) / 21
    offset = np.convolve(raw, kernel, mode="same") * amplitude
    if broad_waves:
        # Broad, uneven undulations like the upper contact in the source scheme.
        offset += broad_waves * (
            0.68 * np.sin(t * 10.5 * np.pi + 0.35)
            + 0.22 * np.sin(t * 21.0 * np.pi + 1.1)
        )
    x0, y0 = start
    x1, y1 = end
    xs = x0 + (x1 - x0) * t
    ys = y0 + (y1 - y0) * t
    if normal_axis == "y":
        ys += offset
    else:
        xs += offset
    return list(zip(xs.round().astype(int), ys.round().astype(int)))


def make_polygon(left: int, top: int, right: int, bottom: int):
    top_edge = irregular_edge((left, top), (right, top), "y", 18, 11, broad_waves=33)
    right_edge = irregular_edge((right, top), (right, bottom), "x", 13, 12)
    bottom_edge = irregular_edge((right, bottom), (left, bottom), "y", 15, 13,
                                 broad_waves=4)
    left_edge = irregular_edge((left, bottom), (left, top), "x", 13, 14)
    return top_edge + right_edge + bottom_edge + left_edge


def fit_font(text: str, max_width: int, start_size: int):
    size = start_size
    while size > 24:
        font = ImageFont.truetype(FONT_PATH, size=size)
        box = font.getbbox(text)
        if box[2] - box[0] <= max_width:
            return font
        size -= 4
    return ImageFont.truetype(FONT_PATH, size=24)


canvas = Image.new("RGBA", (W, H), (255, 255, 255, 0))
left, top, right, bottom = 85, 90, W - 85, H - 90
polygon = make_polygon(left, top, right, bottom)

mask = Image.new("L", (W, H), 0)
ImageDraw.Draw(mask).polygon(polygon, fill=255)

# Several translucent, differently scaled washes reproduce the pale coral
# watercolour used for the Gubinskaya Formation in the correlation schemes.
fine = smooth_noise(W, H, scale=18, blur=10)
coarse = smooth_noise(W, H, scale=48, blur=36)
granule = rng.normal(0.0, 1.0, (H, W)).astype(np.float32)
density = np.clip(0.67 + 0.19 * coarse + 0.08 * fine + 0.025 * granule, 0.52, 0.94)

rgb = np.empty((H, W, 4), dtype=np.uint8)
paper_mix = np.array([255, 252, 249], dtype=np.float32)
mix = (0.62 + 0.16 * coarse)[..., None]
rgb[..., :3] = np.clip(CORAL * mix + paper_mix * (1.0 - mix), 0, 255).astype(np.uint8)
rgb[..., 3] = np.uint8(232 * density)
wash = Image.fromarray(rgb, "RGBA")
wash.putalpha(Image.composite(wash.getchannel("A"), Image.new("L", (W, H), 0), mask))
canvas = Image.alpha_composite(canvas, wash)

# Faint pooled pigment and pale blooms make the fill feel painted, not flat.
bloom = Image.new("RGBA", (W, H), (255, 255, 255, 0))
bloom_draw = ImageDraw.Draw(bloom, "RGBA")
for _ in range(42):
    cx = int(rng.integers(left + 80, right - 80))
    cy = int(rng.integers(top + 55, bottom - 55))
    rx = int(rng.integers(80, 310))
    ry = int(rng.integers(35, 150))
    if rng.random() < 0.58:
        colour = (255, 247, 240, int(rng.integers(7, 18)))
    else:
        colour = (184, 68, 43, int(rng.integers(4, 12)))
    bloom_draw.ellipse((cx - rx, cy - ry, cx + rx, cy + ry), fill=colour)
bloom = bloom.filter(ImageFilter.GaussianBlur(55))
bloom.putalpha(Image.composite(bloom.getchannel("A"), Image.new("L", (W, H), 0), mask))
canvas = Image.alpha_composite(canvas, bloom)

# Slightly doubled, irregular outline as in the hand-drawn columns.
outline = Image.new("RGBA", (W, H), (255, 255, 255, 0))
od = ImageDraw.Draw(outline, "RGBA")
od.line(polygon + [polygon[0]], fill=(20, 23, 22, 140), width=12, joint="curve")
inner = [(x + int(rng.normal(0, 1.5)), y + int(rng.normal(0, 1.5))) for x, y in polygon]
od.line(inner + [inner[0]], fill=INK, width=6, joint="curve")
canvas = Image.alpha_composite(canvas, outline)

text = "Губинская\nсвита"
font = fit_font("Губинская", right - left - 600, 265)
draw = ImageDraw.Draw(canvas)
bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=-12,
                               align="center", stroke_width=0)
tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
tx = (W - tw) / 2 - bbox[0]
ty = (H - th) / 2 - bbox[1] - 8
draw.multiline_text((tx, ty), text, font=font, fill=(24, 27, 25, 255),
                    spacing=-12, align="center")

png = OUT / "gubinskaya_formation_block_16cm_600dpi.png"
canvas.save(png, dpi=(600, 600), optimize=True)
print(png)
