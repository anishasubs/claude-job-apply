"""Render the get-me-a-job skill banner as a PNG, matching the Claude Code launch style.

Pixel-art tiny person at a desk in coral, dark terminal background, monospace text.
"""
from PIL import Image, ImageDraw, ImageFont
import os

# Canvas — 2x for crispness on Retina/HiDPI displays
SCALE = 2
W = 1200 * SCALE
H = 240 * SCALE
BG = (10, 10, 10)
FG = (245, 245, 245)
DIM = (138, 138, 138)
ACCENT = (212, 193, 112)        # yellow (matches Claude prompt style)
CORAL = (228, 124, 95)          # mascot color

img = Image.new("RGB", (W, H), BG)
draw = ImageDraw.Draw(img)

# Try a few fonts in order of preference
def find_font(size):
    candidates = [
        r"C:\Windows\Fonts\CascadiaCode.ttf",
        r"C:\Windows\Fonts\CascadiaMono.ttf",
        r"C:\Windows\Fonts\consola.ttf",
        r"C:\Windows\Fonts\consolab.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except OSError:
                continue
    return ImageFont.load_default()

f_prompt = find_font(28 * SCALE)
f_title  = find_font(30 * SCALE)
f_dim    = find_font(24 * SCALE)

# ── Top row: PS prompt ──────────────────────────────────────────────
prompt_y = 30 * SCALE
draw.text((40 * SCALE, prompt_y), "PS C:\\Users\\Anisha> ", font=f_prompt, fill=FG)
# Measure prefix to position the command in accent color
prefix = "PS C:\\Users\\Anisha> "
prefix_w = draw.textlength(prefix, font=f_prompt)
draw.text((40 * SCALE + prefix_w, prompt_y), "/get-me-a-job", font=f_prompt, fill=ACCENT)

# ── Pixel mascot: tiny person at a desk ─────────────────────────────
# Grid is 28 wide × 14 tall. Each cell = PX pixels.
PX = 8 * SCALE
mascot_left = 70 * SCALE
mascot_top  = 90 * SCALE

# Coffee mug with handle, hot steam curls, and a visible coffee surface.
# 22 cols × 14 rows. # = coral, . = transparent.
GRID = [
    "......................",  # 0
    "....##..##..##........",  # 1   steam wisp tops
    "...#...#...#..........",  # 2   steam wave (shifted left)
    "....##..##..##........",  # 3   steam tails
    "......................",  # 4   gap
    "##############........",  # 5   rim
    "##############........",  # 6   coffee surface
    "#............#####....",  # 7   wall + handle top
    "#............#...#....",  # 8   wall + handle middle
    "#............#...#....",  # 9   wall + handle middle
    "#............#####....",  # 10  wall + handle bottom
    "#............#........",  # 11  wall
    "#............#........",  # 12  wall
    "##############........",  # 13  bottom edge
]

for r, row in enumerate(GRID):
    for c, ch in enumerate(row):
        if ch == "#":
            x = mascot_left + c * PX
            y = mascot_top  + r * PX
            draw.rectangle([x, y, x + PX - 1, y + PX - 1], fill=CORAL)

# ── Right side: title + subtitle + repo path ────────────────────────
text_left = mascot_left + 30 * PX
title_y   = 88 * SCALE

draw.text((text_left, title_y), "get-me-a-job", font=f_title, fill=FG)
title_w = draw.textlength("get-me-a-job", font=f_title)
draw.text((text_left + title_w + 12 * SCALE, title_y + 4 * SCALE), "v1.0", font=f_dim, fill=DIM)

draw.text((text_left, title_y + 42 * SCALE), "Tailor resumes · cover letters · outreach", font=f_dim, fill=DIM)
draw.text((text_left, title_y + 76 * SCALE), "anishasubs/get-me-a-job", font=f_dim, fill=DIM)

# Save
out = os.path.join(os.path.dirname(__file__), "banner.png")
img.save(out, "PNG", optimize=True)
print(f"Wrote {out}  ({W}×{H})")
