from pathlib import Path
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "site" / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

width, height = 900, 500
image = Image.new("RGB", (width, height), "#dff5ec")
draw = ImageDraw.Draw(image)

for y in range(height):
    mix = y / height
    r = int(223 * (1 - mix) + 244 * mix)
    g = int(245 * (1 - mix) + 247 * mix)
    b = int(236 * (1 - mix) + 245 * mix)
    draw.line([(0, y), (width, y)], fill=(r, g, b))

draw.rectangle((0, 350, width, height), fill="#bdded0")
draw.rectangle((80, 120, 820, 360), fill="#ffffff", outline="#9dcfbd", width=8)
draw.rectangle((130, 165, 770, 360), fill="#eaf7f1")

for x in range(160, 720, 120):
    draw.rectangle((x, 190, x + 70, 310), fill="#147f96")
    draw.rectangle((x + 10, 204, x + 60, 296), fill="#d8f1f6")

draw.rectangle((395, 230, 505, 360), fill="#0c8f63")
draw.ellipse((420, 258, 480, 318), fill="#dff5ec")

for x, color in [(180, "#d84f45"), (260, "#e4a927"), (640, "#0c8f63"), (720, "#147f96")]:
    draw.ellipse((x, 330, x + 44, 374), fill=color)
    draw.rectangle((x + 10, 370, x + 34, 430), fill="#17201b")

draw.polygon([(90, 120), (450, 40), (810, 120)], fill="#17201b")
draw.rectangle((58, 112, 842, 140), fill="#17201b")
draw.rectangle((58, 140, 92, 360), fill="#17201b")
draw.rectangle((808, 140, 842, 360), fill="#17201b")

image.save(ASSETS / "campus.png", optimize=True)
