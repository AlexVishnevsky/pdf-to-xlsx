"""Малює icon.ico: аркуш PDF із червоною міткою і зеленою таблицею Excel."""
import os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
S = 1024
im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
# аркуш з загнутим кутом
d.rounded_rectangle((150, 60, 800, 900), 40, fill=(250, 250, 250), outline=(110, 110, 110), width=22)
d.polygon([(640, 60), (800, 220), (640, 220)], fill=(205, 205, 205), outline=(110, 110, 110))
# червона мітка PDF
d.rounded_rectangle((90, 150, 520, 360), 30, fill=(210, 40, 40))
f = ImageFont.truetype(os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arialbd.ttf"), 150)
d.text((305, 255), "PDF", font=f, fill="white", anchor="mm")
# зелена таблиця Excel
x0, y0, x1, y1 = 420, 470, 980, 980
d.rounded_rectangle((x0, y0, x1, y1), 40, fill=(33, 115, 70))
for i in range(1, 4):
    y = y0 + (y1 - y0) * i / 4
    d.line((x0 + 40, y, x1 - 40, y), fill=(255, 255, 255), width=18)
for i in range(1, 3):
    x = x0 + (x1 - x0) * i / 3
    d.line((x, y0 + 40, x, y1 - 40), fill=(255, 255, 255), width=18)
im.save(os.path.join(HERE, "icon.ico"), sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
im.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, "icon_preview.png"))
print("ok")
