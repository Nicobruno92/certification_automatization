from PIL import Image, ImageDraw, ImageFont
import os

img = Image.open("congreso_neurociencias_2026/certificate_poster.png").convert("RGB")
draw = ImageDraw.Draw(img)

# draw a vertical line with ticks every 100px
for y in range(0, 2250, 100):
    draw.line([(100, y), (200, y)], fill="red", width=5)
    draw.text((220, y-10), str(y), fill="red", font=ImageFont.load_default(size=40))

img.save("congreso_neurociencias_2026/ruler.png")
