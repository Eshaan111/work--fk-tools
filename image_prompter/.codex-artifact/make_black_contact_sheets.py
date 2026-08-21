from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

root = Path(r"C:\Users\ESHAAN\HAKUR\work--fk-tools\image_prompter\IMAGES-FINAL\BLACK-BAGGY")
out = Path(r"C:\Users\ESHAAN\HAKUR\work--fk-tools\image_prompter\.codex-artifact")
font = ImageFont.load_default()

for batch in range(22, 31):
    tiles = []
    for number in range(1, 6):
        path = root / str(batch) / f"{number}.png"
        image = Image.open(path).convert("RGB")
        image.thumbnail((240, 320))
        tile = Image.new("RGB", (260, 360), "white")
        x = (260 - image.width) // 2
        y = 28 + (320 - image.height) // 2
        tile.paste(image, (x, y))
        ImageDraw.Draw(tile).text((10, 8), f"current {number}.png", fill="black", font=font)
        tiles.append(tile)
    sheet = Image.new("RGB", (260 * 5, 360), "#dddddd")
    for index, tile in enumerate(tiles):
        sheet.paste(tile, (index * 260, 0))
    sheet.save(out / f"black-batch-{batch}-mapping.jpg", quality=90)
