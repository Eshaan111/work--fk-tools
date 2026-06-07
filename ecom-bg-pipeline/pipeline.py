import os
import io
import base64
import pandas as pd
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
# pyrefly: ignore [missing-import]
from rembg import remove
# pyrefly: ignore [missing-import]
from openai import OpenAI
# pyrefly: ignore [missing-import]
from dotenv import load_dotenv
from datetime import datetime

# -----------------------------
# LOAD ENV
# -----------------------------
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# -----------------------------
# PATHS
# -----------------------------
CSV_PATH = "background_styles.csv"
INPUT_DIR = "input_images"
OUTPUT_DIR = "outputs"
BG_DIR = "backgrounds"

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(BG_DIR, exist_ok=True)

# -----------------------------
# LOAD CSV
# -----------------------------
def load_styles():
    if os.path.exists(CSV_PATH):
        return pd.read_csv(CSV_PATH)
    return pd.DataFrame(columns=[
        "style_id","color_palette","props",
        "lighting_style","mood","notes","image_file"
    ])

# -----------------------------
# BUILD PROMPT
# -----------------------------
def build_prompt(df):
    return f"""
Generate a NEW premium ecommerce studio background.

Avoid repeating these styles:
{df.to_string(index=False)}

Requirements:
- Modern, minimal, high-end
- Muted colors
- Soft diffused lighting
- Clean center composition
- No product

Also return metadata as JSON:
color_palette, props, lighting_style, mood, notes
"""

# -----------------------------
# GENERATE BACKGROUND
# -----------------------------
def generate_background(prompt):
    result = client.images.generate(
        model="gpt-image-1",
        prompt=prompt,
        size="1024x1024"
    )

    img_data = base64.b64decode(result.data[0].b64_json)

    filename = f"bg_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    path = os.path.join(BG_DIR, filename)

    with open(path, "wb") as f:
        f.write(img_data)

    return path

# -----------------------------
# REMOVE BACKGROUND
# -----------------------------
def remove_bg(image_path):
    with open(image_path, "rb") as f:
        output = remove(f.read())

    return Image.open(io.BytesIO(output)).convert("RGBA")

# -----------------------------
# ADD SHADOW
# -----------------------------
def add_shadow(subject):
    shadow = subject.copy()
    alpha = shadow.split()[3]

    shadow = Image.new("RGBA", subject.size, (0,0,0,120))
    shadow.putalpha(alpha)
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))

    return shadow

# -----------------------------
# COMPOSITE
# -----------------------------
def composite(subject, background):
    bg = background.resize(subject.size).convert("RGBA")

    shadow = add_shadow(subject)
    bg.paste(shadow, (0,10), shadow)

    final = Image.alpha_composite(bg, subject)
    return final

# -----------------------------
# ENHANCE IMAGE
# -----------------------------
def enhance(img):
    img = ImageEnhance.Brightness(img).enhance(1.05)
    img = ImageEnhance.Contrast(img).enhance(1.1)
    img = ImageEnhance.Sharpness(img).enhance(1.1)
    return img

# -----------------------------
# APPEND TO CSV
# -----------------------------
def append_to_csv(df, bg_path):
    new_id = df["style_id"].max() + 1 if not df.empty else 1

    new_row = {
        "style_id": new_id,
        "color_palette": "generated",
        "props": "generated",
        "lighting_style": "generated",
        "mood": "generated",
        "notes": "auto-added",
        "image_file": bg_path
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_csv(CSV_PATH, index=False)

# -----------------------------
# MAIN RUN
# -----------------------------
def run():
    df = load_styles()

    print("Generating background...")
    prompt = build_prompt(df)
    bg_path = generate_background(prompt)

    background = Image.open(bg_path).convert("RGBA")

    images = sorted(os.listdir(INPUT_DIR))[:5]

    for i, img_name in enumerate(images):
        print(f"Processing {img_name}")

        img_path = os.path.join(INPUT_DIR, img_name)

        subject = remove_bg(img_path)
        final = composite(subject, background)
        final = enhance(final)

        output_path = os.path.join(OUTPUT_DIR, f"final_{i+1}.png")
        final.save(output_path)

    append_to_csv(df, bg_path)

    print("✅ Done! 5 images generated.")

# -----------------------------
if __name__ == "__main__":
    run()