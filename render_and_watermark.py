import asyncio
from playwright.async_api import async_playwright
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

HTML_FILE = Path("output/song_notations.html")
PNG_FILE = Path("output/song_notations.png")
WM_FILE = Path("output/song_notations_wm.png")

async def html_to_png(html_file, png_file):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        content = html_file.read_text(encoding="utf-8")
        await page.set_content(content)
        await page.screenshot(path=str(png_file), full_page=True)
        await browser.close()
    print(f"📸 Rendered {png_file}")

def watermark_image(png_path, wm_path, text="© 2025 YourCompany"):
    img = Image.open(png_path).convert("RGBA")
    txt = Image.new("RGBA", img.size, (255,255,255,0))
    draw = ImageDraw.Draw(txt)
    font = ImageFont.load_default()
    w, h = img.size
    for x in range(0, w, 400):
        for y in range(0, h, 400):
            draw.text((x, y), text, fill=(255,255,255,100), font=font)
    combined = Image.alpha_composite(img, txt)
    combined.convert("RGB").save(wm_path, "PNG")
    print(f"💧 Watermark added → {wm_path}")

if __name__ == "__main__":
    asyncio.run(html_to_png(HTML_FILE, PNG_FILE))
    watermark_image(PNG_FILE, WM_FILE)
