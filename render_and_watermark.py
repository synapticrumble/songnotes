import asyncio
import sys
from pathlib import Path
from playwright.async_api import async_playwright
from PIL import Image, ImageDraw, ImageFont

# File paths
HTML_FILE = Path("output/song_notations.html")
PNG_FILE = Path("output/song_notations.png")
WM_FILE = Path("output/song_notations_wm.png")

async def html_to_png(html_file, png_file):
    """Convert HTML file to PNG screenshot"""
    try:
        if not html_file.exists():
            print(f"[ERROR] HTML file not found: {html_file}")
            return False
            
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()
            
            # Set viewport for consistent rendering
            await page.set_viewport_size({"width": 1200, "height": 800})
            
            # Read and set HTML content
            content = html_file.read_text(encoding="utf-8")
            await page.set_content(content, wait_until="networkidle")
            
            # Take full page screenshot
            await page.screenshot(
                path=str(png_file), 
                full_page=True,
                type="png"
            )
            await browser.close()
            
        print(f"[OK] Successfully rendered: {png_file}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error rendering HTML to PNG: {e}")
        return False

def watermark_image(png_path, wm_path, text="© 2025 Bansuri Notations"):
    """Add watermark to PNG image"""
    try:
        if not png_path.exists():
            print(f"[ERROR] PNG file not found: {png_path}")
            return False
            
        # Open image and convert to RGBA
        img = Image.open(png_path).convert("RGBA")
        
        # Create transparent overlay
        txt_layer = Image.new("RGBA", img.size, (255, 255, 255, 0))
        draw = ImageDraw.Draw(txt_layer)
        
        # Try to load a better font, fallback to default
        try:
            font = ImageFont.truetype("arial.ttf", 24)
        except:
            try:
                font = ImageFont.load_default()
            except:
                font = None
        
        w, h = img.size
        
        # Add watermarks in a grid pattern
        spacing_x, spacing_y = 400, 300
        for x in range(0, w, spacing_x):
            for y in range(0, h, spacing_y):
                # Semi-transparent white watermark
                draw.text(
                    (x + 20, y + 20), 
                    text, 
                    fill=(255, 255, 255, 80), 
                    font=font
                )
        
        # Combine original image with watermark
        watermarked = Image.alpha_composite(img, txt_layer)
        
        # Convert back to RGB and save
        final_img = watermarked.convert("RGB")
        final_img.save(wm_path, "PNG", optimize=True)
        
        print(f"[OK] Watermark added: {wm_path}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error adding watermark: {e}")
        return False

async def main():
    """Main execution function"""
    print("[INFO] Starting PNG generation process...")
    
    # Ensure output directory exists
    PNG_FILE.parent.mkdir(exist_ok=True)
    
    # Check if HTML file exists, if not suggest running the pipeline
    if not HTML_FILE.exists():
        print(f"[ERROR] HTML file not found: {HTML_FILE}")
        print("[INFO] Please run the conversion pipeline first:")
        print("   1. python reformat.py")
        print("   2. python convert_and_push.py")
        print("   3. python render_and_watermark.py")
        sys.exit(1)
    
    # Step 1: Convert HTML to PNG
    success = await html_to_png(HTML_FILE, PNG_FILE)
    if not success:
        print("[ERROR] Failed to generate PNG from HTML")
        sys.exit(1)
    
    # Step 2: Add watermark
    success = watermark_image(PNG_FILE, WM_FILE)
    if not success:
        print("[ERROR] Failed to add watermark")
        sys.exit(1)
    
    print("[OK] PNG generation completed successfully!")

if __name__ == "__main__":
    asyncio.run(main())
