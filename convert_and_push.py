import mammoth
import os
import subprocess
from datetime import datetime

# === CONFIG ===
REPO_PATH = os.path.dirname(os.path.abspath(__file__))
INPUT_DOCX = os.path.join(REPO_PATH, "songs_reformatted.docx")
OUTPUT_HTML = os.path.join(REPO_PATH, "songs_reformatted.html")

# INPUT_DOCX = "P:\\ShareDownloads\\songnotes\\songs_reformatted.docx"
# OUTPUT_HTML = "P:\\ShareDownloads\\songnotes\\song_notations.html"
# REPO_PATH = "."  # current folder
COMMIT_MESSAGE = f"Auto-update HTML from Word on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

def convert_docx_to_html(input_path, output_path):
    with open(input_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value
    # add no-copy script
    protect_js = """
    <style>
    body { user-select: none; -webkit-user-select: none; }
    </style>
    <script>
    document.addEventListener('contextmenu', e => e.preventDefault());
    document.addEventListener('copy', e => e.preventDefault());
    document.addEventListener('cut', e => e.preventDefault());
    document.addEventListener('paste', e => e.preventDefault());
    </script>

    <script async src="https://www.googletagmanager.com/gtag/js?id=G-XXXXXXX"></script>
    <script>
    window.dataLayer = window.dataLayer || [];
    function gtag(){dataLayer.push(arguments);}
    gtag('js', new Date());
    gtag('config', 'G-XXXXXXX');
    </script>

    """
    html = html + protect_js
    with open(output_path, "w", encoding="utf-8") as html_file:
        html_file.write(html)
    print(f"✅ Converted and protected {input_path} → {output_path}")


# def convert_docx_to_html(input_path, output_path):
#     """Convert a .docx file to clean HTML using mammoth"""
#     with open(input_path, "rb") as docx_file:
#         result = mammoth.convert_to_html(docx_file)
#         html = result.value
#         messages = result.messages
#     with open(output_path, "w", encoding="utf-8") as html_file:
#         html_file.write(html)
#     print(f"✅ Converted {input_path} → {output_path}")
#     if messages:
#         print("Messages:", messages)

def git_commit_and_push(repo_path, message):
    """Commit and push changes to GitHub"""
    subprocess.run(["git", "-C", repo_path, "add", "."], check=True)
    subprocess.run(["git", "-C", repo_path, "commit", "-m", message], check=True)
    subprocess.run(["git", "-C", repo_path, "push"], check=True)
    print("🚀 Changes pushed to GitHub.")

if __name__ == "__main__":
    convert_docx_to_html(INPUT_DOCX, OUTPUT_HTML)
    git_commit_and_push(REPO_PATH, COMMIT_MESSAGE)
