import mammoth
import os
import subprocess
from datetime import datetime

# === CONFIG ===
REPO_PATH = os.path.dirname(os.path.abspath(__file__))
INPUT_DOCX = os.path.join(REPO_PATH, "songs_reformatted.docx")
OUTPUT_HTML = os.path.join(REPO_PATH, "song_notations.html")

# INPUT_DOCX = "P:\\ShareDownloads\\songnotes\\songs_reformatted.docx"
# OUTPUT_HTML = "P:\\ShareDownloads\\songnotes\\song_notations.html"
# REPO_PATH = "."  # current folder
COMMIT_MESSAGE = f"Auto-update HTML from Word on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

def convert_docx_to_html(input_path, output_path):
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
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


    <script>
// Add this to your GitHub Pages
(function() {
    // REPLACE with your actual Apps Script URL from deployment
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxexIuN99-anjKyMhITQF-nrda2fPZV4gbak8i0kRA/exec';
    
    const visitData = {
        url: window.location.href,
        title: document.title,
        timestamp: new Date().toISOString(),
        referrer: document.referrer || 'direct',
        screen: `${screen.width}x${screen.height}`,
        language: navigator.language
    };
    
    // Use sendBeacon for better reliability
    if (navigator.sendBeacon) {
        const blob = new Blob([JSON.stringify(visitData)], {type: 'application/json'});
        navigator.sendBeacon(APPS_SCRIPT_URL, blob);
    } else {
        // Fallback to fetch
        fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Origin': window.location.origin
            },
            body: JSON.stringify(visitData),
            mode: 'no-cors' // Important for cross-origin requests
        }).catch(err => console.log('Visit tracking failed:', err));
    }
})();
</script>

    """
    html = html + protect_js
    with open(output_path, "w", encoding="utf-8") as html_file:
        html_file.write(html)
    print(f"[OK] Converted and protected {input_path} -> {output_path}")

def git_commit_and_push(repo_path, message):
    """Commit and push changes to GitHub"""
    subprocess.run(["git", "-C", repo_path, "add", "."], check=True)
    subprocess.run(["git", "-C", repo_path, "commit", "-m", message], check=True)
    subprocess.run(["git", "-C", repo_path, "push"], check=True)
    print("[OK] Changes pushed to GitHub.")

if __name__ == "__main__":
    convert_docx_to_html(INPUT_DOCX, OUTPUT_HTML)
    
    # Skip git operations in CI environment
    if not os.getenv('GITHUB_ACTIONS'):
        git_commit_and_push(REPO_PATH, COMMIT_MESSAGE)
    else:
        print("[INFO] Running in GitHub Actions - skipping git operations")
