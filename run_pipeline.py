import os
import subprocess
import sys
from datetime import datetime

def run_step(description, command):
    """Run a subprocess step with logging and error capture."""
    print(f"[INFO] Running {description}...")
    try:
        subprocess.run(command, check=True)
        print(f"[✅] {description} completed successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[❌] {description} failed:\n{e}")
        return False


def process_docx_if_available():
    """Process the BansuriMusic.docx if found, otherwise skip."""
    # Local and GitHub paths
    local_input = r"P:\ShareDownloads\BansuriMusic.docx"
    repo_input = "input/BansuriMusic.docx"
    output_file = "./songs_reformatted.docx"

    # Resolve actual existing file
    if os.path.exists(local_input):
        input_file = local_input
        print(f"[INFO] Found local input file: {input_file}")
    elif os.path.exists(repo_input):
        input_file = repo_input
        print(f"[INFO] Found repo input file: {input_file}")
    else:
        print("[⚠️] Input DOCX not found. Skipping reformat step.")
        return True  # skip gracefully instead of erroring out

    # Call reformat.py
    result = run_step("reformat.py", [sys.executable, "reformat.py", input_file, output_file])
    return result


def main():
    print("[INFO] Starting complete Bansuri processing pipeline...")

    steps = [
        ("Reformat Word document", process_docx_if_available),
        ("Convert DOCX → HTML", lambda: run_step("convert_and_push.py", [sys.executable, "convert_and_push.py"])),
        ("Post-process HTML protection", lambda: run_step("render_and_watermark.py", [sys.executable, "render_and_watermark.py"])),
    ]

    success_all = True
    for desc, func in steps:
        print(f"\n[STEP] {desc}")
        success = func()
        if not success:
            success_all = False
            # Do not exit early; continue to next step

    print("\n" + "=" * 60)
    if success_all:
        print("[🎉] Pipeline completed successfully at", datetime.now())
    else:
        print("[⚠️] Pipeline finished with some errors at", datetime.now())
    print("=" * 60)


if __name__ == "__main__":
    main()
