#!/usr/bin/env python3
"""
Master pipeline script - runs all processing steps in order
"""
import subprocess
import sys
from pathlib import Path

def run_script(script_name):
    """Run a Python script and handle errors"""
    try:
        print(f"[INFO] Running {script_name}...")
        result = subprocess.run([sys.executable, script_name], 
                              capture_output=True, text=True, check=True)
        print(f"[OK] {script_name} completed successfully")
        if result.stdout.strip():
            print(result.stdout.strip())
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] {script_name} failed:")
        if e.stderr:
            print(e.stderr)
        if e.stdout:
            print(e.stdout)
        return False

def main():
    """Run the complete pipeline"""
    print("[INFO] Starting complete Bansuri processing pipeline...")
    
    scripts = [
        "reformat.py",
        "convert_and_push.py", 
        "render_and_watermark.py"
    ]
    
    for script in scripts:
        if not Path(script).exists():
            print(f"[ERROR] Script not found: {script}")
            continue
            
        if not run_script(script):
            print(f"[ERROR] Pipeline failed at {script}")
            sys.exit(1)
    
    print("[OK] Complete pipeline finished successfully!")

if __name__ == "__main__":
    main()
