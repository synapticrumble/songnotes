"""
Local File Watcher Pipeline
Monitors BansuriMusic.docx for changes and triggers the processing pipeline
"""

import time
import subprocess
import sys
import os
from pathlib import Path
from datetime import datetime

# Try different polling approaches
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False
    print("⚠️ Watchdog not available, using polling method")

# Configuration
WATCH_FILE = Path(r"P:\ShareDownloads\BansuriMusic.docx")
WATCH_DIR = WATCH_FILE.parent
PIPELINE_SCRIPTS = [
    "reformat.py",
    "convert_and_push.py"
]

class DocxFileHandler:
    def __init__(self):
        self.last_processed = 0
        self.last_modified = 0
        self.cooldown = 5  # 5 second cooldown to avoid multiple triggers
        
    def check_file_changed(self):
        """Check if file has been modified using polling"""
        try:
            if not WATCH_FILE.exists():
                return False
                
            current_mtime = WATCH_FILE.stat().st_mtime
            current_time = time.time()
            
            # Check if file was modified and cooldown has passed
            if (current_mtime > self.last_modified and 
                current_time - self.last_processed > self.cooldown):
                
                self.last_modified = current_mtime
                self.last_processed = current_time
                print(f"\n🔔 Detected change in {WATCH_FILE.name} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                self.run_pipeline()
                return True
            elif current_mtime > self.last_modified:
                print(f"⏳ Cooldown active, skipping trigger")
                self.last_modified = current_mtime
                
        except Exception as e:
            print(f"❌ Error checking file: {e}")
        return False
    
    def on_modified(self, event):
        """Watchdog event handler"""
        if event.is_directory:
            return
            
        file_path = Path(event.src_path)
        
        # Check if it's our target file (case-insensitive comparison)
        if file_path.name.lower() == WATCH_FILE.name.lower():
            current_time = time.time()
            
            # Cooldown check
            if current_time - self.last_processed < self.cooldown:
                print(f"⏳ Cooldown active, skipping trigger")
                return
                
            self.last_processed = current_time
            print(f"\n🔔 Detected change in {WATCH_FILE.name} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            self.run_pipeline()
    
    def run_pipeline(self):
        """Execute the processing pipeline"""
        print("🚀 Starting processing pipeline...")
        
        for script in PIPELINE_SCRIPTS:
            try:
                # Check if script exists
                script_path = Path(script)
                if not script_path.exists():
                    print(f"❌ Script not found: {script}")
                    print(f"   Current directory: {os.getcwd()}")
                    print(f"   Looking for: {script_path.absolute()}")
                    continue
                
                print(f"📝 Running {script}...")
                
                # Run script with proper error handling
                result = subprocess.run([
                    sys.executable, str(script_path)
                ], capture_output=True, text=True, timeout=300, cwd=os.getcwd())
                
                if result.returncode == 0:
                    print(f"✅ {script} completed successfully")
                    if result.stdout.strip():
                        print(f"   Output: {result.stdout.strip()}")
                else:
                    print(f"❌ {script} failed with return code {result.returncode}")
                    if result.stderr.strip():
                        print(f"   Error: {result.stderr.strip()}")
                    if result.stdout.strip():
                        print(f"   Output: {result.stdout.strip()}")
                    break
                    
            except subprocess.TimeoutExpired:
                print(f"⏰ {script} timed out after 5 minutes")
                break
            except FileNotFoundError as e:
                print(f"❌ Python executable or script not found: {e}")
                break
            except Exception as e:
                print(f"❌ Unexpected error running {script}: {e}")
                break
        
        print("🏁 Pipeline execution completed\n")

def main():
    """Main file watcher function"""
    print("🎵 Bansuri Music File Watcher Starting...")
    print(f"📁 Current working directory: {os.getcwd()}")
    
    # Validate watch file exists
    if not WATCH_FILE.exists():
        print(f"❌ Watch file not found: {WATCH_FILE}")
        print("Please update WATCH_FILE path in the script")
        return False
    
    # Validate watch directory exists
    if not WATCH_DIR.exists():
        print(f"❌ Watch directory not found: {WATCH_DIR}")
        return False
    
    # Validate pipeline scripts exist
    missing_scripts = []
    for script in PIPELINE_SCRIPTS:
        if not Path(script).exists():
            missing_scripts.append(script)
    
    if missing_scripts:
        print(f"⚠️ Warning: Missing scripts: {missing_scripts}")
        print("Pipeline will skip missing scripts")
    
    print(f"👀 Watching for changes in: {WATCH_FILE}")
    print(f"📁 Watch directory: {WATCH_DIR}")
    print(f"🔄 Pipeline will run: {' → '.join(PIPELINE_SCRIPTS)}")
    
    # Initialize handler
    handler = DocxFileHandler()
    
    if WATCHDOG_AVAILABLE:
        # Try watchdog first
        print("🔍 Using watchdog file monitoring...")
        print("Press Ctrl+C to stop watching...\n")
        
        observer = None
        try:
            if WATCHDOG_AVAILABLE:
                from watchdog.events import FileSystemEventHandler
                
                class WatchdogHandler(FileSystemEventHandler):
                    def __init__(self, handler):
                        self.handler = handler
                    
                    def on_modified(self, event):
                        self.handler.on_modified(event)
                
                observer = Observer()
                observer.schedule(WatchdogHandler(handler), str(WATCH_DIR), recursive=False)
                observer.start()
                print("✅ Watchdog file watcher started successfully")
                
                # Keep the script running
                while True:
                    time.sleep(1)
                    
        except KeyboardInterrupt:
            print("\n🛑 Stopping file watcher...")
        except Exception as e:
            print(f"⚠️ Watchdog failed: {e}")
            print("🔄 Falling back to polling method...")
            WATCHDOG_AVAILABLE = False
        finally:
            if observer:
                try:
                    observer.stop()
                    observer.join(timeout=2)
                except:
                    pass
    
    if not WATCHDOG_AVAILABLE:
        # Fallback to polling
        print("🔍 Using polling file monitoring...")
        print("Press Ctrl+C to stop watching...\n")
        
        # Initialize last modified time
        if WATCH_FILE.exists():
            handler.last_modified = WATCH_FILE.stat().st_mtime
        
        print("✅ Polling file watcher started successfully")
        
        try:
            while True:
                handler.check_file_changed()
                time.sleep(2)  # Check every 2 seconds
        except KeyboardInterrupt:
            print("\n🛑 Stopping file watcher...")
    
    print("👋 File watcher stopped")
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

