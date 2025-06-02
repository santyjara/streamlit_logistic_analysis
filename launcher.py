import subprocess
import sys
import os
import webbrowser
import time
import threading
from pathlib import Path

def get_app_path():
    """Get the path to the Streamlit app file"""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller executable
        bundle_dir = Path(sys._MEIPASS)
        app_path = bundle_dir / "analysis.py"
    else:
        # Running as regular Python script
        app_path = Path(__file__).parent / "analysis.py"
    
    return str(app_path)

def check_port_available(port=8501):
    """Check if the port is available"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('localhost', port))
            return True
        except OSError:
            return False

def find_available_port(start_port=8501):
    """Find an available port starting from start_port"""
    port = start_port
    while port < start_port + 100:  # Try 100 ports
        if check_port_available(port):
            return port
        port += 1
    return start_port  # Fallback

def open_browser_delayed(url, delay=3):
    """Open browser after a delay"""
    def _open():
        time.sleep(delay)
        print(f"Opening browser: {url}")
        webbrowser.open(url)
    
    thread = threading.Thread(target=_open)
    thread.daemon = True
    thread.start()

def launch_streamlit():
    """Launch the Streamlit application"""
    app_path = get_app_path()
    
    if not os.path.exists(app_path):
        print(f"❌ Error: Streamlit app not found at {app_path}")
        input("Press Enter to exit...")
        return
    
    # Find available port
    port = find_available_port()
    url = f"http://localhost:{port}"
    
    print("🚀 Starting Streamlit App...")
    print(f"📍 App location: {app_path}")
    print(f"🌐 Will open: {url}")
    print("⏳ Please wait while the app loads...")
    
    # Prepare streamlit command
    cmd = [
        sys.executable, "-m", "streamlit", "run",
        app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.address", "localhost"
    ]
    
    try:
        # Start browser opener thread
        open_browser_delayed(url, delay=4)
        
        # Start Streamlit process
        print("⚡ Starting server...")
        process = subprocess.run(cmd, check=False)
        
    except KeyboardInterrupt:
        print("\n⏹️  Shutting down...")
    except FileNotFoundError:
        print("❌ Error: Streamlit not found. Make sure it's installed.")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"❌ Error starting app: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    print("=" * 50)
    print("📊 Procesador de Archivos - Launcher")
    print("=" * 50)
    launch_streamlit()