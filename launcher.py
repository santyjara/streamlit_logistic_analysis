import subprocess
import sys
import os
import webbrowser
import time
import threading
from pathlib import Path
import socket

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

def get_python_executable():
    """Get the correct Python executable path"""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller executable - find Python in PATH
        try:
            result = subprocess.run(["python", "--version"], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=5)
            if result.returncode == 0:
                return "python"
        except:
            pass
        
        # Try python3
        try:
            result = subprocess.run(["python3", "--version"], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=5)
            if result.returncode == 0:
                return "python3"
        except:
            pass
        
        return None
    else:
        # Running as regular Python script
        return sys.executable

def check_port_available(port=8501):
    """Check if the port is available"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
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

def wait_for_server(url, max_attempts=30):
    """Wait for the Streamlit server to be ready"""
    import urllib.request
    import urllib.error
    
    for attempt in range(max_attempts):
        try:
            urllib.request.urlopen(url, timeout=1)
            return True
        except (urllib.error.URLError, OSError):
            time.sleep(1)
    return False

def open_browser_when_ready(url):
    """Open browser only when server is ready"""
    def _open():
        print(" Waiting for server to start...")
        if wait_for_server(url):
            print(f" Opening browser: {url}")
            webbrowser.open(url)
        else:
            print(" Server didn't start in time")
    
    thread = threading.Thread(target=_open)
    thread.daemon = True
    thread.start()

def launch_streamlit():
    """Launch the Streamlit application"""
    app_path = get_app_path()
    
    if not os.path.exists(app_path):
        print(f"Error: Streamlit app not found at {app_path}")
        input("Press Enter to exit...")
        return
    
    # Get Python executable
    python_exe = get_python_executable()
    
    if python_exe is None:
        print("Error: Python not found. Please install Python and add it to PATH.")
        input("Press Enter to exit...")
        return
    
    # Find available port
    port = find_available_port()
    url = f"http://localhost:{port}"
    
    print("Starting Streamlit App...")
    print(f"App location: {app_path}")
    print(f"Will open: {url}")
    print("Please wait while the app loads...")
    
    # Prepare streamlit command
    cmd = [
        python_exe, "-m", "streamlit", "run",
        app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.runOnSave", "false",
        "--browser.gatherUsageStats", "false",
        "--server.address", "localhost",
        "--global.developmentMode", "false"
    ]
    
    try:
        # Start browser opener thread
        open_browser_when_ready(url)
        
        # Start Streamlit process
        print("Starting server...")
        
        # Create process with suppressed output
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
        
        print("Server started successfully!")
        print("App should open in your browser shortly...")
        print("Close this window to stop the server")
        
        # Wait for process to complete or be interrupted
        try:
            process.wait()
        except KeyboardInterrupt:
            print("\n Shutting down...")
            process.terminate()
            process.wait()
            
    except KeyboardInterrupt:
        print("\n Shutting down...")
    except FileNotFoundError:
        print(" Error: Streamlit not found. Make sure it's installed.")
        print("Try: pip install streamlit")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"Error starting app: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    print("=" * 50)
    print("Procesador de Archivos - Launcher")
    print("=" * 50)
    launch_streamlit()