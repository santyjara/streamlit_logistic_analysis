import subprocess
import sys
import os
import webbrowser
import threading
import time
import socket
from contextlib import closing

def find_free_port():
    """Find a free port to run Streamlit on"""
    with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def open_browser(port):
    """Open browser after Streamlit starts"""
    time.sleep(3)  # Wait for Streamlit to start
    webbrowser.open(f'http://localhost:{port}')

def main():
    # Prevent recursive calls
    if os.environ.get('STREAMLIT_LAUNCHER_RUNNING'):
        return
    
    os.environ['STREAMLIT_LAUNCHER_RUNNING'] = '1'
    
    # Get the directory where the executable is located
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        bundle_dir = sys._MEIPASS
        python_executable = sys.executable
    else:
        # Running as script
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        python_executable = sys.executable
    
    # Path to your main Streamlit app
    app_path = os.path.join(bundle_dir, 'analysis.py')
    
    # Check if the app file exists
    if not os.path.exists(app_path):
        print(f"Error: Streamlit app not found at {app_path}")
        input("Press Enter to exit...")
        return
    
    # Find a free port
    port = find_free_port()
    
    # Start browser in a separate thread
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    
    # Start Streamlit with specific port and headless mode
    try:
        subprocess.run([
            python_executable, '-m', 'streamlit', 'run', 
            app_path, 
            f'--server.port={port}',
            '--server.headless=true',
            '--server.enableCORS=false',
            '--server.enableXsrfProtection=false'
        ])
    except KeyboardInterrupt:
        print("Application stopped by user")
    except Exception as e:
        print(f"Error running Streamlit: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()