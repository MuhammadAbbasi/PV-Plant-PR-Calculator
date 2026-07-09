import os
import sys
import subprocess
import time
import webbrowser

def check_python_packages():
    """Ensure required python modules are available."""
    required = ["pandas", "numpy", "openpyxl", "fastapi", "uvicorn"]
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"Librerie Python mancanti: {', '.join(missing)}")
        print("Installazione in corso...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install"] + missing, check=True)
            print("Installazione completata con successo.")
        except subprocess.CalledProcessError as e:
            print(f"Errore durante l'installazione delle librerie: {e}")
            sys.exit(1)

def build_frontend():
    """Check and compile the React frontend bundle if needed."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    frontend_dir = os.path.join(base_dir, "pr_dashboard", "frontend")
    dist_dir = os.path.join(frontend_dir, "dist")
    node_modules_dir = os.path.join(frontend_dir, "node_modules")

    # If dist folder exists, we can skip compilation to start faster
    if os.path.exists(dist_dir):
        print("Frontend già compilato (cartella dist presente). Salto la compilazione.")
        return

    print("Compilazione del frontend React in corso (questa operazione avviene solo al primo avvio)...")
    
    # 1. Check if node_modules exists, if not run npm install
    if not os.path.exists(node_modules_dir):
        print("Installazione dei pacchetti node (npm install)...")
        try:
            # Run npm install
            subprocess.run("npm install", shell=True, cwd=frontend_dir, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Errore durante npm install: {e}")
            sys.exit(1)
            
    # 2. Run npm run build
    print("Compilazione degli asset statici (npm run build)...")
    try:
        subprocess.run("npm run build", shell=True, cwd=frontend_dir, check=True)
        print("Compilazione completata con successo.")
    except subprocess.CalledProcessError as e:
        print(f"Errore durante la compilazione del frontend: {e}")
        sys.exit(1)

def start_server():
    """Start the FastAPI backend server using uvicorn and open browser."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Port requested by user to avoid conflicts
    port = 5896
    url = f"http://127.0.0.1:{port}"
    
    print("\n" + "="*50)
    print(f"Avvio del server PR Dashboard su {url}")
    print("="*50 + "\n")
    
    # Run uvicorn in a subprocess
    cmd = [sys.executable, "-m", "uvicorn", "pr_dashboard.backend.main:app", "--host", "127.0.0.1", "--port", str(port)]
    
    # Wait 1.5 seconds and open browser automatically
    def open_browser():
        time.sleep(2.0)
        print(f"Apertura automatica del browser all'indirizzo {url}...")
        webbrowser.open(url)
        
    import threading
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    try:
        subprocess.run(cmd, cwd=base_dir, check=True)
    except KeyboardInterrupt:
        print("\nPR Dashboard arrestato.")
    except Exception as e:
        print(f"Errore durante l'esecuzione del server: {e}")

if __name__ == "__main__":
    check_python_packages()
    build_frontend()
    start_server()
