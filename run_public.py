"""
run_public.py
Run this file in VS Code to start the app AND get a public URL.
Usage: python run_public.py
"""
import subprocess
import sys
import time
import os
import threading

# ── CONFIG ──────────────────────────────────────────────
NGROK_AUTH_TOKEN = ""   # paste your free token from https://dashboard.ngrok.com/get-started/your-authtoken
PORT = 8501
# ────────────────────────────────────────────────────────

def start_streamlit():
    python = sys.executable
    subprocess.Popen(
        [python, "-m", "streamlit", "run", "app.py",
         "--server.port", str(PORT),
         "--server.headless", "true"],
        cwd=os.path.dirname(os.path.abspath(__file__))
    )

def start_tunnel():
    from pyngrok import ngrok, conf

    if NGROK_AUTH_TOKEN:
        ngrok.set_auth_token(NGROK_AUTH_TOKEN)

    print("\n⏳ Waiting for Streamlit to start...")
    time.sleep(5)

    tunnel = ngrok.connect(PORT, "http")
    public_url = tunnel.public_url

    print("\n" + "="*60)
    print("🐳  AI Docker NL Health Dashboard is LIVE!")
    print("="*60)
    print(f"🌐  Public URL  : {public_url}")
    print(f"💻  Local URL   : http://localhost:{PORT}")
    print("="*60)
    print("Share the Public URL with anyone to access your app.")
    print("Press Ctrl+C to stop.\n")

    try:
        ngrok.connect(PORT)
        input()
    except KeyboardInterrupt:
        print("\n🛑 Shutting down...")
        ngrok.kill()

if __name__ == "__main__":
    print("🚀 Starting AI Docker NL Dashboard...")
    start_streamlit()
    start_tunnel()
