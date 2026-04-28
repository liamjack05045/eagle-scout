from threading import Thread
import webview
from app import app

def start_flask():
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)

if __name__ == "__main__":
    # Launch Flask server in a background thread
    Thread(target=start_flask, daemon=True).start()
    
    # Create a native window that loads your Flask site
    webview.create_window("Script Formatter", "http://127.0.0.1:5000", width=1200, height=800)
    webview.start()
