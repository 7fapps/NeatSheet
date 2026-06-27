"""
desktop.py
----------
Launches NeatSheet as a desktop app (Windows/Mac) instead of needing
to open a browser manually.

This works by:
  1. Starting the Flask server (app.py) in the background, silently.
  2. Opening a native window (using pywebview) that displays the
     Flask app, just like a regular desktop application.

Run with:  python desktop.py
(No need to also run app.py separately - this does both.)
"""

import threading
import webview

from app import app  # reuse the exact same Flask app from app.py


def start_flask_server():
    """
    Runs the Flask server in the background.
    use_reloader=False is important here - the auto-reload feature
    is meant for development in a browser, and can cause problems
    when running inside a packaged desktop app.
    """
    app.run(port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    # Start Flask in a separate background thread, so it doesn't
    # block the desktop window from opening.
    flask_thread = threading.Thread(target=start_flask_server)
    flask_thread.daemon = True  # closes automatically when the app quits
    flask_thread.start()

    # Open the native desktop window, pointing at our local Flask server.
    webview.create_window(
        "NeatSheet",
        "http://127.0.0.1:5000",
        width=900,
        height=700,
        resizable=True
    )
    webview.start()
