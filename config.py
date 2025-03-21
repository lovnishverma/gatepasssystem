import os

class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "your-default-secret-key")  # Required for Flask-WTF forms
    WTF_CSRF_ENABLED = False  # Completely disable CSRF protection
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    STATIC_DIR = os.path.join(BASE_DIR, "static")
    LOG_PATH = os.path.join(BASE_DIR, "app.log")
    DOMAIN = "https://ravinder2115115.pythonanywhere.com"
