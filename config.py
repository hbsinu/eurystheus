import os


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "eurystheus-dev-secret-key")
    UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
    ALLOWED_EXTENSIONS = {"pdf", "docx", "doc", "txt"}
    MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32 MB
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
