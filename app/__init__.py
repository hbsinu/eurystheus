from flask import Flask
from config import Config


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    from app.routes.main import main_bp
    from app.routes.researcher import researcher_bp

    app.register_blueprint(main_bp)
    app.register_blueprint(researcher_bp, url_prefix="/researcher")

    return app
