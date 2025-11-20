import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from flask_login import LoginManager
from authlib.integrations.flask_client import OAuth
from dotenv import load_dotenv
from flask_bootstrap import Bootstrap5
from flask_wtf.csrf import CSRFProtect
from flask_ckeditor import CKEditor

csrf = CSRFProtect()

load_dotenv()


db = SQLAlchemy()
migrate = Migrate()
login_manager = LoginManager()
oauth = OAuth()
bootstrap = Bootstrap5()
ckeditor = CKEditor()

report_types = ['weekly', 'quarterly']

def create_app(config_object=None):
    app = Flask(__name__, template_folder='templates', static_folder='static')
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev')
    app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///reportit.db')
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


    db.init_app(app)
    migrate.init_app(app, db)
    csrf.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'

    
    #Bootstrap5(app)
    bootstrap.init_app(app)
    ckeditor.init_app(app)

    # OAuth (Google)
    oauth.init_app(app)
    oauth.register(
    name='google',
    client_id=os.getenv('GOOGLE_CLIENT_ID'),
    client_secret=os.getenv('GOOGLE_CLIENT_SECRET'),
    server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
    client_kwargs={'scope': 'openid email profile'},
    )

    # Blueprints
    from .routes import main_bp, auth_bp
    app.register_blueprint(main_bp)
    app.register_blueprint(auth_bp)

    # Create tables and seed report types inside the application context
    def create_report_types():
        from app.models import ReportType
        for t in report_types:
            if not ReportType.query.filter_by(type=t).first():
                db.session.add(ReportType(type=t))
        db.session.commit()

    with app.app_context():
        db.create_all()
        create_report_types()

    return app

    