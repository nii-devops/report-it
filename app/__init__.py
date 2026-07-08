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
from flask_mail import Mail
from sqlalchemy import inspect, text
from config import Config


csrf = CSRFProtect()

load_dotenv()


db = SQLAlchemy()
migrate = Migrate()
login_manager = LoginManager()
oauth = OAuth()
bootstrap = Bootstrap5()
ckeditor = CKEditor()
mail = Mail()
report_types = ['weekly', 'quarterly']
default_categories = [
    'Printers and Peripherals',
    'Network',
    'Meeting',
    'Software',
    'General IT Support',
    'Audio-Visual',
    'Procurement',
    'Hardware',
    'Others',
]

def create_app(config_object=None):
    app = Flask(__name__, template_folder='templates', static_folder='static')
    #app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev')
    #app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///reportit.db')
    #app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

    # Load configuration
    if config_object:
        app.config.from_object(config_object)  # Use the config you pass
    else:
        app.config.from_object(Config)  # Default fallback


    db.init_app(app)
    migrate.init_app(app, db)
    csrf.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'
    mail.init_app(app)
    print("DATABASE URI:", app.config['SQLALCHEMY_DATABASE_URI'])
    
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

    def create_categories():
        from app.models import Category
        for c in default_categories:
            if not Category.query.filter_by(name=c).first():
                db.session.add(Category(name=c))
        db.session.commit()

    def ensure_task_category_column():
        # db.create_all() only creates missing tables — it won't add a new
        # column to a 'tasks' table that already exists on a live database,
        # so the category_id FK is added here if it isn't there yet.
        inspector = inspect(db.engine)
        if 'tasks' not in inspector.get_table_names():
            return
        existing_columns = [col['name'] for col in inspector.get_columns('tasks')]
        if 'category_id' not in existing_columns:
            db.session.execute(text('ALTER TABLE tasks ADD COLUMN category_id INTEGER REFERENCES categories(id)'))
            db.session.commit()

    with app.app_context():
        db.create_all()
        ensure_task_category_column()
        create_report_types()
        create_categories()

    return app

    