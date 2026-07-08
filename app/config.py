
from dotenv import load_dotenv
load_dotenv()
import os



class Config:

    # Flask security
    SECRET_KEY = os.environ.get("SECRET_KEY")

    # Database
    #SQLALCHEMY_DATABASE_URI = os.environ.get("SQLITE_DATABASE_URL", 'sqlite:///reportit.db')
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL',)
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    #API Keys
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
    ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

    # Mail settings
    MAIL_SERVER = os.environ.get("MAIL_SERVER", "smtp.gmail.com")
    MAIL_PORT = int(os.environ.get("MAIL_PORT", 587))
    MAIL_USE_TLS = True
    MAIL_USE_SSL = False
    MAIL_USERNAME = os.environ.get("MAIL_USERNAME")
    MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD")
    MAIL_DEFAULT_SENDER = os.environ.get("MAIL_USERNAME")


class DevelopmentConfig(Config):
    DEBUG = False

class ProductionConfig(Config):
    DEBUG = True 

