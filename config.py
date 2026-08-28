import os
from dotenv import load_dotenv

basedir = os.path.abspath(os.path.dirname(__file__))

# Cargar variables de entorno desde .env con ruta absoluta
env_path = os.path.join(basedir, '.env')
load_dotenv(env_path)

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-fallback-key-cambiar-en-produccion')
    
    # Forzamos la ruta al archivo que está dentro de la raíz
    db_path = os.path.join(basedir, 'roboclass.db')
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + db_path
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # Carpeta para guardar los documentos privados
    UPLOAD_FOLDER = os.path.join(basedir, 'uploads', 'personal')
    
    # Configuraciones de correo
    MAIL_SERVER = os.environ.get('MAIL_SERVER', 'smtp.gmail.com')
    MAIL_PORT = int(os.environ.get('MAIL_PORT', 587))
    MAIL_USE_TLS = os.environ.get('MAIL_USE_TLS', 'True').lower() in ('true', '1', 'yes')
    MAIL_USERNAME = os.environ.get('MAIL_USER', 'eepdanieloleary9@gmail.com')
    MAIL_PASSWORD = os.environ.get('MAIL_PASS', '')
    MAIL_DEFAULT_SENDER = os.environ.get('MAIL_USER', 'eepdanieloleary9@gmail.com')


class DevelopmentConfig(Config):
    DEBUG = True


class ProductionConfig(Config):
    DEBUG = False


class TestingConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///:memory:'


config_by_name = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
}
