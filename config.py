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
    MAIL_USERNAME = os.environ.get('MAIL_USER')
    MAIL_PASSWORD = os.environ.get('MAIL_PASS')
