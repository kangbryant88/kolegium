"""Script de diagnóstico para verificar email"""
from app import app, db
from models import Usuario, TokenRecuperacion
from flask_mail import Message
from extensions import mail

with app.app_context():
    # 1. Verificar tabla TokenRecuperacion
    try:
        count = TokenRecuperacion.query.count()
        print(f"OK: Tabla TokenRecuperacion existe ({count} registros)")
    except Exception as e:
        print(f"ERROR tabla: {e}")
        db.create_all()
        print("Tabla creada con db.create_all()")

    # 2. Verificar config de correo
    username = app.config.get("MAIL_USERNAME")
    password = app.config.get("MAIL_PASSWORD")
    print(f"MAIL_SERVER: {app.config.get('MAIL_SERVER')}")
    print(f"MAIL_PORT: {app.config.get('MAIL_PORT')}")
    print(f"MAIL_USE_TLS: {app.config.get('MAIL_USE_TLS')}")
    print(f"MAIL_USERNAME: {username}")
    print(f"MAIL_PASSWORD: {'***' + password[-4:] if password else 'NO CONFIGURADA'}")

    # 3. Intentar enviar correo de prueba
    try:
        msg = Message("Test Kolegium", sender=username, recipients=[username])
        msg.body = "Prueba de correo desde Kolegium - diagnóstico"
        mail.send(msg)
        print("OK: Correo enviado exitosamente!")
    except Exception as e:
        print(f"ERROR correo: {e}")
        import traceback
        traceback.print_exc()
