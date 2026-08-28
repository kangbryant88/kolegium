"""
Inyecta (o actualiza) el usuario administrador real de la EEP Daniel
O'Leary. Es un script de una sola ejecucion manual, nunca se importa
desde el resto de la app.

Uso:
    python seed_admin.py

Usa los modelos de SQLAlchemy vía create_app(), asi que escribe en lo que
sea que SQLALCHEMY_DATABASE_URI apunte en ese momento (hoy: SQLite local
en roboclass.db; el dia que migren a MariaDB, este script no cambia).

La contraseña se pide de forma interactiva (getpass) y nunca queda escrita
en este archivo ni en el historial de git -- no hardcodees credenciales
reales en un script versionado.
"""
import getpass
import sys

from werkzeug.security import generate_password_hash

from app import create_app
from app.extensions import db
from app.models import Rol, Usuario

NOMBRE_COMPLETO = 'Michael Brian Laprea Torrealba'
USERNAME = 'mias'
EMAIL = 'mias@kolegium.com'
ROL_NOMBRE = 'Administrador Supremo'
AREA_TRABAJO = 'Administrador Supremo'


def main():
    app = create_app()

    with app.app_context():
        rol = Rol.query.filter_by(nombre=ROL_NOMBRE).first()
        if not rol:
            sys.exit(
                f'No existe el rol "{ROL_NOMBRE}". Arranca la app una vez '
                'primero (create_app siembra los roles) y vuelve a correr esto.'
            )

        password = getpass.getpass(f'Contraseña para "{USERNAME}": ')
        password_confirm = getpass.getpass('Confírmala: ')

        if password != password_confirm:
            sys.exit('Las contraseñas no coinciden. No se guardó nada.')
        if len(password) < 8:
            sys.exit('Usa al menos 8 caracteres. No se guardó nada.')

        usuario = Usuario.query.filter(
            (Usuario.username == USERNAME) | (Usuario.email == EMAIL)
        ).first()

        if usuario:
            print(f'Ya existe (id={usuario.id}). Actualizando datos y contraseña...')
        else:
            usuario = Usuario(username=USERNAME, email=EMAIL)
            db.session.add(usuario)
            print('Usuario nuevo, creando...')

        usuario.nombre_completo = NOMBRE_COMPLETO
        usuario.username = USERNAME
        usuario.email = EMAIL
        usuario.area_trabajo = AREA_TRABAJO
        usuario.rol_id = rol.id
        usuario.activo = True
        usuario.password = generate_password_hash(password)

        db.session.commit()
        print(f'Listo: "{USERNAME}" (id={usuario.id}), rol "{ROL_NOMBRE}".')


if __name__ == '__main__':
    main()
