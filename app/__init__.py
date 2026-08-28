import os

from flask import Flask

from config import config_by_name
from app.extensions import db, migrate, mail


def create_app(config_name=None):
    config_name = config_name or os.environ.get('FLASK_ENV', 'development')

    app = Flask(__name__)
    app.config.from_object(config_by_name[config_name])

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    db.init_app(app)
    migrate.init_app(app, db)
    mail.init_app(app)

    from app.web.auth import auth_bp
    from app.web.admin import admin_bp
    from app.web.academico import academico_bp
    from app.web.main import register_routes
    from app.api import api_bp
    from app.api.auth import auth_api_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(academico_bp)
    app.register_blueprint(api_bp)
    app.register_blueprint(auth_api_bp)
    register_routes(app)

    with app.app_context():
        _seed_datos_iniciales()

    return app


def _seed_datos_iniciales():
    from app.models import Rol, Usuario

    # Mapa maestro de permisos por rol
    PERMISOS_POR_ROL = {
        'Administrador Supremo': 'dashboard,planificador,asistencia,defensoria,configuracion,admin,anuncios',
        'Equipo Directivo (Dirección)': 'dashboard,asistencia,configuracion,anuncios',
        'Administrativo': 'dashboard_general,asistencia,anuncios,planificador',
        'Docente de Aula': 'planificador,asistencia,anuncios',
        'Docente Especialista': 'planificador,asistencia,anuncios',
        'Defensoría Estudiantil': 'dashboard_general,defensoria',
        'Obrero': 'dashboard_general',
        'Personal de Vigilancia': 'dashboard_general',
        'Personal de Cocina': 'dashboard_general',
    }

    if not Rol.query.first():
        for nombre, permisos in PERMISOS_POR_ROL.items():
            db.session.add(Rol(nombre=nombre, permisos=permisos))
        db.session.commit()
    else:
        # Limpieza agresiva de roles "Equipo Directivo" duplicados
        roles_ed = Rol.query.filter(Rol.nombre.in_(['Equipo Directivo', 'Equipo Directivo "Dirección"', 'Equipo Directivo / Administrativo', 'Equipo Directivo (Dirección)'])).all()
        if roles_ed:
            # Mantener solo uno y renombrarlo, reasignar usuarios y borrar el resto
            rol_principal = roles_ed[0]
            rol_principal.nombre = 'Equipo Directivo (Dirección)'

            for rol_dup in roles_ed[1:]:
                # Reasignar usuarios al rol principal
                usuarios_afectados = Usuario.query.filter_by(rol_id=rol_dup.id).all()
                for u in usuarios_afectados:
                    u.rol_id = rol_principal.id
                db.session.delete(rol_dup)

            db.session.commit()

        # Migración: sincronizar permisos de roles existentes y crear faltantes
        for nombre, permisos in PERMISOS_POR_ROL.items():
            rol = Rol.query.filter_by(nombre=nombre).first()
            if rol:
                rol.permisos = permisos
            else:
                db.session.add(Rol(nombre=nombre, permisos=permisos))

        # Limpiar rol duplicado "Vigilante" si existe
        rol_vig = Rol.query.filter_by(nombre='Vigilante').first()
        rol_pv = Rol.query.filter_by(nombre='Personal de Vigilancia').first()
        if rol_vig and rol_pv:
            for u in Usuario.query.filter_by(rol_id=rol_vig.id).all():
                u.rol_id = rol_pv.id
            db.session.delete(rol_vig)

        # Limpiar roles obsoletos "Coordinador Administrativo" / "Coordinador / Administrativo" si existen
        nombres_coord_obsoletos = ['Coordinador / Administrativo', 'Coordinador Administrativo', 'Coordinador/Administrativo']
        roles_coord = Rol.query.filter(Rol.nombre.in_(nombres_coord_obsoletos)).all()
        rol_admin = Rol.query.filter_by(nombre='Administrativo').first()
        if roles_coord and rol_admin:
            for rol_coord in roles_coord:
                for u in Usuario.query.filter_by(rol_id=rol_coord.id).all():
                    u.rol_id = rol_admin.id
                db.session.delete(rol_coord)

        # Limpiar texto de area_trabajo y cargo_solicitado en usuarios (roles de Dirección)
        nombres_ed_legacy = ['Equipo Directivo', 'Equipo Directivo "Dirección"', 'Equipo Directivo / Administrativo']
        usuarios_ed = Usuario.query.filter(Usuario.area_trabajo.in_(nombres_ed_legacy)).all()
        for u in usuarios_ed:
            u.area_trabajo = 'Equipo Directivo (Dirección)'

        # Limpiar texto de area_trabajo y cargo_solicitado en usuarios (rol Administrativo)
        usuarios_admin_legacy = Usuario.query.filter(Usuario.area_trabajo.in_(nombres_coord_obsoletos)).all()
        for u in usuarios_admin_legacy:
            u.area_trabajo = 'Administrativo'

        cargos_ed_legacy = Usuario.query.filter(Usuario.cargo_solicitado.in_(nombres_ed_legacy)).all()
        for u in cargos_ed_legacy:
            u.cargo_solicitado = 'Equipo Directivo (Dirección)'

        cargos_admin_legacy = Usuario.query.filter(Usuario.cargo_solicitado.in_(nombres_coord_obsoletos)).all()
        for u in cargos_admin_legacy:
            u.cargo_solicitado = 'Administrativo'

        db.session.commit()
