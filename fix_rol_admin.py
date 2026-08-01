import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app
from models import db, Rol, Usuario

def fix_rol_admin():
    with app.app_context():
        # Revisar si ya existe el rol 'Administrativo'
        rol_existente = Rol.query.filter_by(nombre='Administrativo').first()
        rol_viejo = Rol.query.filter_by(nombre='Coordinador / Administrativo').first()

        if rol_existente and rol_viejo:
            print("[INFO] El rol 'Administrativo' ya existe. Moviendo usuarios al rol existente y borrando el viejo...")
            usuarios = Usuario.query.filter_by(rol_id=rol_viejo.id).all()
            for u in usuarios:
                u.rol_id = rol_existente.id
            db.session.delete(rol_viejo)
            
        elif rol_viejo and not rol_existente:
            print("[INFO] Renombrando rol 'Coordinador / Administrativo' a 'Administrativo'")
            rol_viejo.nombre = 'Administrativo'
        else:
            print("[INFO] El rol 'Coordinador / Administrativo' no se encontró o ya fue procesado.")

        # Renombrar área de trabajo de usuarios
        usuarios = Usuario.query.filter_by(area_trabajo='Coordinador / Administrativo').all()
        for u in usuarios:
            print(f"[INFO] Renombrando area_trabajo de usuario '{u.nombre_completo}' a 'Administrativo'")
            u.area_trabajo = 'Administrativo'
            
        # Limpiar cargo_solicitado
        pendientes = Usuario.query.filter_by(cargo_solicitado='Coordinador / Administrativo').all()
        for u in pendientes:
            print(f"[INFO] Renombrando cargo_solicitado de usuario '{u.nombre_completo}' a 'Administrativo'")
            u.cargo_solicitado = 'Administrativo'

        db.session.commit()
        print("[OK] Migración de roles finalizada.")

if __name__ == '__main__':
    fix_rol_admin()
