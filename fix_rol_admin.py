import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app
from models import db, Rol, Usuario

def fix_rol_admin():
    with app.app_context():
        # Renombrar Rol
        rol = Rol.query.filter_by(nombre='Coordinador / Administrativo').first()
        if rol:
            print("[INFO] Renombrando rol 'Coordinador / Administrativo' a 'Administrativo'")
            rol.nombre = 'Administrativo'
        
        # Renombrar área de trabajo de usuarios
        usuarios = Usuario.query.filter_by(area_trabajo='Coordinador / Administrativo').all()
        for u in usuarios:
            print(f"[INFO] Renombrando area_trabajo de usuario '{u.nombre_completo}' a 'Administrativo'")
            u.area_trabajo = 'Administrativo'
            
        # Limpiar cargo_solicitado por si acaso
        pendientes = Usuario.query.filter_by(cargo_solicitado='Coordinador / Administrativo').all()
        for u in pendientes:
            print(f"[INFO] Renombrando cargo_solicitado de usuario '{u.nombre_completo}' a 'Administrativo'")
            u.cargo_solicitado = 'Administrativo'

        db.session.commit()
        print("[OK] Migración de roles finalizada.")

if __name__ == '__main__':
    fix_rol_admin()
