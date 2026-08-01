import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app
from models import db, Usuario

def fix_departamentos():
    with app.app_context():
        usuarios = Usuario.query.filter_by(departamento_asignado='Pedagógica').all()
        actualizados = 0
        for u in usuarios:
            print(f"[INFO] Renombrando departamento de '{u.nombre_completo}' a 'Coord. Pedagógica'")
            u.departamento_asignado = 'Coord. Pedagógica'
            actualizados += 1
            
        if actualizados > 0:
            db.session.commit()
            print(f"[OK] Se actualizaron {actualizados} usuarios.")
        else:
            print("[INFO] No se encontraron usuarios con el departamento 'Pedagógica'.")

if __name__ == '__main__':
    fix_departamentos()
