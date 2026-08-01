import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app
from models import db, Usuario

def arreglar_cargos():
    with app.app_context():
        usuarios = Usuario.query.all()
        actualizados = 0
        for u in usuarios:
            if u.area_trabajo == 'Equipo Directivo / Administrativo':
                print(f"[MODIFICANDO] {u.nombre_completo}: 'Equipo Directivo / Administrativo' -> 'Equipo Directivo'")
                u.area_trabajo = 'Equipo Directivo'
                actualizados += 1
        
        if actualizados > 0:
            db.session.commit()
            print(f"[OK] Se actualizaron {actualizados} usuarios.")
        else:
            print("[INFO] No se encontraron usuarios con el cargo antiguo.")

if __name__ == '__main__':
    arreglar_cargos()
