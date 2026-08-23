import os
import sys

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app
from models import db

def migrar_activo():
    with app.app_context():
        inspector = db.inspect(db.engine)
        columnas = [col['name'] for col in inspector.get_columns('usuario')]

        if 'activo' not in columnas:
            print("[INFO] Añadiendo columna 'activo' a la tabla 'usuario'...")
            try:
                db.session.execute(db.text("ALTER TABLE usuario ADD COLUMN activo BOOLEAN DEFAULT 1"))
                db.session.commit()
                print("  [OK] Columna añadida correctamente.")
            except Exception as e:
                print(f"  [ERROR] No se pudo añadir la columna: {e}")
                db.session.rollback()
        else:
            print("[INFO] La columna 'activo' ya existe en la tabla 'usuario'.")

if __name__ == '__main__':
    migrar_activo()
