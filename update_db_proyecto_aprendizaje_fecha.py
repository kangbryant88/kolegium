import os
import sys

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app
from models import db

def migrar_fecha_creacion():
    with app.app_context():
        inspector = db.inspect(db.engine)
        columnas = [col['name'] for col in inspector.get_columns('proyecto_aprendizaje')]

        if 'fecha_creacion' not in columnas:
            print("[INFO] Añadiendo columna 'fecha_creacion' a la tabla 'proyecto_aprendizaje'...")
            try:
                db.session.execute(db.text("ALTER TABLE proyecto_aprendizaje ADD COLUMN fecha_creacion DATETIME"))
                db.session.execute(db.text("UPDATE proyecto_aprendizaje SET fecha_creacion = CURRENT_TIMESTAMP WHERE fecha_creacion IS NULL"))
                db.session.commit()
                print("  [OK] Columna añadida correctamente.")
            except Exception as e:
                print(f"  [ERROR] No se pudo añadir la columna: {e}")
                db.session.rollback()
        else:
            print("[INFO] La columna 'fecha_creacion' ya existe en la tabla 'proyecto_aprendizaje'.")

if __name__ == '__main__':
    migrar_fecha_creacion()
