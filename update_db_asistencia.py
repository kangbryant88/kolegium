"""
Script de migración para crear las tablas de Asistencia por Estudiante y Alertas de Defensoría.
Ejecutar una sola vez: python update_db_asistencia.py
"""
import os
import sys

# Fix encoding for Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app
from models import db

def migrar():
    with app.app_context():
        # Verificar si las tablas ya existen
        inspector = db.inspect(db.engine)
        tablas_existentes = inspector.get_table_names()
        
        tablas_nuevas = []
        
        if 'alerta_defensoria' not in tablas_existentes:
            tablas_nuevas.append('alerta_defensoria')
        
        # Verificar si asistencia_estudiante necesita actualización (nuevas columnas)
        if 'asistencia_estudiante' in tablas_existentes:
            columnas = [col['name'] for col in inspector.get_columns('asistencia_estudiante')]
            if 'estatus' not in columnas:
                print("[INFO] Actualizando tabla asistencia_estudiante...")
                try:
                    if 'grado_id' not in columnas:
                        db.session.execute(db.text("ALTER TABLE asistencia_estudiante ADD COLUMN grado_id INTEGER REFERENCES grado(id)"))
                        print("  [OK] Columna grado_id añadida")
                    if 'estatus' not in columnas:
                        db.session.execute(db.text("ALTER TABLE asistencia_estudiante ADD COLUMN estatus VARCHAR(20) DEFAULT 'Presente'"))
                        print("  [OK] Columna estatus añadida")
                    if 'asistio' in columnas:
                        db.session.execute(db.text("UPDATE asistencia_estudiante SET estatus = CASE WHEN asistio = 1 THEN 'Presente' ELSE 'Ausente' END WHERE estatus IS NULL"))
                        print("  [OK] Datos migrados de 'asistio' a 'estatus'")
                    db.session.commit()
                except Exception as e:
                    print(f"  [WARN] Error al actualizar columnas (puede que ya existan): {e}")
                    db.session.rollback()
        else:
            tablas_nuevas.append('asistencia_estudiante')
        
        if tablas_nuevas:
            print(f"[INFO] Creando tablas nuevas: {', '.join(tablas_nuevas)}")
        
        db.create_all()
        print("[OK] Base de datos verificada. Todas las tablas estan al dia.")
        
        # Verificación final
        tablas_finales = db.inspect(db.engine).get_table_names()
        print(f"\n[INFO] Tablas en la base de datos ({len(tablas_finales)}):")
        for t in sorted(tablas_finales):
            print(f"   - {t}")

if __name__ == '__main__':
    migrar()
