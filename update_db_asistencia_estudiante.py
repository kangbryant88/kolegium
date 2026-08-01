import sqlite3
import os
import sys

def update_db():
    # Asegurarnos de usar EXACTAMENTE la misma base de datos que usa la app
    # (la que está en la raíz de la carpeta, no en 'instance')
    basedir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(basedir, 'roboclass.db')
    
    print(f"Conectando a la base de datos oficial en: {db_path}")
    
    if not os.path.exists(db_path):
        print(f"Error: No se encontró el archivo de base de datos en {db_path}")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        cursor.execute("ALTER TABLE asistencia_estudiante ADD COLUMN estatus VARCHAR(20) NOT NULL DEFAULT 'Presente'")
        print("ÉXITO: Columna 'estatus' agregada con éxito a asistencia_estudiante.")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("INFO: La columna 'estatus' ya existe en asistencia_estudiante.")
        else:
            print(f"ERROR al agregar columna 'estatus': {e}")
            
    try:
        cursor.execute("ALTER TABLE asistencia_estudiante ADD COLUMN grado_id INTEGER DEFAULT NULL")
        print("ÉXITO: Columna 'grado_id' agregada con éxito a asistencia_estudiante.")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("INFO: La columna 'grado_id' ya existe en asistencia_estudiante.")
        else:
            print(f"ERROR al agregar columna 'grado_id': {e}")

    conn.commit()
    conn.close()
    print("Actualización de la base de datos completada.")

if __name__ == '__main__':
    update_db()
