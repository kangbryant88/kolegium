import sqlite3
import os

def update_db():
    db_path = os.path.join(os.path.dirname(__file__), 'instance', 'roboclass.db')
    if not os.path.exists(db_path):
        db_path = os.path.join(os.path.dirname(__file__), 'roboclass.db')
        
    print(f"Connecting to database at: {db_path}")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Añadir la columna 'estatus' a asistencia_estudiante
        cursor.execute("ALTER TABLE asistencia_estudiante ADD COLUMN estatus VARCHAR(20) NOT NULL DEFAULT 'Presente'")
        print("Columna 'estatus' agregada con éxito a asistencia_estudiante.")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("La columna 'estatus' ya existe en asistencia_estudiante.")
        else:
            print(f"Error al agregar columna 'estatus': {e}")
            
    try:
        # Ya que estamos aquí, verifiquemos si falta el grado_id también, por si acaso
        cursor.execute("ALTER TABLE asistencia_estudiante ADD COLUMN grado_id INTEGER DEFAULT NULL")
        print("Columna 'grado_id' agregada con éxito a asistencia_estudiante.")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("La columna 'grado_id' ya existe en asistencia_estudiante.")
        else:
            print(f"Error al agregar columna 'grado_id': {e}")

    conn.commit()
    conn.close()
    print("Actualización de la base de datos completada.")

if __name__ == '__main__':
    update_db()
