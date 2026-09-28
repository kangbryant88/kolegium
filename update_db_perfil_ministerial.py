import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'roboclass.db')

COLUMNAS = [
    # Nivel 1 (Hard Required)
    ("cedula", "VARCHAR(20)"),
    ("nombres", "VARCHAR(100)"),
    ("apellidos", "VARCHAR(100)"),
    ("fecha_nacimiento", "DATE"),
    ("sexo", "VARCHAR(20)"),
    # Nivel 2 (Soft Required)
    ("fecha_ingreso", "DATE"),
    ("cargo", "VARCHAR(100)"),
    ("codigo_rac", "VARCHAR(50)"),
    ("turno", "VARCHAR(20)"),
]

def update_db():
    print(f"Connecting to {db_path}...")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for nombre, tipo in COLUMNAS:
        try:
            cursor.execute(f"ALTER TABLE usuario ADD COLUMN {nombre} {tipo}")
            print(f"Added {nombre}.")
        except sqlite3.OperationalError as e:
            print(f"{nombre}: {e}")

    conn.commit()
    conn.close()
    print("Database updated successfully!")

if __name__ == '__main__':
    update_db()
