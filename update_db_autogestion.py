import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'roboclass.db')

def update_db():
    print(f"Connecting to {db_path}...")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    try:
        cursor.execute("ALTER TABLE usuario ADD COLUMN curriculum_path VARCHAR(250)")
        print("Added curriculum_path.")
    except sqlite3.OperationalError as e:
        print(f"curriculum_path: {e}")

    try:
        cursor.execute("ALTER TABLE usuario ADD COLUMN rif_path VARCHAR(250)")
        print("Added rif_path.")
    except sqlite3.OperationalError as e:
        print(f"rif_path: {e}")
        
    try:
        cursor.execute("ALTER TABLE usuario ADD COLUMN cedula_path VARCHAR(250)")
        print("Added cedula_path.")
    except sqlite3.OperationalError as e:
        print(f"cedula_path: {e}")
        
    try:
        cursor.execute("ALTER TABLE usuario ADD COLUMN foto_perfil_path VARCHAR(250)")
        print("Added foto_perfil_path.")
    except sqlite3.OperationalError as e:
        print(f"foto_perfil_path: {e}")

    conn.commit()
    conn.close()
    print("Database updated successfully!")

if __name__ == '__main__':
    update_db()
