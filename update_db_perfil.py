import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'roboclass.db')

def update_db():
    print(f"Connecting to {db_path}...")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    try:
        cursor.execute("ALTER TABLE usuario ADD COLUMN cargo_solicitado VARCHAR(100)")
        print("Added cargo_solicitado.")
    except sqlite3.OperationalError as e:
        print(f"cargo_solicitado: {e}")

    conn.commit()
    conn.close()
    print("Database updated successfully!")

if __name__ == '__main__':
    update_db()
