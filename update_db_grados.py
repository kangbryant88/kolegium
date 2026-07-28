import sqlite3

def migrate_grados():
    conn = sqlite3.connect('roboclass.db')
    cursor = conn.cursor()

    # Create the new association table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS grado_docente (
        grado_id INTEGER NOT NULL,
        usuario_id INTEGER NOT NULL,
        PRIMARY KEY (grado_id, usuario_id),
        FOREIGN KEY(grado_id) REFERENCES grado (id),
        FOREIGN KEY(usuario_id) REFERENCES usuario (id)
    )
    ''')

    # Get existing assignments from grado table
    try:
        cursor.execute("SELECT id, usuario_id FROM grado WHERE usuario_id IS NOT NULL")
        grados = cursor.fetchall()
        for grado_id, usuario_id in grados:
            # Insert into the association table if not exists
            cursor.execute("INSERT OR IGNORE INTO grado_docente (grado_id, usuario_id) VALUES (?, ?)", (grado_id, usuario_id))
        
        # We should NOT drop the column from SQLite immediately because ALTER TABLE DROP COLUMN has some constraints 
        # in older SQLite versions. However, we can just leave it there or let Alembic/flask-migrate drop it later.
        # But wait, SQLite 3.35+ supports DROP COLUMN. Let's try it.
        try:
            cursor.execute("ALTER TABLE grado DROP COLUMN usuario_id")
        except sqlite3.OperationalError as e:
            print(f"Note: Could not drop column (might be old SQLite version, which is fine): {e}")

        conn.commit()
        print("Migration successful.")
    except Exception as e:
        print(f"Error during migration: {e}")
    finally:
        conn.close()

if __name__ == '__main__':
    migrate_grados()
