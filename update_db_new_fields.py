import sqlite3
import os

db_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'roboclass.db')
print("DB Path:", db_path)
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

def add_column(table, column, datatype):
    try:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {datatype}")
        print(f"Added column {column} to {table}")
    except sqlite3.OperationalError as e:
        print(f"Column {column} might already exist or error: {e}")

add_column('estudiante', 'literal_escolar', 'VARCHAR(2)')
add_column('estudiante', 'plantel_procedencia', 'VARCHAR(150)')
add_column('estudiante', 'email_estudiante', 'VARCHAR(120)')

add_column('representante', 'email_representante', 'VARCHAR(120)')
add_column('representante', 'direccion_completa', 'TEXT')

conn.commit()
conn.close()
print("Done")
