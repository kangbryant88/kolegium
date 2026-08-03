import sqlite3
import pandas as pd

def check_db():
    conn = sqlite3.connect('roboclass.db')
    print("--- ROLES ---")
    roles = pd.read_sql_query("SELECT * FROM rol", conn)
    print(roles)
    
    print("\n--- USUARIOS ---")
    usuarios = pd.read_sql_query("SELECT id, username, nombre_completo, area_trabajo, rol_id, departamento_asignado FROM usuario", conn)
    print(usuarios)
    conn.close()

if __name__ == "__main__":
    check_db()
