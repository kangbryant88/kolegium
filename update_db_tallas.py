import sqlite3

def upgrade():
    conn = sqlite3.connect('roboclass.db')
    cursor = conn.cursor()
    try:
        cursor.execute('ALTER TABLE estudiante ADD COLUMN talla_camisa VARCHAR(50)')
        print("Columna 'talla_camisa' añadida exitosamente.")
    except sqlite3.OperationalError:
        print("La columna 'talla_camisa' ya existe en la base de datos.")
        
    try:
        cursor.execute('ALTER TABLE estudiante ADD COLUMN talla_pantalon VARCHAR(50)')
        print("Columna 'talla_pantalon' añadida exitosamente.")
    except sqlite3.OperationalError:
        print("La columna 'talla_pantalon' ya existe en la base de datos.")
        
    conn.commit()
    conn.close()
    print("Base de datos actualizada.")

if __name__ == '__main__':
    upgrade()
