import sqlite3
import os

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'roboclass.db')

# Ficha ministerial completa (RAC Nominal + RAC Beneficios)
COLUMNAS = [
    # Personales
    ("nacionalidad", "VARCHAR(20)"),
    ("lugar_nacimiento", "VARCHAR(150)"),
    ("estado_civil", "VARCHAR(30)"),
    ("nivel_instruccion", "VARCHAR(100)"),
    ("profesion", "VARCHAR(150)"),
    ("telefono_habitacion", "VARCHAR(20)"),
    ("telefono_oficina", "VARCHAR(20)"),
    # Laborales (RAC)
    ("tipo_personal", "VARCHAR(50)"),
    ("horas_academicas", "FLOAT"),
    ("horas_adm", "FLOAT"),
    ("grado_imparte", "VARCHAR(50)"),
    ("seccion_imparte", "VARCHAR(20)"),
    ("especialidad", "VARCHAR(150)"),
    ("situacion_trabajador", "VARCHAR(50)"),
    ("observacion", "TEXT"),
    # Beneficios / Salud
    ("talla_camisa", "VARCHAR(10)"),
    ("talla_pantalon", "VARCHAR(10)"),
    ("talla_zapato", "VARCHAR(10)"),
    ("actividad_deportiva", "VARCHAR(150)"),
    ("actividad_cultural", "VARCHAR(150)"),
    ("tipo_vivienda", "VARCHAR(50)"),
    ("condicion_vivienda", "VARCHAR(50)"),
    ("tipo_material", "VARCHAR(50)"),
    ("tipo_enfermedad", "VARCHAR(200)"),
    ("medicamento", "VARCHAR(200)"),
    ("posee_discapacidad", "VARCHAR(100)"),
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
