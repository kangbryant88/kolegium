import sqlite3

NUEVAS_COLUMNAS_ESTUDIANTE = [
    ('nombres', 'VARCHAR(100)'),
    ('apellidos', 'VARCHAR(100)'),
    ('municipio', 'VARCHAR(100)'),
    ('parroquia', 'VARCHAR(100)'),
    ('edad', 'INTEGER'),
    ('nacionalidad', 'VARCHAR(50)'),
    ('direccion_alumno', 'TEXT'),
    ('telefono_habitacion', 'VARCHAR(20)'),
    ('posee_canaima', 'BOOLEAN DEFAULT 0'),
    ('posee_enfermedad', 'BOOLEAN DEFAULT 0'),
    ('enfermedad_detalle', 'VARCHAR(200)'),
    ('toma_medicamento', 'BOOLEAN DEFAULT 0'),
    ('medicamento_detalle', 'VARCHAR(200)'),
    ('alergico_medicamento', 'BOOLEAN DEFAULT 0'),
    ('alergia_medicamento_detalle', 'VARCHAR(200)'),
    ('madre_nombre', 'VARCHAR(150)'),
    ('madre_ci', 'VARCHAR(20)'),
    ('madre_ocupacion', 'VARCHAR(100)'),
    ('madre_telefono', 'VARCHAR(20)'),
    ('madre_direccion', 'TEXT'),
    ('padre_nombre', 'VARCHAR(150)'),
    ('padre_ci', 'VARCHAR(20)'),
    ('padre_ocupacion', 'VARCHAR(100)'),
    ('padre_telefono', 'VARCHAR(20)'),
    ('padre_direccion', 'TEXT'),
    ('telefono_familiar_extra', 'VARCHAR(20)'),
    ('autorizacion_odontologica', 'BOOLEAN DEFAULT 0'),
]

NUEVAS_COLUMNAS_REPRESENTANTE = [
    ('ocupacion', 'VARCHAR(100)'),
    ('lugar_direccion_trabajo', 'TEXT'),
    ('banco_nombre', 'VARCHAR(100)'),
    ('banco_cuenta_numero', 'VARCHAR(30)'),
    ('banco_cuenta_tipo', 'VARCHAR(20)'),
    ('banco_titular_nombre', 'VARCHAR(150)'),
    ('banco_titular_ci', 'VARCHAR(20)'),
]


def upgrade():
    conn = sqlite3.connect('roboclass.db')
    cursor = conn.cursor()

    for tabla, columnas in (('estudiante', NUEVAS_COLUMNAS_ESTUDIANTE),
                             ('representante', NUEVAS_COLUMNAS_REPRESENTANTE)):
        for nombre, tipo in columnas:
            try:
                cursor.execute(f'ALTER TABLE {tabla} ADD COLUMN {nombre} {tipo}')
                print(f"Columna '{nombre}' añadida a '{tabla}'.")
            except sqlite3.OperationalError:
                print(f"La columna '{nombre}' ya existe en '{tabla}'.")

    conn.commit()
    conn.close()
    print("Base de datos actualizada (Ficha de Admisión).")


if __name__ == '__main__':
    upgrade()
