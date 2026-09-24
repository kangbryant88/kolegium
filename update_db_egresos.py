import sqlite3
import os

db_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'roboclass.db')
print("DB Path:", db_path)
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS egreso (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    estudiante_id INTEGER NOT NULL,
    nombre_completo VARCHAR(150) NOT NULL,
    cedula_escolar VARCHAR(30) NOT NULL,
    grado_nombre VARCHAR(50),
    motivo VARCHAR(200) DEFAULT 'Retiro / Egreso',
    fecha_egreso DATETIME,
    usuario_id INTEGER,
    FOREIGN KEY (estudiante_id) REFERENCES estudiante (id),
    FOREIGN KEY (usuario_id) REFERENCES usuario (id)
)
""")
print("Tabla 'egreso' creada o ya existente")

conn.commit()
conn.close()
print("Done")
