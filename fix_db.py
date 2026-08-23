import sqlite3
import glob

archivos_db = glob.glob('*.db') + glob.glob('instance/*.db')

for db_path in archivos_db:
    print(f"🔧 Revisando: {db_path}")
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("ALTER TABLE usuario ADD COLUMN activo BOOLEAN DEFAULT 1;")
        conn.commit()
        print(f"✅ ¡Éxito! Columna creada en {db_path}")
    except sqlite3.OperationalError as e:
        print(f"⚠️ Omitido (quizás ya la tenía): {e}")
    finally:
        conn.close()