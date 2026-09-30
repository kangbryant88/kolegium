"""
Limpia las fechas absurdas que dejó el importador antes del fix 80e3a64: un
año escrito como número (2022) se leía como serial de Excel y se guardaba en
1905. Pone en NULL fecha_ingreso y fecha_nacimiento para que la próxima
importación (que solo rellena campos vacíos) las cargue bien.

Si la fecha correcta no viene en el Excel, fecha_nacimiento vacía activa el
Muro de Contención y el usuario deberá cargarla él mismo: es intencional.

Uso:
    python fix_fechas_excel.py            -> solo muestra lo que cambiaría
    python fix_fechas_excel.py --aplicar  -> respalda la BD y aplica
"""
import os
import shutil
import sqlite3
import sys
from datetime import date, datetime

basedir = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(basedir, 'roboclass.db')

# Campo -> fecha mínima razonable (anterior = error). Posterior a hoy también es error.
# Nadie en nómina activa ingresó antes de 1940 ni nació antes de 1920.
CAMPOS = {
    'fecha_ingreso': '1940-01-01',
    'fecha_nacimiento': '1920-01-01',
}


def _condicion(campo):
    return f"{campo} IS NOT NULL AND ({campo} < ? OR {campo} > ?)"


def main(aplicar):
    hoy = date.today().isoformat()
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    total = 0
    for campo, minima in CAMPOS.items():
        cur.execute(f"SELECT id, nombre_completo, cedula, {campo} FROM usuario "
                    f"WHERE {_condicion(campo)} ORDER BY id", (minima, hoy))
        filas = cur.fetchall()
        total += len(filas)
        print(f"{campo} absurda (< {minima} o futura): {len(filas)}")
        for fila in filas:
            print("   #%s  %-35s %-12s %s" % fila)
    conn.close()

    if not aplicar or not total:
        print("\nNo se modificó nada." + (" Ejecuta con --aplicar para limpiar." if total else ""))
        return

    respaldo = f"{db_path}.bak_{datetime.now():%Y%m%d_%H%M%S}"
    shutil.copy2(db_path, respaldo)
    print(f"\nRespaldo: {respaldo}")

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    for campo, minima in CAMPOS.items():
        cur.execute(f"UPDATE usuario SET {campo} = NULL WHERE {_condicion(campo)}", (minima, hoy))
        print(f"{campo}: {cur.rowcount} filas puestas en NULL")
    conn.commit()
    conn.close()


if __name__ == '__main__':
    main('--aplicar' in sys.argv)
