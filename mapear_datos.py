"""
mapear_datos.py

Diagnostico:
El Representante (nombre/CI/telefono/direccion) siempre vivio en la tabla
`representante` y se lee igual antes y despues de la Ficha de Admision
ampliada -- por eso su tarjeta ("Representante y Banco") NUNCA deja de
mostrar datos para un estudiante con `representante_id` valido.

Lo que SI quedo vacio para los estudiantes inscritos antes de esta feature
son las columnas nuevas y dedicadas `madre_*` / `padre_*` de `Estudiante`:
antes, un solo Representante (con su campo `parentesco` = 'Madre' o 'Padre')
concentraba esa informacion; ahora el formulario separa explicitamente a
Madre y Padre. La migracion (`update_db_admision.py`) fue un ALTER TABLE
aditivo -- no pudo inventar datos en columnas que no existian, asi que
quedaron en NULL para todo alumno previo.

Este script recorre los estudiantes existentes y, cuando el Representante
vinculado tiene parentesco 'Madre' o 'Padre', copia sus datos hacia las
columnas madre_* / padre_* correspondientes -- SOLO si esas columnas siguen
vacias (nunca sobreescribe algo ya cargado a mano). Es seguro ejecutarlo
varias veces (idempotente).

Tambien reporta (sin poder corregirlos, porque no existe ninguna fuente de
la cual copiar) los estudiantes cuyo representante_id es NULL o apunta a un
Representante inexistente.

Uso:
    python mapear_datos.py
"""

from app import app
from models import db, Estudiante


def upgrade():
    estudiantes = Estudiante.query.all()

    con_representante = 0
    actualizados_madre = 0
    actualizados_padre = 0
    sin_representante = []

    for est in estudiantes:
        rep = est.representante_info

        if not rep:
            sin_representante.append((est.id, est.nombre_completo, est.cedula_escolar))
            continue

        con_representante += 1
        parentesco = (rep.parentesco or '').strip().lower()

        if parentesco == 'madre' and not est.madre_nombre:
            est.madre_nombre = rep.nombre_completo
            est.madre_ci = est.madre_ci or rep.cedula
            est.madre_telefono = est.madre_telefono or rep.telefono
            est.madre_ocupacion = est.madre_ocupacion or rep.ocupacion
            est.madre_direccion = est.madre_direccion or rep.direccion_habitacion or rep.direccion_completa
            actualizados_madre += 1

        elif parentesco == 'padre' and not est.padre_nombre:
            est.padre_nombre = rep.nombre_completo
            est.padre_ci = est.padre_ci or rep.cedula
            est.padre_telefono = est.padre_telefono or rep.telefono
            est.padre_ocupacion = est.padre_ocupacion or rep.ocupacion
            est.padre_direccion = est.padre_direccion or rep.direccion_habitacion or rep.direccion_completa
            actualizados_padre += 1

    db.session.commit()

    print(f"Estudiantes revisados: {len(estudiantes)}")
    print(f"Con representante vinculado: {con_representante}")
    print(f"Backfill aplicado a madre_*: {actualizados_madre}")
    print(f"Backfill aplicado a padre_*: {actualizados_padre}")

    if sin_representante:
        print()
        print(f"ATENCION: {len(sin_representante)} estudiante(s) SIN representante vinculado "
              f"(representante_id nulo o huerfano). No hay ninguna fuente de la cual migrar "
              f"sus datos; requieren carga manual desde 'Editar Datos':")
        for eid, nombre, cesc in sin_representante:
            print(f"  - ID {eid} | {nombre} | CESC {cesc}")


if __name__ == '__main__':
    with app.app_context():
        upgrade()
