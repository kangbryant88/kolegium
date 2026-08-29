"""API móvil del módulo 'Mi Aula'.

Ruta paralela a academico.mi_aula (app/web/academico.py): ejecuta las mismas
consultas (matrícula, asistencia promedio, incidencias) pero responde JSON en
lugar de render_template. Protegida con el token de acceso de la app.
"""
from datetime import date, timedelta

from flask import jsonify, request

from app.api import api_bp
from app.api.security import token_required
from app.models import (
    AsistenciaEstudiante,
    Estudiante,
    Grado,
    Incidencia,
    ProyectoAprendizaje,
    Usuario,
    db,
)

# Mismas categorías que usa la vista web.
CATEGORIAS_INCIDENCIA = ('Conducta', 'Académico', 'Salud', 'Familiar')


@api_bp.route('/mi_aula/<int:grado_id>', methods=['GET'])
@token_required
def mi_aula(grado_id):
    grado = Grado.query.get(grado_id)
    if grado is None:
        return jsonify(error='Grado no encontrado.'), 404

    # --- Matrícula activa del salón ---
    estudiantes = (
        Estudiante.query
        .filter_by(grado_id=grado.id, estatus='Activo')
        .order_by(Estudiante.nombre_completo.asc())
        .all()
    )
    estudiante_ids = [e.id for e in estudiantes]

    # --- Asistencia histórica por estudiante (% de días 'Presente') ---
    asistencia_por_estudiante = {}
    for est in estudiantes:
        total_dias = AsistenciaEstudiante.query.filter_by(
            estudiante_id=est.id
        ).count() or 0
        asistencias = AsistenciaEstudiante.query.filter_by(
            estudiante_id=est.id, estatus='Presente'
        ).count() or 0
        porcentaje = (asistencias / total_dias * 100) if total_dias > 0 else 0.0
        asistencia_por_estudiante[est.id] = round(porcentaje, 1)

    # --- Incidencias del mes en curso, por categoría ---
    datos_incidencias = {categoria: 0 for categoria in CATEGORIAS_INCIDENCIA}

    # --- Asistencia promedio del salón (últimos 30 días) ---
    asistencia_promedio_salon = 100.0

    if estudiantes:
        mes_actual = date.today().month

        incidencias_salon = Incidencia.query.filter(
            Incidencia.estudiante_id.in_(estudiante_ids)
        ).all()
        for inc in incidencias_salon:
            if (
                inc.fecha
                and inc.fecha.month == mes_actual
                and inc.categoria in datos_incidencias
            ):
                datos_incidencias[inc.categoria] += 1

        treinta_dias_atras = date.today() - timedelta(days=30)
        asist_records = AsistenciaEstudiante.query.filter(
            AsistenciaEstudiante.estudiante_id.in_(estudiante_ids),
            AsistenciaEstudiante.fecha >= treinta_dias_atras,
        ).all()
        if asist_records:
            total_dias_salon = len(asist_records)
            asistencias_positivas = sum(
                1 for a in asist_records if a.estatus == 'Presente'
            )
            asistencia_promedio_salon = (
                round((asistencias_positivas / total_dias_salon) * 100, 1)
                if total_dias_salon > 0
                else 0.0
            )

    total_varones = sum(1 for e in estudiantes if e.genero == 'Masculino')
    total_hembras = sum(1 for e in estudiantes if e.genero == 'Femenino')
    docentes = [d.nombre_completo for d in grado.docentes]
    docente_titular = ', '.join(docentes) if docentes else 'Docente no asignado'

    # --- Proyectos de Aprendizaje oficiales del salón ---
    proyectos_oficiales = (
        ProyectoAprendizaje.query
        .filter_by(grado_id=grado.id)
        .order_by(ProyectoAprendizaje.fecha_creacion.desc())
        .all()
    )

    return jsonify(
        grado={
            'id': grado.id,
            'nombre': grado.nombre,
            'docentes': docentes,
            'docente_titular': docente_titular,
        },
        resumen={
            'total_matricula': len(estudiantes),
            'total_varones': total_varones,
            'total_hembras': total_hembras,
            'asistencia_promedio_salon': asistencia_promedio_salon,
            'total_incidencias_mes': sum(datos_incidencias.values()),
        },
        incidencias_mes=datos_incidencias,
        estudiantes=[
            {
                'id': e.id,
                'nombre_completo': e.nombre_completo,
                'nombres': e.nombres,
                'apellidos': e.apellidos,
                'cedula_escolar': e.cedula_escolar,
                'genero': e.genero,
                'estatus': e.estatus,
                'asistencia_porcentaje': asistencia_por_estudiante.get(e.id, 0.0),
            }
            for e in estudiantes
        ],
        proyectos_aprendizaje=[
            {
                'id': p.id,
                'tema': p.tema,
                'momento_pedagogico': p.momento_pedagogico,
                'fecha_inicio': p.fecha_inicio.isoformat() if p.fecha_inicio else None,
                'fecha_culminacion': (
                    p.fecha_culminacion.isoformat() if p.fecha_culminacion else None
                ),
                'fecha_entrega': p.fecha_entrega.isoformat() if p.fecha_entrega else None,
                'entregado': bool(p.entregado),
            }
            for p in proyectos_oficiales
        ],
    ), 200


@api_bp.route('/mi_aula/<int:grado_id>/asistencia', methods=['POST'])
@token_required
def guardar_asistencia(grado_id):
    """Pase de asistencia del día.

    Body: ``{"presentes": [id, id, ...]}``. Misma lógica de upsert por fecha
    que ``academico.guardar_asistencia_aula``: los estudiantes del grado que
    no estén en la lista quedan como 'Ausente'.
    """
    grado = Grado.query.get(grado_id)
    if grado is None:
        return jsonify(error='Grado no encontrado.'), 404

    data = request.get_json(silent=True) or {}
    presentes = data.get('presentes')
    if not isinstance(presentes, list):
        return jsonify(error='Se esperaba {"presentes": [ids...]}.'), 400
    try:
        presentes_ids = {int(x) for x in presentes}
    except (TypeError, ValueError):
        return jsonify(error='"presentes" debe contener IDs numéricos.'), 400

    estudiantes = Estudiante.query.filter_by(grado_id=grado.id).all()
    fecha_hoy = date.today()

    for est in estudiantes:
        estatus = 'Presente' if est.id in presentes_ids else 'Ausente'
        registro = AsistenciaEstudiante.query.filter_by(
            estudiante_id=est.id, fecha=fecha_hoy
        ).first()
        if registro:
            registro.estatus = estatus
            registro.grado_id = grado.id
        else:
            db.session.add(AsistenciaEstudiante(
                fecha=fecha_hoy,
                estatus=estatus,
                estudiante_id=est.id,
                grado_id=grado.id,
            ))

    db.session.commit()
    return jsonify(
        ok=True,
        fecha=fecha_hoy.isoformat(),
        total=len(estudiantes),
        presentes=sum(1 for e in estudiantes if e.id in presentes_ids),
    ), 200


@api_bp.route('/mi_aula/incidencia', methods=['POST'])
@token_required
def agregar_incidencia():
    """Registra una incidencia ('Añadir Nota').

    Body: ``{"estudiante_id", "grado_id", "categoria", "descripcion"}``.
    Misma lógica que ``academico.agregar_incidencia``.
    """
    data = request.get_json(silent=True) or {}
    estudiante_id = data.get('estudiante_id')
    grado_id = data.get('grado_id')
    categoria = (data.get('categoria') or '').strip()
    descripcion = (data.get('descripcion') or '').strip()

    if not estudiante_id or not categoria or not descripcion:
        return jsonify(
            error='Faltan campos: estudiante_id, categoria y descripcion.'
        ), 400
    if categoria not in CATEGORIAS_INCIDENCIA:
        return jsonify(
            error='Categoría inválida. Use una de: '
                  + ', '.join(CATEGORIAS_INCIDENCIA) + '.'
        ), 400

    estudiante = Estudiante.query.get(estudiante_id)
    if estudiante is None:
        return jsonify(error='Estudiante no encontrado.'), 404

    # La API aún no lleva identidad de usuario: la incidencia se atribuye al
    # docente titular del grado (o al primer usuario del sistema como último
    # recurso), ya que Incidencia.usuario_id es obligatorio.
    autor_id = None
    grado = Grado.query.get(grado_id) if grado_id else None
    if grado and grado.docentes:
        autor_id = grado.docentes[0].id
    if autor_id is None:
        primer_usuario = Usuario.query.order_by(Usuario.id.asc()).first()
        autor_id = primer_usuario.id if primer_usuario else None
    if autor_id is None:
        return jsonify(error='No hay un usuario al cual atribuir la nota.'), 500

    incidencia = Incidencia(
        categoria=categoria,
        descripcion=descripcion,
        estudiante_id=estudiante.id,
        usuario_id=autor_id,
    )
    db.session.add(incidencia)
    db.session.commit()
    return jsonify(ok=True, id=incidencia.id), 200
