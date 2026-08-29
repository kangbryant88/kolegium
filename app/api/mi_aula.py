"""API móvil del módulo 'Mi Aula'.

Ruta paralela a academico.mi_aula (app/web/academico.py): ejecuta las mismas
consultas (matrícula, asistencia promedio, incidencias) pero responde JSON en
lugar de render_template. Protegida con el token de acceso de la app.
"""
from datetime import date, timedelta

from flask import jsonify

from app.api import api_bp
from app.api.security import token_required
from app.models import (
    AsistenciaEstudiante,
    Estudiante,
    Grado,
    Incidencia,
    ProyectoAprendizaje,
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
