from flask import Blueprint, render_template, request, redirect, url_for, session, flash, send_file, make_response, jsonify
from datetime import datetime, date, timedelta
import io
import pandas as pd
import openpyxl
from docx import Document
from fpdf import FPDF
import urllib.parse
import uuid
from models import db, Bitacora, Grado, Tema, AsistenciaDiaria, AsistenciaPersonal, Estudiante, Representante, Incidencia, AsistenciaEstudiante, AlertaDefensoria, EnlaceTemporal, SolicitudEnlace, SolicitudActualizacion

academico_bp = Blueprint('academico', __name__, url_prefix='/academico')

# ==========================================
# --- 4. PLANIFICADOR DOCENTE ---
# ==========================================

@academico_bp.route('/planificador')
def planificador():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    uid = session['usuario_id']
    clases = Bitacora.query.filter_by(usuario_id=uid).order_by(Bitacora.fecha.desc()).all()
    grados = Grado.query.all(); temas = Tema.query.all()
    return render_template('planificador.html', registros=clases, grados=grados, temas=temas,
                           total=len(clases), completadas=len([c for c in clases if c.estado == 'Completado']),
                           pendientes=len([c for c in clases if c.estado == 'Pendiente']))

@academico_bp.route('/agregar', methods=['POST'])
def agregar():
    hora = request.form.get('hora', '07:00')
    f = datetime.strptime(f"{datetime.now().strftime('%Y-%m-%d')} {hora}", "%Y-%m-%d %H:%M")
    db.session.add(Bitacora(fecha=f, grado=request.form['grado'], actividad=request.form['actividad'], 
                            estado=request.form['estado'], usuario_id=session['usuario_id']))
    db.session.commit(); return redirect(url_for('academico.planificador'))

@academico_bp.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar(id):
    reg = Bitacora.query.get_or_404(id)
    if request.method == 'POST':
        reg.grado, reg.actividad, reg.estado = request.form['grado'], request.form['actividad'], request.form['estado']
        reg.fecha = datetime.strptime(f"{reg.fecha.strftime('%Y-%m-%d')} {request.form['hora']}", "%Y-%m-%d %H:%M")
        db.session.commit(); return redirect(url_for('academico.planificador'))
    return render_template('editar.html', registro=reg, grados=Grado.query.all())

@academico_bp.route('/eliminar/<int:id>')
def eliminar(id):
    db.session.delete(Bitacora.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('academico.planificador'))

@academico_bp.route('/reporte_diario')
def reporte_diario():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    hoy = datetime.now().date()
    regs = [r for r in Bitacora.query.filter_by(usuario_id=session['usuario_id']).all() if r.fecha.date() == hoy]
    return render_template('reporte.html', registros=regs)

@academico_bp.route('/reporte_general')
def reporte_general():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    regs = Bitacora.query.filter_by(usuario_id=session['usuario_id']).order_by(Bitacora.fecha.desc()).all()
    return render_template('reporte.html', registros=regs)

# ==========================================
# --- 5. EXPORTACIONES (EXCEL / WORD) ---
# ==========================================

@academico_bp.route('/exportar_excel')
def exportar_excel():
    regs = Bitacora.query.filter_by(usuario_id=session['usuario_id']).all()
    datos = [{'Fecha': r.fecha.strftime('%d/%m/%Y'), 'Hora': r.fecha.strftime('%I:%M %p'), 
              'Grado': r.grado, 'Actividad': r.actividad, 'Estado': r.estado} for r in regs]
    df = pd.DataFrame(datos); output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Reporte', startrow=4, index=False)
        ws = writer.sheets['Reporte']; fmt = writer.book.add_format({'bold': True, 'font_size': 14})
        ws.write('A1', 'REPORTE DE ACTIVIDADES - EDUPALNNER OS', fmt)
        ws.write('A2', f"Docente: {session.get('nombre_completo')}")
        ws.write('A3', f"Área: {session.get('area_trabajo')}")
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     download_name="Reporte_EduPlanner.xlsx", as_attachment=True)

@academico_bp.route('/exportar_word')
def exportar_word():
    doc = Document(); doc.add_heading('REPORTE OFICIAL DE ACTIVIDADES', 0)
    p = doc.add_paragraph(); p.add_run('Docente: ').bold = True; p.add_run(f"{session.get('nombre_completo')}\n")
    p.add_run('Área: ').bold = True; p.add_run(f"{session.get('area_trabajo')}")
    table = doc.add_table(rows=1, cols=5); table.style = 'Table Grid'
    for i, t in enumerate(['Fecha', 'Hora', 'Grado', 'Actividad', 'Estado']): table.rows[0].cells[i].text = t
    for r in Bitacora.query.filter_by(usuario_id=session['usuario_id']).all():
        row = table.add_row().cells
        row[0].text, row[1].text = r.fecha.strftime('%d/%m/%Y'), r.fecha.strftime('%I:%M %p')
        row[2].text, row[3].text, row[4].text = r.grado, r.actividad, r.estado
    f = io.BytesIO(); doc.save(f); f.seek(0)
    return send_file(f, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
                     download_name="Reporte_EduPlanner.docx", as_attachment=True)

# ==========================================
# ==========================================
# --- 7. CONTROL DE ASISTENCIA ---
# ==========================================

@academico_bp.route('/asistencia', methods=['GET', 'POST'])
def asistencia():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    if request.method == 'POST':
        grado = Grado.query.get(request.form['grado_id'])
        v, h = int(request.form['varones'] or 0), int(request.form['hembras'] or 0)
        t = grado.total_varones + grado.total_hembras
        p = round(((v + h) / t) * 100, 1) if t > 0 else 0
        db.session.add(AsistenciaDiaria(fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(), 
                       grado_seccion=grado.nombre, matricula_total=t, varones=v, hembras=h, asistentes=v+h, 
                       porcentaje=p, usuario_id=session['usuario_id']))
        db.session.commit(); return redirect(url_for('academico.asistencia'))
    if session.get('rol_id') in [1, 2]:
        regs = AsistenciaDiaria.query.order_by(AsistenciaDiaria.fecha.desc()).all()
    else:
        regs = AsistenciaDiaria.query.filter_by(usuario_id=session['usuario_id']).order_by(AsistenciaDiaria.fecha.desc()).all()
        
    from collections import OrderedDict
    
    meses_es = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 
                7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
    dias_es = {0:'Lunes', 1:'Martes', 2:'Miércoles', 3:'Jueves', 4:'Viernes', 5:'Sábado', 6:'Domingo'}

    def formato_fecha_es(fecha_obj):
        return f"{dias_es[fecha_obj.weekday()]}, {fecha_obj.day} de {meses_es[fecha_obj.month]} de {fecha_obj.year}"

    semanas = OrderedDict()
    
    for r in regs:
        dt = r.fecha
        lunes = dt - timedelta(days=dt.weekday())
        viernes = lunes + timedelta(days=4)
        semana_key = f"Semana del {lunes.strftime('%d/%m/%Y')} al {viernes.strftime('%d/%m/%Y')}"
        
        if semana_key not in semanas:
            semanas[semana_key] = OrderedDict()
            
        fecha_str = formato_fecha_es(dt)
        if fecha_str not in semanas[semana_key]:
            semanas[semana_key][fecha_str] = []
            
        semanas[semana_key][fecha_str].append(r)

    return render_template('asistencia.html', registros=regs, semanas=semanas, grados=Grado.query.all(), hoy=datetime.now().strftime('%Y-%m-%d'))

@academico_bp.route('/eliminar_asistencia/<int:id>', methods=['POST'])
def eliminar_asistencia(id):
    db.session.delete(AsistenciaDiaria.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('academico.asistencia'))

@academico_bp.route('/asistencia_personal', methods=['GET', 'POST'])
def asistencia_personal():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    if request.method == 'POST':
        mat_base = int(request.form.get('matricula_base') or 0)
        asistentes = int(request.form.get('asistentes') or 0)
        pct = round((asistentes / mat_base) * 100, 1) if mat_base > 0 else 0.0
        db.session.add(AsistenciaPersonal(
            fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(), 
            categoria=request.form['categoria'], 
            matricula_base=mat_base, 
            asistentes=asistentes, 
            porcentaje=pct, 
            usuario_id=session['usuario_id']
        ))
        db.session.commit()
        return redirect(url_for('academico.asistencia_personal'))
    if session.get('rol_id') in [1, 2]:
        regs = AsistenciaPersonal.query.order_by(AsistenciaPersonal.fecha.desc()).all()
    else:
        regs = AsistenciaPersonal.query.filter_by(usuario_id=session['usuario_id']).order_by(AsistenciaPersonal.fecha.desc()).all()
    
    # Memoria: Obtener el último valor registrado por categoría
    bases_conocidas = {}
    categorias = ['Docentes', 'Administrativos', 'Obreros', 'Vigilantes', 'Cocina']
    for cat in categorias:
        ultimo = AsistenciaPersonal.query.filter_by(usuario_id=session['usuario_id'], categoria=cat)\
            .order_by(AsistenciaPersonal.fecha.desc(), AsistenciaPersonal.id.desc()).first()
        if ultimo:
            bases_conocidas[cat] = ultimo.matricula_base

    return render_template('asistencia_personal.html', registros=regs, hoy=datetime.now().strftime('%Y-%m-%d'), bases=bases_conocidas)

@academico_bp.route('/eliminar_asistencia_personal/<int:id>')
def eliminar_asistencia_personal(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    db.session.delete(AsistenciaPersonal.query.get_or_404(id))
    db.session.commit()
    return redirect(url_for('academico.asistencia_personal'))

@academico_bp.route('/editar_asistencia_personal/<int:id>', methods=['GET', 'POST'])
def editar_asistencia_personal(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    reg = AsistenciaPersonal.query.get_or_404(id)
    if request.method == 'POST':
        mat_base = int(request.form.get('matricula_base') or 0)
        asistentes = int(request.form.get('asistentes') or 0)
        
        reg.fecha = datetime.strptime(request.form['fecha'], '%Y-%m-%d').date()
        reg.categoria = request.form['categoria']
        reg.matricula_base = mat_base
        reg.asistentes = asistentes
        reg.porcentaje = round((asistentes / mat_base) * 100, 1) if mat_base > 0 else 0.0
        
        db.session.commit()
        return redirect(url_for('academico.asistencia_personal'))
    
    return render_template('editar_asistencia_personal.html', registro=reg)

# ==========================================
# --- 8. CONFIGURACIONES GLOBALES ---
# ==========================================

@academico_bp.route('/historial_global')
def historial_global():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    tiene_acceso = (
        session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo (Dirección)'] or 
        (session.get('nombre_rol') == 'Administrativo' and session.get('departamento_asignado') == 'Dirección')
    )
    if not tiene_acceso:
        return redirect(url_for('index'))
    
    fecha_str = request.args.get('fecha')
    if fecha_str:
        try:
            fecha_busqueda = datetime.strptime(fecha_str, '%Y-%m-%d').date()
        except ValueError:
            fecha_busqueda = date.today() - timedelta(days=1)
    else:
        fecha_busqueda = date.today() - timedelta(days=1)
        
    regs_estudiantes = AsistenciaDiaria.query.filter_by(fecha=fecha_busqueda).all()
    regs_personal = AsistenciaPersonal.query.filter_by(fecha=fecha_busqueda).all()
    
    t_estudiantes = sum(r.asistentes for r in regs_estudiantes)
    t_personal = sum(r.asistentes for r in regs_personal)
    
    return render_template('historial.html', 
                           fecha_busqueda=fecha_busqueda.strftime('%Y-%m-%d'),
                           regs_estudiantes=regs_estudiantes,
                           regs_personal=regs_personal,
                           t_estudiantes=t_estudiantes,
                           t_personal=t_personal)

# ==========================================
# --- 10. MÓDULO DE ESTADÍSTICA ---
# ==========================================

def generar_cedula_escolar(nro_parto, anio_nino, cedula_rep):
    # [Nro Parto] + [Últimos 2 dígitos Año del Niño] + [Cédula Madre/Representante]
    anio_str = str(anio_nino)[-2:]
    return f"{str(nro_parto)}{anio_str}{str(cedula_rep)}"

@academico_bp.route('/estadistica_global')
def estadistica_global():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))
        
    estudiantes = Estudiante.query.all() or []
    total_estudiantes = len(estudiantes)
    ultimos_ingresos = Estudiante.query.filter_by(estatus='Activo').order_by(Estudiante.fecha_registro.desc()).limit(3).all() or []
    ultimos_egresos = Estudiante.query.filter_by(estatus='Egreso').order_by(Estudiante.fecha_registro.desc()).limit(3).all() or []
    
    return render_template('estadistica.html', 
                           estudiantes=estudiantes, 
                           total_estudiantes=total_estudiantes,
                           ultimos_ingresos=ultimos_ingresos,
                           ultimos_egresos=ultimos_egresos,
                           grados=Grado.query.all() or [])

@academico_bp.route('/api/estudiantes/<int:grado_id>')
def api_estudiantes_por_grado(grado_id):
    """API ligera: devuelve estudiantes de un grado en JSON para el modal dinámico."""
    if not session.get('logeado'):
        return jsonify({'error': 'No autorizado'}), 401

    grado = Grado.query.get_or_404(grado_id)
    estudiantes = Estudiante.query.filter_by(grado_id=grado_id).order_by(Estudiante.nombre_completo.asc()).all()
    
    resultado = {
        'grado_nombre': grado.nombre,
        'total': len(estudiantes),
        'estudiantes': [{
            'id': est.id,
            'cedula_escolar': est.cedula_escolar,
            'nombre_completo': est.nombre_completo,
            'inicial': est.nombre_completo[0].upper() if est.nombre_completo else '?',
            'genero': est.genero or 'No especificado',
            'estatus': est.estatus or 'Activo',
            'cedula_rep': est.representante_info.cedula if est.representante_info else 'N/A'
        } for est in estudiantes]
    }
    return jsonify(resultado)

@academico_bp.route('/registrar_estudiante', methods=['POST'])
def registrar_estudiante():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))

    # Obtener datos del representante
    cedula_rep = request.form.get('cedula_representante')
    
    # Verificar si el representante ya existe para no duplicarlo
    representante = Representante.query.filter_by(cedula=cedula_rep).first()
    if not representante:
        representante = Representante(
            cedula=cedula_rep,
            nombre_completo=request.form.get('nombre_representante'),
            telefono=request.form.get('telefono_representante'),
            email=request.form.get('email_representante'),
            email_representante=request.form.get('email_representante'),
            parentesco=request.form.get('parentesco'),
            direccion_habitacion=request.form.get('direccion_habitacion'),
            direccion_completa=request.form.get('direccion_completa')
        )
        db.session.add(representante)
        db.session.flush() # Para poder usar el id del representante
    else:
        # Update if not set or just leave it
        if request.form.get('direccion_habitacion'):
            representante.direccion_habitacion = request.form.get('direccion_habitacion')
        if request.form.get('direccion_completa'):
            representante.direccion_completa = request.form.get('direccion_completa')
        if request.form.get('email_representante'):
            representante.email = request.form.get('email_representante')
            representante.email_representante = request.form.get('email_representante')
        db.session.flush()
        
    # Calcular datos para generar cédula escolar
    nro_parto = request.form.get('nro_parto', '1')
    fecha_nac_str = request.form.get('fecha_nacimiento')
    fecha_nac = datetime.strptime(fecha_nac_str, '%Y-%m-%d').date()
    anio_nino = fecha_nac.year
    
    cedula_escolar_gen = generar_cedula_escolar(nro_parto, anio_nino, cedula_rep)
    
    estudiante_existente = Estudiante.query.filter_by(cedula_escolar=cedula_escolar_gen).first()
    if estudiante_existente:
        grado_nombre = estudiante_existente.grado.nombre if estudiante_existente.grado else "el sistema"
        flash(f'Error: Este niño ya ha sido registrado anteriormente y se encuentra en {grado_nombre}.', 'error')
        return redirect(url_for('academico.estadistica_global'))
    
    # Valores numéricos o nulos
    talla_val = request.form.get('talla')
    peso_val = request.form.get('peso')

    # Guardar Estudiante
    estudiante = Estudiante(
        cedula_escolar=cedula_escolar_gen,
        nombre_completo=request.form.get('nombre_estudiante'),
        fecha_nacimiento=fecha_nac,
        lugar_nacimiento=request.form.get('lugar_nacimiento'),
        genero=request.form.get('genero'),
        num_acta=request.form.get('num_acta'),
        num_oficio=request.form.get('num_oficio'),
        talla=float(talla_val) if talla_val else None,
        peso=float(peso_val) if peso_val else None,
        calzado=request.form.get('calzado'),
        talla_camisa=request.form.get('talla_camisa'),
        talla_pantalon=request.form.get('talla_pantalon'),
        tipaje=request.form.get('tipaje'),
        vacunacion_completa=request.form.get('vacunacion_completa'),
        alergias=request.form.get('alergias'),
        neurodivergencia=True if request.form.get('neurodivergencia') == 'on' else False,
        neuro_detalle=request.form.get('neuro_detalle'),
        literal=request.form.get('literal_escolar'),
        literal_escolar=request.form.get('literal_escolar'),
        procedencia=request.form.get('plantel_procedencia'),
        plantel_procedencia=request.form.get('plantel_procedencia'),
        email_estudiante=request.form.get('email_estudiante'),
        es_repetidor=True if request.form.get('es_repetidor') == 'on' else False,
        doc_partida=True if request.form.get('doc_partida') == 'on' else False,
        doc_sano=True if request.form.get('doc_sano') == 'on' else False,
        doc_vacuna=True if request.form.get('doc_vacuna') == 'on' else False,
        lateralidad=request.form.get('lateralidad'),
        nuevo_ingreso=True if request.form.get('nuevo_ingreso') == 'on' else False,
        institucion_procedencia=request.form.get('institucion_procedencia'),
        estatus=request.form.get('estatus', 'Activo'),
        grado_id=request.form.get('grado_id'),
        representante_id=representante.id
    )
    
    try:
        db.session.add(estudiante)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        flash('Ocurrió un error inesperado al guardar el estudiante. Inténtalo de nuevo.', 'error')
    
    return redirect(url_for('academico.estadistica_global'))

@academico_bp.route('/perfil_estudiante/<int:id>')
def perfil_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    estudiante = Estudiante.query.get_or_404(id)
    return render_template('perfil_estudiante.html', estudiante=estudiante, grados=Grado.query.all())

@academico_bp.route('/editar_estudiante/<int:id>', methods=['POST'])
def editar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    est = Estudiante.query.get_or_404(id)
    
    est.nombre_completo = request.form.get('nombre_estudiante')
    est.fecha_nacimiento = datetime.strptime(request.form.get('fecha_nacimiento'), '%Y-%m-%d').date()
    est.lugar_nacimiento = request.form.get('lugar_nacimiento')
    est.genero = request.form.get('genero')
    est.num_acta = request.form.get('num_acta')
    est.num_oficio = request.form.get('num_oficio')
    
    talla_val = request.form.get('talla')
    peso_val = request.form.get('peso')
    est.talla = float(talla_val) if talla_val else None
    est.peso = float(peso_val) if peso_val else None
    
    est.calzado = request.form.get('calzado')
    est.talla_camisa = request.form.get('talla_camisa')
    est.talla_pantalon = request.form.get('talla_pantalon')
    est.tipaje = request.form.get('tipaje')
    est.vacunacion_completa = request.form.get('vacunacion_completa')
    est.alergias = request.form.get('alergias')
    est.neurodivergencia = True if request.form.get('neurodivergencia') == 'on' else False
    est.neuro_detalle = request.form.get('neuro_detalle')
    est.literal = request.form.get('literal_escolar')
    est.literal_escolar = request.form.get('literal_escolar')
    est.procedencia = request.form.get('plantel_procedencia')
    est.plantel_procedencia = request.form.get('plantel_procedencia')
    est.email_estudiante = request.form.get('email_estudiante')
    est.es_repetidor = True if request.form.get('es_repetidor') == 'on' else False
    est.doc_partida = True if request.form.get('doc_partida') == 'on' else False
    est.doc_sano = True if request.form.get('doc_sano') == 'on' else False
    est.doc_vacuna = True if request.form.get('doc_vacuna') == 'on' else False
    est.lateralidad = request.form.get('lateralidad')
    est.institucion_procedencia = request.form.get('institucion_procedencia')
    est.grado_id = request.form.get('grado_id')
    
    db.session.commit()
    return redirect(url_for('academico.perfil_estudiante', id=est.id))

@academico_bp.route('/egresar_estudiante/<int:id>', methods=['POST'])
def egresar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    est = Estudiante.query.get_or_404(id)
    est.estatus = 'Activo' if est.estatus == 'Egreso' else 'Egreso'
    db.session.commit()
    return redirect(url_for('academico.perfil_estudiante', id=est.id))

@academico_bp.route('/eliminar_estudiante/<int:id>', methods=['POST'])
def eliminar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    est = Estudiante.query.get_or_404(id)
    db.session.delete(est)
    db.session.commit()
    return redirect(url_for('academico.estadistica_global'))

# ==========================================
# --- 11. MÓDULO MI AULA ---
# ==========================================

@academico_bp.route('/mi_aula', methods=['GET', 'POST'])
def mi_aula():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    
    rol = session.get('nombre_rol')
    if rol not in ['Administrador Supremo', 'Equipo Directivo (Dirección)', 'Docente de Aula']:
        return redirect(url_for('index'))
        
    grados = []
    grado_seleccionado = None
    estudiantes = []
    asistencia_porcentaje = {}
    grado_id = request.args.get('grado_id')

    if rol == 'Docente de Aula':
        if not grado_id:
            usuario_id = session.get('usuario_id')
            salon = Grado.query.filter(Grado.docentes.any(id=usuario_id)).first()
            if salon:
                return redirect(url_for('academico.mi_aula', grado_id=salon.id))
            else:
                flash("Aún no tienes un grado asignado.", "warning")
                return redirect(url_for('index'))
        else:
            grado_seleccionado = Grado.query.get(grado_id)
            if grado_seleccionado and session.get('usuario_id') in [d.id for d in grado_seleccionado.docentes]:
                estudiantes = Estudiante.query.filter_by(grado_id=grado_seleccionado.id).order_by(Estudiante.nombre_completo.asc()).all()
            else:
                flash("No tienes acceso a este salón.", "danger")
                return redirect(url_for('index'))
    else:
        grados = Grado.query.all()
        if grado_id:
            grado_seleccionado = Grado.query.get(grado_id)
            if grado_seleccionado:
                estudiantes = Estudiante.query.filter_by(grado_id=grado_seleccionado.id).order_by(Estudiante.nombre_completo.asc()).all()

    for est in estudiantes:
        total_dias = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id).count() or 0
        asistencias = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id, estatus='Presente').count() or 0
        porcentaje = (asistencias / total_dias * 100) if total_dias > 0 else 0.0
        asistencia_porcentaje[est.id] = round(porcentaje, 1)
        
    datos_incidencias = {'Conducta': 0, 'Académico': 0, 'Salud': 0, 'Familiar': 0}
    asistencia_promedio_salon = 100.0
    total_matricula = len(estudiantes)

    if estudiantes:
        estudiante_ids = [e.id for e in estudiantes]
        mes_actual = date.today().month
        
        incidencias_salon = Incidencia.query.filter(Incidencia.estudiante_id.in_(estudiante_ids)).all()
        for inc in incidencias_salon:
            if inc.fecha and inc.fecha.month == mes_actual and inc.categoria in datos_incidencias:
                datos_incidencias[inc.categoria] += 1
                
        treinta_dias_atras = date.today() - timedelta(days=30)
        asist_records = AsistenciaEstudiante.query.filter(
            AsistenciaEstudiante.estudiante_id.in_(estudiante_ids),
            AsistenciaEstudiante.fecha >= treinta_dias_atras
        ).all()
        
        if asist_records:
            total_dias_salon = len(asist_records)
            asistencias_positivas = sum(1 for a in asist_records if a.estatus == 'Presente')
            asistencia_promedio_salon = round((asistencias_positivas / total_dias_salon) * 100, 1) if total_dias_salon > 0 else 0.0
            
    total_varones = sum(1 for e in estudiantes if e.genero == 'Masculino')
    total_hembras = sum(1 for e in estudiantes if e.genero == 'Femenino')
    docente_titular = ", ".join([d.nombre_completo for d in grado_seleccionado.docentes]) if (grado_seleccionado and grado_seleccionado.docentes) else "Docente no asignado"

    estado_solicitudes = {}
    if estudiantes:
        mis_solicitudes = SolicitudEnlace.query.filter(
            SolicitudEnlace.estudiante_id.in_([e.id for e in estudiantes])
        ).order_by(SolicitudEnlace.fecha_solicitud.desc()).all()
        for sol in mis_solicitudes:
            if sol.estudiante_id not in estado_solicitudes:
                estado_solicitudes[sol.estudiante_id] = sol

    # Nuevas métricas: Gráfica mensual y Alertas de Defensoría
    datos_mensuales = {}
    alertas_activas = {}
    
    if estudiantes:
        hoy = date.today()
        # Obtener asistencias del mes actual para todos los estudiantes
        asistencias_mes = AsistenciaEstudiante.query.filter(
            AsistenciaEstudiante.estudiante_id.in_([e.id for e in estudiantes])
        ).all()
        
        for est in estudiantes:
            datos_mensuales[est.id] = {'Presente': 0, 'Ausente': 0, 'Justificado': 0}
            alertas_activas[est.id] = False
            
        for a in asistencias_mes:
            if a.fecha and a.fecha.month == hoy.month and a.fecha.year == hoy.year:
                if a.estatus in datos_mensuales[a.estudiante_id]:
                    datos_mensuales[a.estudiante_id][a.estatus] += 1
                    
        # Buscar Alertas de Defensoría activas (Pendiente) para estos estudiantes
        alertas = AlertaDefensoria.query.filter(
            AlertaDefensoria.estudiante_id.in_([e.id for e in estudiantes]),
            AlertaDefensoria.estatus_atencion == 'Pendiente'
        ).all()
        
        for alerta in alertas:
            # Simplificar el motivo para mostrarlo en el badge
            tipo_alerta = "Inasistencia" if alerta.motivo and "inasistencia" in alerta.motivo.lower() else "Incidencia"
            alertas_activas[alerta.estudiante_id] = {'tipo': tipo_alerta, 'id': alerta.id}

    return render_template('mi_aula.html', 
                           grado=grado_seleccionado, 
                           grados=grados, 
                           estudiantes=estudiantes, 
                           asistencia_porcentaje=asistencia_porcentaje,
                           datos_mensuales_estudiantes=datos_mensuales,
                           alertas_estudiantes=alertas_activas,
                           total_matricula=total_matricula,
                           total_varones=total_varones,
                           total_hembras=total_hembras,
                           docente_titular=docente_titular,
                           datos_incidencias=datos_incidencias,
                           asistencia_promedio_salon=asistencia_promedio_salon,
                           estado_solicitudes=estado_solicitudes)

@academico_bp.route('/guardar_asistencia_aula', methods=['POST'])
def guardar_asistencia_aula():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    grado_id = request.form.get('grado_id')
    estudiantes = Estudiante.query.filter_by(grado_id=grado_id).all()
    fecha_hoy = date.today()
    
    for est in estudiantes:
        vino = request.form.get(f'asistio_{est.id}') == 'on'
        registro = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id, fecha=fecha_hoy).first()
        if registro:
            registro.estatus = 'Presente' if vino else 'Ausente'
            registro.grado_id = grado_id
        else:
            nuevo_registro = AsistenciaEstudiante(fecha=fecha_hoy, estatus='Presente' if vino else 'Ausente', estudiante_id=est.id, grado_id=grado_id)
            db.session.add(nuevo_registro)
            
    db.session.commit()
    url = url_for('academico.mi_aula')
    if session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo (Dirección)'] and grado_id:
        url = url_for('academico.mi_aula', grado_id=grado_id)
    return redirect(url)

@academico_bp.route('/agregar_incidencia', methods=['POST'])
def agregar_incidencia():
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    estudiante_id = request.form.get('estudiante_id')
    categoria = request.form.get('categoria')
    descripcion = request.form.get('descripcion')
    grado_id = request.form.get('grado_id')
    
    if estudiante_id and categoria and descripcion:
        incidencia = Incidencia(
            categoria=categoria,
            descripcion=descripcion,
            estudiante_id=estudiante_id,
            usuario_id=session.get('usuario_id')
        )
        db.session.add(incidencia)
        db.session.commit()
        
    url = url_for('academico.mi_aula')
    if session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo (Dirección)'] and grado_id:
        url = url_for('academico.mi_aula', grado_id=grado_id)
    return redirect(url)

@academico_bp.route('/descargar_inscripcion_inicial/<int:grado_id>')
def descargar_inscripcion_inicial(grado_id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    
    grado = Grado.query.get_or_404(grado_id)
    estudiantes = Estudiante.query.filter_by(grado_id=grado.id).order_by(Estudiante.nombre_completo.asc()).all()
    
    docente_nombre = ", ".join([d.nombre_completo for d in grado.docentes]) if grado.docentes else "No asignado"
    
    pdf = FPDF(orientation='L', unit='mm', format='Legal')
    pdf.add_page()
    
    # Membrete Oficial
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(330, 7, f"INSCRIPCIÓN INICIAL - {date.today().year}", ln=1, align='C')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(330, 5, f"Grado y Sección: {grado.nombre} | Docente: {docente_nombre}", ln=1, align='C')
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 7)
    
    # Definición de anchos de las 19 columnas (~335mm disponibles, sumaremos 332)
    w = [6, 38, 16, 14, 8, 8, 24, 8, 38, 17, 18, 50, 45, 7, 7, 7, 7, 8, 8]
    
    # Calcular edad
    def calc_edad(fn):
        if not fn: return 'S/F'
        hoy = date.today()
        return str(hoy.year - fn.year - ((hoy.month, hoy.day) < (fn.month, fn.day)))
    
    headers = ["N°", "Apellidos y Nombres", "C.I. / C.E.", "F. Nac.", "Edad", "Sexo", "Procedencia", "Repet.", "Representante", "C.I. Rep.", "Teléfono", "Correo Representante", "Dirección Habitación", "Lit.", "Talla", "Peso", "Calz.", "T. Cam.", "T. Pan."]
    for i in range(len(headers)):
        pdf.cell(w[i], 6, headers[i], border=1, align='C')
    pdf.ln()
    
    def imprimir_celda_ajustada(ancho, texto, alineacion='L'):
        tamano_original = 7
        tamano_actual = tamano_original
        pdf.set_font_size(tamano_actual)
        while pdf.get_string_width(texto) > (ancho - 1) and tamano_actual > 3:
            tamano_actual -= 0.5
            pdf.set_font_size(tamano_actual)
        pdf.cell(ancho, 6, texto, border=1, align=alineacion)
        pdf.set_font_size(tamano_original)

    pdf.set_font('Arial', '', 7)
    for i, est in enumerate(estudiantes, 1):
        rep_nombre = est.representante_info.nombre_completo if est.representante_info else "Sin registro"
        rep_ci = est.representante_info.cedula if est.representante_info else "Sin registro"
        rep_tlf = est.representante_info.telefono if est.representante_info else "Sin registro"
        rep_email = est.representante_info.email_representante if est.representante_info and est.representante_info.email_representante else (est.representante_info.email if est.representante_info else "Sin registro")
        rep_dir = est.representante_info.direccion_completa if est.representante_info and est.representante_info.direccion_completa else (est.representante_info.direccion_habitacion if est.representante_info else "Sin registro")
        
        nombre_str = est.nombre_completo
        rep_str = rep_nombre
        email_str = rep_email
        dir_str = rep_dir
        proc = est.plantel_procedencia or est.procedencia or est.institucion_procedencia
        proc_str = proc if proc else "Ninguna"
        
        pdf.cell(w[0], 6, str(i), border=1, align='C')
        imprimir_celda_ajustada(w[1], nombre_str, 'L')
        pdf.cell(w[2], 6, str(est.cedula_escolar), border=1, align='C')
        pdf.cell(w[3], 6, est.fecha_nacimiento.strftime('%d/%m/%y') if est.fecha_nacimiento else 'S/F', border=1, align='C')
        pdf.cell(w[4], 6, calc_edad(est.fecha_nacimiento), border=1, align='C')
        pdf.cell(w[5], 6, "M" if est.genero == "Masculino" else ("F" if est.genero == "Femenino" else "-"), border=1, align='C')
        imprimir_celda_ajustada(w[6], proc_str, 'L')
        pdf.cell(w[7], 6, "Sí" if est.es_repetidor else "No", border=1, align='C')
        imprimir_celda_ajustada(w[8], rep_str, 'L')
        pdf.cell(w[9], 6, str(rep_ci), border=1, align='C')
        pdf.cell(w[10], 6, rep_tlf, border=1, align='C')
        imprimir_celda_ajustada(w[11], email_str, 'L')
        imprimir_celda_ajustada(w[12], dir_str, 'L')
        pdf.cell(w[13], 6, est.literal_escolar or est.literal or "-", border=1, align='C')
        pdf.cell(w[14], 6, f"{est.talla or '-'}", border=1, align='C')
        pdf.cell(w[15], 6, f"{est.peso or '-'}", border=1, align='C')
        pdf.cell(w[16], 6, str(est.calzado or '-'), border=1, align='C')
        pdf.cell(w[17], 6, str(est.talla_camisa or '-'), border=1, align='C')
        pdf.cell(w[18], 6, str(est.talla_pantalon or '-'), border=1, align='C')
        pdf.ln()

    from flask import make_response
    response = make_response(pdf.output(dest='S').encode('latin1'))
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'attachment; filename=Inscripcion_Inicial_Grado_{grado.id}.pdf'
    return response

# ==========================================
# --- 10. COMPARTIR EXPEDIENTE SEGURO ---
# ==========================================

@academico_bp.route('/solicitar_enlace/<int:estudiante_id>', methods=['POST'])
def solicitar_enlace(estudiante_id):
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))
        
    motivo = request.form.get('motivo')
    if not motivo:
        flash("Debe proporcionar un motivo para la solicitud.", "error")
        return redirect(url_for('academico.mi_aula'))
        
    nueva_solicitud = SolicitudEnlace(
        docente_id=session['usuario_id'],
        estudiante_id=estudiante_id,
        motivo=motivo,
        estado='Pendiente'
    )
    
    db.session.add(nueva_solicitud)
    db.session.commit()
    
    flash("Solicitud de enlace de expediente enviada a Dirección exitosamente.", "success")
    return redirect(url_for('academico.mi_aula'))

@academico_bp.route('/generar_enlace/<int:id>')
def generar_enlace(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    
    import uuid
    estudiante = Estudiante.query.get_or_404(id)
    token = uuid.uuid4().hex
    
    nuevo_enlace = EnlaceTemporal(token=token, estudiante_id=estudiante.id, usado=False)
    db.session.add(nuevo_enlace)
    db.session.commit()
    
    url_acceso = url_for('academico.ver_expediente', token=token, _external=True)
    mensaje = f"Aquí tienes el expediente temporal de {estudiante.nombre_completo}: {url_acceso}\n\n*Nota: Este enlace es de un solo uso y expirará después de abrirlo por razones de seguridad.*"
    
    whatsapp_url = f"https://wa.me/?text={urllib.parse.quote(mensaje)}"
    return redirect(whatsapp_url)

@academico_bp.route('/aprobar_solicitud/<int:id>', methods=['POST', 'GET'])
def aprobar_solicitud(id):
    if session.get('nombre_rol') not in ['Administrador Supremo', 'Equipo Directivo (Dirección)']:
        flash("No tiene permisos para realizar esta acción.", "danger")
        return redirect(url_for('index'))
        
    solicitud = SolicitudEnlace.query.get_or_404(id)
    solicitud.estado = 'Aprobado'
    solicitud.token_generado = uuid.uuid4().hex[:8]
    db.session.commit()
    
    flash(f"Solicitud aprobada exitosamente. Se ha generado un enlace para {solicitud.estudiante.nombre_completo}.", "success")
    return redirect(url_for('index'))

@academico_bp.route('/rechazar_solicitud/<int:id>', methods=['POST', 'GET'])
def rechazar_solicitud(id):
    if session.get('nombre_rol') not in ['Administrador Supremo', 'Equipo Directivo (Dirección)']:
        flash("No tiene permisos para realizar esta acción.", "danger")
        return redirect(url_for('index'))
        
    solicitud = SolicitudEnlace.query.get_or_404(id)
    solicitud.estado = 'Rechazado'
    db.session.commit()
    
    flash("Solicitud rechazada.", "info")
    return redirect(url_for('index'))

@academico_bp.route('/expediente_publico/<token>')
def expediente_publico(token):
    solicitud = SolicitudEnlace.query.filter_by(token_generado=token, estado='Aprobado').first()
    
    if not solicitud:
        return render_template('enlace_expirado.html'), 404
        
    estudiante_data = solicitud.estudiante
    solicitud.estado = 'Utilizado'
    db.session.commit()
        
    # Prevenir que el navegador guarde la página en caché para mayor seguridad
    from flask import make_response
    response = make_response(render_template('expediente_publico.html', estudiante=estudiante_data))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

@academico_bp.route('/ver_expediente/<token>')
def ver_expediente(token):
    enlace = EnlaceTemporal.query.filter_by(token=token).first()
    
    if not enlace or enlace.usado:
        return render_template('enlace_expirado.html'), 403
        
    enlace.usado = True
    db.session.commit()
    
    # Prevenir que el navegador guarde la página en caché
    from flask import make_response
    response = make_response(render_template('ver_expediente_temporal.html', estudiante=enlace.estudiante))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# ==========================================
# --- 11. GESTIÓN DESCENTRALIZADA (FASE 19) ---
# ==========================================

@academico_bp.route('/nuevo_ingreso_rapido', methods=['POST'])
def nuevo_ingreso_rapido():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))

    # ── 1. Datos del Representante ──
    cedula_rep = request.form.get('cedula_representante')
    if not cedula_rep:
        flash("Debe indicar la cédula del representante.", "danger")
        return redirect(request.referrer or url_for('academico.mi_aula'))

    representante = Representante.query.filter_by(cedula=cedula_rep).first()
    if not representante:
        representante = Representante(
            cedula=cedula_rep,
            nombre_completo=request.form.get('nombre_representante', ''),
            telefono=request.form.get('telefono_representante'),
            email=request.form.get('email_representante'),
            email_representante=request.form.get('email_representante'),
            parentesco=request.form.get('parentesco'),
            direccion_habitacion=request.form.get('direccion_habitacion'),
            direccion_completa=request.form.get('direccion_completa')
        )
        db.session.add(representante)
        db.session.flush()
    else:
        # Actualizar datos del representante si vienen nuevos valores
        if request.form.get('nombre_representante'):
            representante.nombre_completo = request.form.get('nombre_representante')
        if request.form.get('telefono_representante'):
            representante.telefono = request.form.get('telefono_representante')
        if request.form.get('email_representante'):
            representante.email = request.form.get('email_representante')
            representante.email_representante = request.form.get('email_representante')
        if request.form.get('parentesco'):
            representante.parentesco = request.form.get('parentesco')
        if request.form.get('direccion_habitacion'):
            representante.direccion_habitacion = request.form.get('direccion_habitacion')
        if request.form.get('direccion_completa'):
            representante.direccion_completa = request.form.get('direccion_completa')
        db.session.flush()

    # ── 2. Generar Cédula Escolar ──
    fecha_nacimiento_str = request.form.get('fecha_nacimiento')
    try:
        fecha_nac = datetime.strptime(fecha_nacimiento_str, '%Y-%m-%d').date()
    except (ValueError, TypeError):
        flash("Fecha de nacimiento inválida.", "danger")
        return redirect(request.referrer or url_for('academico.mi_aula'))

    nro_parto = request.form.get('nro_parto', '1')
    anio_nino = fecha_nac.year
    cedula_escolar_gen = generar_cedula_escolar(nro_parto, anio_nino, cedula_rep)

    # ── 3. Verificar duplicados ──
    estudiante_existente = Estudiante.query.filter_by(cedula_escolar=cedula_escolar_gen).first()
    if estudiante_existente:
        grado_nombre = estudiante_existente.grado.nombre if estudiante_existente.grado else "el sistema"
        flash(f'La solicitud no fue enviada: Este estudiante ya se encuentra registrado en {grado_nombre}.', 'warning')
        return redirect(request.referrer or url_for('academico.mi_aula'))

    # ── 4. Captura exhaustiva de datos del estudiante ──
    talla_val = request.form.get('talla')
    peso_val = request.form.get('peso')

    nuevo_estudiante = Estudiante(
        # Identidad y Legal
        cedula_escolar=cedula_escolar_gen,
        nombre_completo=request.form.get('nombre_estudiante', ''),
        fecha_nacimiento=fecha_nac,
        lugar_nacimiento=request.form.get('lugar_nacimiento'),
        genero=request.form.get('genero'),
        num_acta=request.form.get('num_acta'),
        num_oficio=request.form.get('num_oficio'),
        # Salud y Antropometría
        talla=float(talla_val) if talla_val else None,
        peso=float(peso_val) if peso_val else None,
        calzado=request.form.get('calzado'),
        talla_camisa=request.form.get('talla_camisa'),
        talla_pantalon=request.form.get('talla_pantalon'),
        tipaje=request.form.get('tipaje'),
        vacunacion_completa=request.form.get('vacunacion_completa'),
        alergias=request.form.get('alergias'),
        neurodivergencia=True if request.form.get('neurodivergencia') == 'on' else False,
        neuro_detalle=request.form.get('neuro_detalle'),
        lateralidad=request.form.get('lateralidad'),
        # Procedencia y Estado Académico
        literal=request.form.get('literal_escolar'),
        literal_escolar=request.form.get('literal_escolar'),
        procedencia=request.form.get('plantel_procedencia'),
        plantel_procedencia=request.form.get('plantel_procedencia'),
        email_estudiante=request.form.get('email_estudiante'),
        es_repetidor=True if request.form.get('es_repetidor') == 'on' else False,
        nuevo_ingreso=True if request.form.get('nuevo_ingreso') == 'on' else True,
        estatus='Activo',
        # Relaciones
        grado_id=request.form.get('grado_id'),
        representante_id=representante.id
    )

    try:
        db.session.add(nuevo_estudiante)
        db.session.commit()
        flash(f"Estudiante {nuevo_estudiante.nombre_completo} registrado exitosamente.", "success")
    except Exception as e:
        db.session.rollback()
        flash('Ocurrió un error inesperado al guardar el estudiante. Inténtalo de nuevo.', 'danger')

    return redirect(request.referrer or url_for('academico.mi_aula'))

@academico_bp.route('/actualizar_estudiante_rapido/<int:id>', methods=['POST'])
def actualizar_estudiante_rapido(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    
    estudiante = Estudiante.query.get_or_404(id)
    
    # Crear la solicitud en lugar de un UPDATE directo
    nueva_solicitud = SolicitudActualizacion(
        docente_id=session['usuario_id'],
        estudiante_id=estudiante.id,
        rep_telefono=request.form.get('rep_telefono'),
        rep_direccion=request.form.get('rep_direccion'),
        alergias=request.form.get('alergias'),
        neuro_detalle=request.form.get('neuro_detalle'),
        estado='Pendiente'
    )
    
    db.session.add(nueva_solicitud)
    db.session.commit()
    
    flash(f"Solicitud de actualización para {estudiante.nombre_completo} enviada a Dirección para su revisión.", "info")
    return redirect(request.referrer or url_for('academico.mi_aula'))

@academico_bp.route('/actualizar_medidas/<int:id>', methods=['POST'])
def actualizar_medidas(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))

    estudiante = Estudiante.query.get_or_404(id)
    grado_id = request.form.get('grado_id')

    def _a_float(valor):
        try:
            return float(valor) if valor not in (None, '') else None
        except ValueError:
            return None

    # Guardado directo: a diferencia de "Actualizar" (contacto/ficha médica),
    # las medidas físicas no pasan por aprobación de Dirección.
    estudiante.peso = _a_float(request.form.get('peso'))
    estudiante.talla = _a_float(request.form.get('talla'))
    estudiante.talla_camisa = request.form.get('talla_camisa')
    estudiante.talla_pantalon = request.form.get('talla_pantalon')
    estudiante.calzado = request.form.get('calzado')

    db.session.commit()
    flash(f"Medidas de {estudiante.nombre_completo} actualizadas correctamente.", "success")

    url = url_for('academico.mi_aula')
    if session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo (Dirección)'] and grado_id:
        url = url_for('academico.mi_aula', grado_id=grado_id)
    return redirect(url)

@academico_bp.route('/aprobar_actualizacion/<int:id>', methods=['POST', 'GET'])
def aprobar_actualizacion(id):
    if session.get('nombre_rol') not in ['Administrador Supremo', 'Equipo Directivo (Dirección)']:
        flash("No tiene permisos para realizar esta acción.", "danger")
        return redirect(url_for('index'))
        
    solicitud = SolicitudActualizacion.query.get_or_404(id)
    estudiante = solicitud.estudiante
    
    if estudiante.representante_info:
        if solicitud.rep_telefono:
            estudiante.representante_info.telefono = solicitud.rep_telefono
        if solicitud.rep_direccion:
            estudiante.representante_info.direccion_completa = solicitud.rep_direccion
            
    if solicitud.alergias is not None:
        estudiante.alergias = solicitud.alergias
    if solicitud.neuro_detalle is not None:
        estudiante.neuro_detalle = solicitud.neuro_detalle
        estudiante.neurodivergencia = bool(solicitud.neuro_detalle.strip())
        
    solicitud.estado = 'Aprobado'
    db.session.commit()
    
    flash(f"Actualización para {estudiante.nombre_completo} aprobada y aplicada.", "success")
    return redirect(url_for('index'))

@academico_bp.route('/rechazar_actualizacion/<int:id>', methods=['POST', 'GET'])
def rechazar_actualizacion(id):
    if session.get('nombre_rol') not in ['Administrador Supremo', 'Equipo Directivo (Dirección)']:
        flash("No tiene permisos para realizar esta acción.", "danger")
        return redirect(url_for('index'))
        
    solicitud = SolicitudActualizacion.query.get_or_404(id)
    solicitud.estado = 'Rechazado'
    db.session.commit()
    
    flash("Solicitud de actualización rechazada.", "info")
    return redirect(url_for('index'))

# ==========================================
# --- ASISTENCIA POR ESTUDIANTE Y ALERTAS ---
# ==========================================

@academico_bp.route('/asistencia_estudiantes', methods=['GET'])
def asistencia_estudiantes():
    """Página principal de asistencia por estudiante."""
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))
    
    grados = Grado.query.all()
    fecha_hoy = date.today().strftime('%Y-%m-%d')
    
    return render_template('asistencia_estudiantes.html', grados=grados, fecha_hoy=fecha_hoy)

@academico_bp.route('/api/asistencia_estudiantes/<int:grado_id>/<string:fecha>', methods=['GET'])
def api_asistencia_estudiantes(grado_id, fecha):
    """API: Devuelve la lista de estudiantes de un grado con su estatus de asistencia para una fecha."""
    if not session.get('logeado'):
        return jsonify({'error': 'No autorizado'}), 401
    
    try:
        fecha_obj = datetime.strptime(fecha, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'Formato de fecha inválido'}), 400
    
    estudiantes = Estudiante.query.filter_by(grado_id=grado_id, estatus='Activo').order_by(Estudiante.nombre_completo.asc()).all()
    
    resultado = []
    for est in estudiantes:
        # Buscar si ya tiene asistencia registrada para esta fecha
        asistencia = AsistenciaEstudiante.query.filter_by(
            estudiante_id=est.id, fecha=fecha_obj
        ).first()
        
        resultado.append({
            'id': est.id,
            'nombre_completo': est.nombre_completo,
            'cedula_escolar': est.cedula_escolar,
            'genero': est.genero or 'N/A',
            'estatus': asistencia.estatus if asistencia else None  # None = no registrado aún
        })
    
    return jsonify({'estudiantes': resultado, 'total': len(resultado)})

@academico_bp.route('/guardar_asistencia_estudiantes', methods=['POST'])
def guardar_asistencia_estudiantes():
    """Guarda la asistencia de todos los estudiantes de un grado y ejecuta el disparador de alertas."""
    if not session.get('logeado'):
        return jsonify({'error': 'No autorizado'}), 401
    
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No se recibieron datos'}), 400
    
    grado_id = data.get('grado_id')
    fecha_str = data.get('fecha')
    asistencias = data.get('asistencias', [])  # [{estudiante_id: X, estatus: 'Presente'|'Ausente'|'Justificado'}]
    
    if not grado_id or not fecha_str or not asistencias:
        return jsonify({'error': 'Faltan datos requeridos'}), 400
    
    try:
        fecha_obj = datetime.strptime(fecha_str, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'Formato de fecha inválido'}), 400
    
    try:
        # Insertar o actualizar asistencia de cada estudiante
        for registro in asistencias:
            est_id = registro.get('estudiante_id')
            estatus = registro.get('estatus', 'Presente')
            
            # Validar estatus
            if estatus not in ('Presente', 'Ausente', 'Justificado'):
                estatus = 'Presente'
            
            # Buscar registro existente para esta fecha y estudiante
            existente = AsistenciaEstudiante.query.filter_by(
                estudiante_id=est_id, fecha=fecha_obj
            ).first()
            
            if existente:
                existente.estatus = estatus
                existente.grado_id = grado_id
            else:
                nuevo = AsistenciaEstudiante(
                    estudiante_id=est_id,
                    grado_id=grado_id,
                    fecha=fecha_obj,
                    estatus=estatus
                )
                db.session.add(nuevo)
        
        db.session.flush()  # Aplicar cambios antes de verificar alertas
        
        # =============================================
        # DISPARADOR: Verificar alertas de Defensoría
        # =============================================
        alertas_generadas = 0
        
        # Calcular rango de la semana (lunes a viernes) de la fecha dada
        dia_semana = fecha_obj.weekday()  # 0=Lunes, 6=Domingo
        lunes = fecha_obj - timedelta(days=dia_semana)
        viernes = lunes + timedelta(days=4)
        semana_iso = fecha_obj.strftime('%G-W%V')  # Ej: '2026-W31'
        
        # Solo verificar estudiantes marcados como Ausentes
        ausentes_ids = [r['estudiante_id'] for r in asistencias if r.get('estatus') == 'Ausente']
        
        for est_id in ausentes_ids:
            # Contar inasistencias en la semana actual
            conteo_ausencias = AsistenciaEstudiante.query.filter(
                AsistenciaEstudiante.estudiante_id == est_id,
                AsistenciaEstudiante.estatus == 'Ausente',
                AsistenciaEstudiante.fecha >= lunes,
                AsistenciaEstudiante.fecha <= viernes
            ).count()
            
            if conteo_ausencias >= 3:
                # Verificar si ya existe alerta pendiente para esta semana
                alerta_existente = AlertaDefensoria.query.filter_by(
                    estudiante_id=est_id,
                    semana_iso=semana_iso
                ).first()
                
                if not alerta_existente:
                    estudiante = Estudiante.query.get(est_id)
                    nombre = estudiante.nombre_completo if estudiante else f'ID {est_id}'
                    grado = Grado.query.get(grado_id)
                    grado_nombre = grado.nombre if grado else ''
                    
                    # Escalación automática: si el estudiante ya tiene 2+ alertas previas
                    # de semanas distintas, escalar directamente a Visita Domiciliaria
                    alertas_previas = AlertaDefensoria.query.filter(
                        AlertaDefensoria.estudiante_id == est_id,
                        AlertaDefensoria.semana_iso != semana_iso
                    ).count()
                    
                    if alertas_previas >= 2:
                        estatus_auto = 'Visita Domiciliaria'
                        motivo_texto = f'⚠️ ESCALADA AUTOMÁTICA: El estudiante {nombre} ({grado_nombre}) acumula {conteo_ausencias} inasistencias en la semana del {lunes.strftime("%d/%m/%Y")} al {viernes.strftime("%d/%m/%Y")}. Este estudiante tiene {alertas_previas} alertas previas de semanas anteriores.'
                    else:
                        estatus_auto = 'Pendiente'
                        motivo_texto = f'El estudiante {nombre} ({grado_nombre}) acumula {conteo_ausencias} inasistencias en la semana del {lunes.strftime("%d/%m/%Y")} al {viernes.strftime("%d/%m/%Y")}.'
                    
                    nueva_alerta = AlertaDefensoria(
                        estudiante_id=est_id,
                        fecha_emision=date.today(),
                        motivo=motivo_texto,
                        estatus_atencion=estatus_auto,
                        semana_iso=semana_iso
                    )
                    db.session.add(nueva_alerta)
                    alertas_generadas += 1
        
        # =============================================
        # ACTUALIZAR REGISTRO DIARIO (AsistenciaDiaria)
        # =============================================
        grado = Grado.query.get(grado_id)
        if grado:
            varones_asist = 0
            hembras_asist = 0
            
            for registro in asistencias:
                if registro.get('estatus') in ('Presente', 'Justificado'):
                    est = Estudiante.query.get(registro.get('estudiante_id'))
                    if est:
                        genero_est = (est.genero or '').upper()
                        if genero_est.startswith('M') or genero_est == 'VARON' or genero_est == 'NIÑO':
                            varones_asist += 1
                        else:
                            hembras_asist += 1
                            
            mat_total = grado.total_varones + grado.total_hembras
            asist_total = varones_asist + hembras_asist
            porc = round((asist_total / mat_total) * 100, 2) if mat_total > 0 else 0
            
            reg_diario = AsistenciaDiaria.query.filter_by(fecha=fecha_obj, grado_seccion=grado.nombre).first()
            if reg_diario:
                reg_diario.varones = varones_asist
                reg_diario.hembras = hembras_asist
                reg_diario.asistentes = asist_total
                reg_diario.porcentaje = porc
            else:
                nuevo_diario = AsistenciaDiaria(
                    fecha=fecha_obj,
                    grado_seccion=grado.nombre,
                    matricula_total=mat_total,
                    varones=varones_asist,
                    hembras=hembras_asist,
                    asistentes=asist_total,
                    porcentaje=porc,
                    usuario_id=session.get('usuario_id', 1)
                )
                db.session.add(nuevo_diario)

        db.session.commit()
        
        mensaje = f'Asistencia guardada correctamente para {len(asistencias)} estudiantes.'
        if alertas_generadas > 0:
            mensaje += f' ⚠️ Se generaron {alertas_generadas} alerta(s) para Defensoría.'
        
        return jsonify({'success': True, 'message': mensaje, 'alertas_generadas': alertas_generadas})
    
    except Exception as e:
        db.session.rollback()
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Error al guardar: {str(e)}'}), 500

@academico_bp.route('/alertas_defensoria')
def alertas_defensoria():
    """Vista de alertas generadas para Defensoría."""
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))
    
    # Solo Defensoría, Directivo o Admin pueden ver
    rol = session.get('nombre_rol', '')
    if rol not in ['Defensoría Estudiantil', 'Equipo Directivo (Dirección)', 'Administrador Supremo']:
        flash('No tienes permisos para acceder a esta sección.', 'error')
        return redirect(url_for('index'))
    
    alertas = AlertaDefensoria.query.order_by(AlertaDefensoria.fecha_emision.desc()).all()
    return render_template('alertas_defensoria.html', alertas=alertas)

@academico_bp.route('/actualizar_alerta/<int:id>', methods=['POST'])
def actualizar_alerta(id):
    """Actualiza el estatus de atención de una alerta."""
    if not session.get('logeado'):
        return jsonify({'error': 'No autorizado'}), 401
    
    alerta = AlertaDefensoria.query.get_or_404(id)
    nuevo_estatus = request.form.get('estatus_atencion', request.json.get('estatus_atencion') if request.is_json else None)
    
    if nuevo_estatus in ('Pendiente', 'Contactado', 'Visita Domiciliaria'):
        alerta.estatus_atencion = nuevo_estatus
        db.session.commit()
        # No usamos flash() aquí para evitar que los mensajes se acumulen y salgan en otros módulos
    
    referer = request.headers.get("Referer")
    if referer and "defensoria" in referer and "alertas_defensoria" not in referer:
        return redirect(url_for('defensoria'))
    return redirect(url_for('academico.alertas_defensoria'))

@academico_bp.route('/enterado_alerta/<int:id>', methods=['POST'])
def enterado_alerta(id):
    if not session.get('logeado'): return redirect(url_for('auth.login'))
    rol = session.get('nombre_rol')
    if rol not in ['Administrador Supremo', 'Equipo Directivo (Dirección)', 'Docente de Aula']:
        return "Acceso Denegado", 403
    alerta = AlertaDefensoria.query.get_or_404(id)
    # Simplemente eliminamos la alerta de incidencia para que ya no aparezca
    db.session.delete(alerta)
    db.session.commit()
    
    grado_id = request.args.get('grado_id')
    return redirect(url_for('academico.mi_aula', grado_id=grado_id))

@academico_bp.route('/eliminar_alerta/<int:id>', methods=['POST'])
def eliminar_alerta(id):
    """Elimina una alerta de defensoría permanentemente."""
    if not session.get('logeado'):
        return jsonify({'error': 'No autorizado'}), 401
    
    # Solo Defensoría, Directivo o Admin pueden borrar
    rol = session.get('nombre_rol', '')
    if rol not in ['Defensoría Estudiantil', 'Equipo Directivo (Dirección)', 'Administrador Supremo']:
        return redirect(url_for('index'))
        
    alerta = AlertaDefensoria.query.get_or_404(id)
    db.session.delete(alerta)
    db.session.commit()
    # No usamos flash() para evitar mensajes huérfanos en portal_trabajador
    
    # Redirect back to where they came from (defensoria or alertas_defensoria)
    referer = request.headers.get("Referer")
    if referer and "defensoria" in referer and "alertas_defensoria" not in referer:
        return redirect(url_for('defensoria'))
    return redirect(url_for('academico.alertas_defensoria'))
