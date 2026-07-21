from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, make_response
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.exc import OperationalError
from flask_migrate import Migrate
from datetime import datetime, date, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
import io
import os
import pandas as pd
import openpyxl
from docx import Document
from fpdf import FPDF
from flask_mail import Mail, Message
import secrets
import urllib.parse
from dotenv import load_dotenv

# Cargar variables de entorno desde .env antes de cualquier configuración
load_dotenv()

# Configuración de rutas para evitar errores en el servidor
basedir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-fallback-key-cambiar-en-produccion')

# Forzamos la ruta al archivo que está dentro de mysite
db_path = os.path.join(basedir, 'roboclass.db')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + db_path

# Carpeta para guardar los documentos privados
app.config['UPLOAD_FOLDER'] = os.path.join(basedir, 'uploads', 'personal')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Imprimimos en el log para estar 100% seguros
print(f"--- CONECTADO A: {db_path} ---")
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAIL_SERVER'] = os.environ.get('MAIL_SERVER', 'smtp.gmail.com')
app.config['MAIL_PORT'] = int(os.environ.get('MAIL_PORT', 587))
app.config['MAIL_USE_TLS'] = os.environ.get('MAIL_USE_TLS', 'True').lower() in ('true', '1', 'yes')
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USER')
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASS')

mail = Mail(app)
db = SQLAlchemy(app)
migrate = Migrate(app, db)

# ==========================================
# --- 1. MODELOS DE BASE DE DATOS ---
# ==========================================

class Rol(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(50), unique=True, nullable=False)
    permisos = db.Column(db.String(500), nullable=False, default="")
    usuarios = db.relationship('Usuario', backref='rol_info', lazy=True)

class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    nombre_completo = db.Column(db.String(100), nullable=False) 
    email = db.Column(db.String(120), unique=True, nullable=False)
    area_trabajo = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(200), nullable=False)
    rol_id = db.Column(db.Integer, db.ForeignKey('rol.id'), nullable=True)
    # Autogestión MPPE y Expediente Digital
    usuario_autogestion = db.Column(db.String(100), nullable=True)
    clave_autogestion = db.Column(db.String(100), nullable=True)
    voucher_path = db.Column(db.String(250), nullable=True)
    constancia_path = db.Column(db.String(250), nullable=True)
    curriculum_path = db.Column(db.String(250), nullable=True)
    rif_path = db.Column(db.String(250), nullable=True)
    cedula_path = db.Column(db.String(250), nullable=True)
    foto_perfil_path = db.Column(db.String(250), nullable=True)
    cargo_solicitado = db.Column(db.String(100), nullable=True)

class Anuncio(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    titulo = db.Column(db.String(100), nullable=False)
    mensaje = db.Column(db.Text, nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now)
    autor_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    autor = db.relationship('Usuario', backref='anuncios_creados')

class Bitacora(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, default=datetime.now)
    grado = db.Column(db.String(50), nullable=False)
    actividad = db.Column(db.String(200), nullable=False)
    estado = db.Column(db.String(50), default='Completado')
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class PlanificacionDefensoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tema_charla = db.Column(db.String(200), nullable=False)
    proposito = db.Column(db.Text, nullable=False)
    poblacion_objetivo = db.Column(db.String(100), nullable=False) 
    fecha_programada = db.Column(db.Date, nullable=False)
    hora_programada = db.Column(db.String(20), nullable=False)
    estado = db.Column(db.String(50), default='Pendiente')
    notas_seguimiento = db.Column(db.Text, nullable=True) 
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class Grado(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(50), nullable=False)
    total_varones = db.Column(db.Integer, default=0)
    total_hembras = db.Column(db.Integer, default=0)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    docente_info = db.relationship('Usuario', backref='grados_asignados', lazy=True)

class Tema(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class AsistenciaDiaria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False)
    grado_seccion = db.Column(db.String(100), nullable=False)
    matricula_total = db.Column(db.Integer, nullable=False)
    varones = db.Column(db.Integer, nullable=False, default=0)
    hembras = db.Column(db.Integer, nullable=False, default=0)
    asistentes = db.Column(db.Integer, nullable=False)
    porcentaje = db.Column(db.Float, nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class AsistenciaPersonal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False)
    categoria = db.Column(db.String(50), nullable=False)
    matricula_base = db.Column(db.Integer, nullable=False, default=0)
    asistentes = db.Column(db.Integer, nullable=False, default=0)
    porcentaje = db.Column(db.Float)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class Representante(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cedula = db.Column(db.String(20), unique=True, nullable=False)
    nombre_completo = db.Column(db.String(100), nullable=False)
    telefono = db.Column(db.String(20))
    parentesco = db.Column(db.String(50)) # Madre, Padre, etc.
    email = db.Column(db.String(120))
    email_representante = db.Column(db.String(120))
    direccion_habitacion = db.Column(db.Text)
    direccion_completa = db.Column(db.Text)
    estudiantes = db.relationship('Estudiante', backref='representante_info', lazy=True)

class Estudiante(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    # Identidad y Legal
    cedula_escolar = db.Column(db.String(30), unique=True, nullable=False)
    nombre_completo = db.Column(db.String(150), nullable=False)
    fecha_nacimiento = db.Column(db.Date, nullable=False)
    lugar_nacimiento = db.Column(db.String(100))
    genero = db.Column(db.String(20))
    num_acta = db.Column(db.String(50))
    num_oficio = db.Column(db.String(50))
    # Salud y Nutrición
    talla = db.Column(db.Float) # En cm
    peso = db.Column(db.Float) # En kg
    calzado = db.Column(db.String(10))
    talla_camisa = db.Column(db.String(50))
    talla_pantalon = db.Column(db.String(50))
    tipaje = db.Column(db.String(10)) # Grupo sanguíneo
    vacunacion_completa = db.Column(db.String(20)) # Si/No/Parcial
    alergias = db.Column(db.String(200))
    neurodivergencia = db.Column(db.Boolean, default=False)
    neuro_detalle = db.Column(db.String(200))
    # Procedencia y Estado
    literal = db.Column(db.String(2))
    literal_escolar = db.Column(db.String(2))
    procedencia = db.Column(db.Text)
    plantel_procedencia = db.Column(db.String(150))
    email_estudiante = db.Column(db.String(120))
    es_repetidor = db.Column(db.Boolean, default=False)
    doc_partida = db.Column(db.Boolean, default=False)
    doc_sano = db.Column(db.Boolean, default=False)
    doc_vacuna = db.Column(db.Boolean, default=False)
    lateralidad = db.Column(db.String(20))
    nuevo_ingreso = db.Column(db.Boolean, default=False)
    institucion_procedencia = db.Column(db.String(150))
    estatus = db.Column(db.String(20), default='Activo') # Activo o Egreso
    fecha_registro = db.Column(db.DateTime, default=datetime.now)
    # Relaciones
    grado_id = db.Column(db.Integer, db.ForeignKey('grado.id'))
    representante_id = db.Column(db.Integer, db.ForeignKey('representante.id'))

class EnlaceTemporal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    token = db.Column(db.String(50), unique=True, nullable=False)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    usado = db.Column(db.Boolean, default=False)
    fecha_creacion = db.Column(db.DateTime, default=datetime.now)
    estudiante = db.relationship('Estudiante', backref='enlaces_temporales')

class Incidencia(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, default=datetime.now)
    categoria = db.Column(db.String(50), nullable=False) # Conducta, Académico, Salud, Familiar
    descripcion = db.Column(db.Text, nullable=False)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

class AsistenciaEstudiante(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, nullable=False)
    asistio = db.Column(db.Boolean, nullable=False)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)

estudiante_brigada = db.Table('estudiante_brigada',
    db.Column('estudiante_id', db.Integer, db.ForeignKey('estudiante.id'), primary_key=True),
    db.Column('brigada_id', db.Integer, db.ForeignKey('brigada.id'), primary_key=True)
)

class Brigada(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    descripcion = db.Column(db.Text)
    estudiantes = db.relationship('Estudiante', secondary=estudiante_brigada, lazy='subquery',
        backref=db.backref('brigadas', lazy=True))

class Acta(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, default=datetime.now)
    contenido_manual = db.Column(db.Text, nullable=False)
    tipo_acta = db.Column(db.String(50), default='General')
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=True)
    defensor_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    # Relationships
    estudiante = db.relationship('Estudiante', backref='actas')
    defensor = db.relationship('Usuario', backref='actas_creadas')

# ==========================================
# --- 2. INICIALIZACIÓN ---
# ==========================================

with app.app_context():
    db.create_all()  # Inicialización Automática: Crea las tablas faltantes si el archivo db es nuevo
    # Mapa maestro de permisos por rol
    PERMISOS_POR_ROL = {
        'Administrador Supremo': 'dashboard,planificador,asistencia,defensoria,configuracion,admin,anuncios',
        'Equipo Directivo': 'dashboard,asistencia,configuracion,admin,anuncios',
        'Docente de Aula': 'planificador,asistencia,anuncios',
        'Docente Especialista': 'planificador,asistencia,anuncios',
        'Defensoría Estudiantil': 'defensoria,anuncios',
        'Obrero': '',
        'Personal de Vigilancia': '',
        'Personal de Cocina': '',
    }

    if not Rol.query.first():
        for nombre, permisos in PERMISOS_POR_ROL.items():
            db.session.add(Rol(nombre=nombre, permisos=permisos))
        db.session.commit()
    else:
        # Migración: sincronizar permisos de roles existentes y crear faltantes
        for nombre, permisos in PERMISOS_POR_ROL.items():
            rol = Rol.query.filter_by(nombre=nombre).first()
            if rol:
                rol.permisos = permisos
            else:
                db.session.add(Rol(nombre=nombre, permisos=permisos))
        db.session.commit()

@app.context_processor
def inyectar_datos_globales():
    return dict(
        mis_permisos=session.get('permisos', ''),
        mi_rol=session.get('nombre_rol', 'Sin Rol'),
        total_anuncios=Anuncio.query.count(),
        hoy_str=datetime.now().strftime('%d/%m/%Y')
    )

def obtener_estudiantes_por_docente(usuario_id):
    grado = Grado.query.filter_by(usuario_id=usuario_id).first()
    if grado:
        return Estudiante.query.filter_by(grado_id=grado.id).order_by(Estudiante.nombre_completo.asc()).all()
    return []

def auth_defensoria():
    if not session.get('logeado'): return False
    rol_permitido = session.get('nombre_rol') in ['Defensoría Estudiantil', 'Equipo Directivo', 'Administrador Supremo']
    return rol_permitido

# ==========================================
# --- 3. DASHBOARD Y AUTENTICACIÓN ---
# ==========================================

@app.route('/')
def index():
    if not session.get('logeado'): return render_template('landing.html')
    permisos = session.get('permisos', '')
    if 'dashboard' not in permisos:
        if session.get('nombre_rol') in ['Obrero', 'Personal de Vigilancia', 'Personal de Cocina']:
            return redirect(url_for('portal_trabajador'))
        if 'planificador' in permisos: return redirect(url_for('planificador'))
        if 'defensoria' in permisos: return redirect(url_for('defensoria'))
        if 'asistencia' in permisos: return redirect(url_for('asistencia'))
        return redirect(url_for('en_espera'))

    ahora = datetime.now()
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    fecha_hoy_esp = f"{dias[ahora.weekday()]}, {ahora.day} de {meses[ahora.month-1]} de {ahora.year}"

    registros_hoy = AsistenciaDiaria.query.filter_by(fecha=ahora.date()).all()
    mat_hoy = sum(r.matricula_total for r in registros_hoy)
    asist_hoy = sum(r.asistentes for r in registros_hoy)
    porc_hoy = round((asist_hoy / mat_hoy) * 100, 1) if mat_hoy > 0 else 0.0

    registros_pers_hoy = AsistenciaPersonal.query.filter_by(fecha=date.today()).all()
    mat_pers_hoy = sum(r.matricula_base for r in registros_pers_hoy)
    asist_pers_hoy = sum(r.asistentes for r in registros_pers_hoy)
    porc_pers_hoy = round((asist_pers_hoy / mat_pers_hoy) * 100, 1) if mat_pers_hoy > 0 else 0.0
    
    desglose_personal = {'Docentes': 0, 'Administrativos': 0, 'Obreros': 0, 'Cocina': 0, 'Vigilantes': 0}
    for r in registros_pers_hoy:
        if r.categoria in desglose_personal: desglose_personal[r.categoria] += r.asistentes

    total_v = sum(g.total_varones for g in Grado.query.all())
    total_h = sum(g.total_hembras for g in Grado.query.all())
    ultimos_anuncios = Anuncio.query.order_by(Anuncio.fecha.desc()).limit(3).all()

    historico = AsistenciaDiaria.query.order_by(AsistenciaDiaria.fecha.desc()).limit(5).all()
    historico.reverse()
    labels_g = [h.fecha.strftime('%d/%m') for h in historico]
    datos_g = [h.porcentaje for h in historico]

    return render_template('index.html', fecha_full=fecha_hoy_esp, matricula_hoy=mat_hoy, 
                           asistentes_hoy=asist_hoy, porcentaje_hoy=porc_hoy,
                           total_v=total_v, total_h=total_h, anuncios=ultimos_anuncios,
                           labels_g=labels_g, datos_g=datos_g, mat_pers_hoy=mat_pers_hoy,
                           asist_pers_hoy=asist_pers_hoy, porc_pers_hoy=porc_pers_hoy,
                           desglose_personal=desglose_personal)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = Usuario.query.filter_by(username=request.form['username']).first()
        if u and check_password_hash(u.password, request.form['password']):
            session.update({
                'logeado': True, 'usuario_id': u.id, 'username': u.username, 
                'nombre_completo': u.nombre_completo, 'area_trabajo': u.area_trabajo,
                'rol_id': u.rol_id, 'foto_perfil_path': u.foto_perfil_path
            })
            if u.rol_info:
                session['permisos'], session['nombre_rol'] = u.rol_info.permisos, u.rol_info.nombre
                if u.rol_info.nombre in ['Obrero', 'Personal de Vigilancia', 'Personal de Cocina']:
                    return redirect(url_for('portal_trabajador'))
                return redirect(url_for('index'))
            session['permisos'], session['nombre_rol'] = '', 'En Espera'
            return redirect(url_for('en_espera'))
        return render_template('login.html', error='Credenciales incorrectas.')
    return render_template('login.html')

@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        u_name, e_mail = request.form['username'], request.form['email']
        if Usuario.query.filter((Usuario.username == u_name) | (Usuario.email == e_mail)).first():
            return render_template('registro.html', error="El usuario o el correo ya están registrados.")
        nuevo = Usuario(nombre_completo=request.form['nombre_completo'], username=u_name, email=e_mail,
                        area_trabajo=request.form['area_trabajo'],
                        password=generate_password_hash(request.form['password'], method='pbkdf2:sha256'))
        db.session.add(nuevo); db.session.commit()
        if Usuario.query.count() == 1:
            nuevo.rol_id = Rol.query.filter_by(nombre='Administrador Supremo').first().id
            db.session.commit()
        return redirect(url_for('login'))
    return render_template('registro.html')

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('index'))

@app.route('/en_espera')
def en_espera():
    return render_template('espera.html')

# ==========================================
# --- 4. PLANIFICADOR DOCENTE ---
# ==========================================

@app.route('/planificador')
def planificador():
    if not session.get('logeado'): return redirect(url_for('login'))
    uid = session['usuario_id']
    clases = Bitacora.query.filter_by(usuario_id=uid).order_by(Bitacora.fecha.desc()).all()
    grados = Grado.query.all(); temas = Tema.query.all()
    return render_template('planificador.html', registros=clases, grados=grados, temas=temas,
                           total=len(clases), completadas=len([c for c in clases if c.estado == 'Completado']),
                           pendientes=len([c for c in clases if c.estado == 'Pendiente']))

@app.route('/agregar', methods=['POST'])
def agregar():
    hora = request.form.get('hora', '07:00')
    f = datetime.strptime(f"{datetime.now().strftime('%Y-%m-%d')} {hora}", "%Y-%m-%d %H:%M")
    db.session.add(Bitacora(fecha=f, grado=request.form['grado'], actividad=request.form['actividad'], 
                            estado=request.form['estado'], usuario_id=session['usuario_id']))
    db.session.commit(); return redirect(url_for('planificador'))

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar(id):
    reg = Bitacora.query.get_or_404(id)
    if request.method == 'POST':
        reg.grado, reg.actividad, reg.estado = request.form['grado'], request.form['actividad'], request.form['estado']
        reg.fecha = datetime.strptime(f"{reg.fecha.strftime('%Y-%m-%d')} {request.form['hora']}", "%Y-%m-%d %H:%M")
        db.session.commit(); return redirect(url_for('planificador'))
    return render_template('editar.html', registro=reg, grados=Grado.query.all())

@app.route('/eliminar/<int:id>')
def eliminar(id):
    db.session.delete(Bitacora.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('planificador'))

@app.route('/reporte_diario')
def reporte_diario():
    if not session.get('logeado'): return redirect(url_for('login'))
    hoy = datetime.now().date()
    regs = [r for r in Bitacora.query.filter_by(usuario_id=session['usuario_id']).all() if r.fecha.date() == hoy]
    return render_template('reporte.html', registros=regs)

@app.route('/reporte_general')
def reporte_general():
    if not session.get('logeado'): return redirect(url_for('login'))
    regs = Bitacora.query.filter_by(usuario_id=session['usuario_id']).order_by(Bitacora.fecha.desc()).all()
    return render_template('reporte.html', registros=regs)

# ==========================================
# --- 5. EXPORTACIONES (EXCEL / WORD) ---
# ==========================================

@app.route('/exportar_excel')
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

@app.route('/exportar_word')
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
# --- 6. DEFENSORÍA ESTUDIANTIL ---
# ==========================================

@app.route('/defensoria', methods=['GET', 'POST'])
def defensoria():
    if not auth_defensoria():
        flash("Acceso denegado: No tienes permisos para ingresar a Defensoría.")
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        db.session.add(PlanificacionDefensoria(
            tema_charla=request.form['tema_charla'], proposito=request.form['proposito'],
            poblacion_objetivo=request.form['poblacion_objetivo'], usuario_id=session['usuario_id'],
            fecha_programada=datetime.strptime(request.form['fecha_programada'], '%Y-%m-%d').date(), 
            hora_programada=request.form['hora_programada']
        ))
        db.session.commit(); return redirect(url_for('defensoria'))
        
    planes = PlanificacionDefensoria.query.filter_by(usuario_id=session['usuario_id']).all()
    estudiantes = Estudiante.query.order_by(Estudiante.nombre_completo).all()
    brigadas = Brigada.query.all()
    actas = Acta.query.order_by(Acta.fecha.desc()).all()
    
    return render_template('defensoria.html', planificaciones=planes, estudiantes=estudiantes, brigadas=brigadas, actas=actas)

@app.route('/editar_defensoria/<int:id>', methods=['POST'])
def editar_defensoria(id):
    reg = PlanificacionDefensoria.query.get_or_404(id)
    reg.tema_charla, reg.proposito, reg.poblacion_objetivo = request.form['tema_charla'], request.form['proposito'], request.form['poblacion_objetivo']
    db.session.commit(); return redirect(url_for('defensoria'))

@app.route('/eliminar_defensoria/<int:id>', methods=['POST'])
def eliminar_defensoria(id):
    db.session.delete(PlanificacionDefensoria.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('defensoria'))

@app.route('/guardar_acta', methods=['POST'])
def guardar_acta():
    if not auth_defensoria(): return "Acceso Denegado", 403
    contenido = request.form.get('contenido_manual')
    tipo = request.form.get('tipo_acta', 'General')
    est_id = request.form.get('estudiante_id')
    estudiante_id = int(est_id) if est_id else None
    
    nueva_acta = Acta(
        contenido_manual=contenido,
        tipo_acta=tipo,
        estudiante_id=estudiante_id,
        defensor_id=session['usuario_id']
    )
    db.session.add(nueva_acta)
    db.session.commit()
    return redirect(url_for('defensoria'))

@app.route('/eliminar_acta/<int:id>', methods=['POST'])
def eliminar_acta(id):
    if not auth_defensoria(): return "Acceso Denegado", 403
    db.session.delete(Acta.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('defensoria'))

@app.route('/crear_brigada', methods=['POST'])
def crear_brigada():
    if not auth_defensoria(): return "Acceso Denegado", 403
    nombre = request.form.get('nombre')
    desc = request.form.get('descripcion')
    if nombre:
        db.session.add(Brigada(nombre=nombre, descripcion=desc))
        db.session.commit()
    return redirect(url_for('defensoria'))

@app.route('/asignar_brigada', methods=['POST'])
def asignar_brigada():
    if not auth_defensoria(): return "Acceso Denegado", 403
    estudiante_id = request.form.get('estudiante_id')
    brigada_id = request.form.get('brigada_id')
    
    if estudiante_id and brigada_id:
        estudiante = Estudiante.query.get(estudiante_id)
        brigada = Brigada.query.get(brigada_id)
        if estudiante and brigada and estudiante not in brigada.estudiantes:
            brigada.estudiantes.append(estudiante)
            db.session.commit()
    return redirect(url_for('defensoria'))

# --- EXPORTAR PDF DEFENSORÍA ---
@app.route('/exportar_pdf_defensoria')
def exportar_pdf_defensoria():
    if not session.get('logeado'): return redirect(url_for('login'))
    
    # Creamos el documento PDF
    pdf = FPDF()
    pdf.add_page()
    
    # Título del Documento
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Cronograma de Defensoria Estudiantil", ln=True, align='C')
    pdf.set_font("Arial", 'I', 12)
    pdf.cell(200, 10, txt=f"Docente: {session.get('nombre_completo')}", ln=True, align='C')
    pdf.ln(10)
    
    # Encabezados de la tabla
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(25, 10, 'Fecha', 1, 0, 'C')
    pdf.cell(25, 10, 'Hora', 1, 0, 'C')
    pdf.cell(40, 10, 'Grado/Seccion', 1, 0, 'C')
    pdf.cell(50, 10, 'Tema de Charla', 1, 0, 'C')
    pdf.cell(50, 10, 'Proposito', 1, 1, 'C')
    
    # Consultamos los datos del usuario activo
    planes = PlanificacionDefensoria.query.filter_by(usuario_id=session['usuario_id']).order_by(PlanificacionDefensoria.fecha_programada.asc()).all()
    
    # Llenamos la tabla
    pdf.set_font("Arial", '', 9)
    for p in planes:
        # Acortamos los textos si son muy largos para que no rompan la tabla
        tema = (p.tema_charla[:25] + '...') if len(p.tema_charla) > 25 else p.tema_charla
        proposito = (p.proposito[:25] + '...') if len(p.proposito) > 25 else p.proposito
        
        pdf.cell(25, 10, p.fecha_programada.strftime('%d/%m/%Y'), 1)
        pdf.cell(25, 10, p.hora_programada, 1)
        pdf.cell(40, 10, p.poblacion_objetivo, 1)
        pdf.cell(50, 10, tema, 1)
        pdf.cell(50, 10, proposito, 1)
        pdf.ln()
        
    # Preparamos la descarga
    response = make_response(pdf.output(dest='S').encode('latin-1'))
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'attachment; filename=Cronograma_Defensoria.pdf'
    return response    

# --- CARGA MASIVA EXCEL ---
@app.route('/descargar_plantilla')
def descargar_plantilla():
    output = io.BytesIO(); wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['TEMA_CHARLA', 'PROPOSITO_OBJETIVO']); wb.save(output); output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, download_name='Plantilla_Defensoria.xlsx')

@app.route('/subir_plan', methods=['POST'])
def subir_plan():
    archivo = request.files.get('archivo_plan')
    if not archivo: return redirect(url_for('defensoria'))
    wb = openpyxl.load_workbook(archivo); ws = wb.active
    extraidas = [{'tema': str(r[0]).strip(), 'proposito': str(r[1]).strip()} for r in ws.iter_rows(min_row=2, values_only=True) if r[0] and r[1]]
    return render_template('asignar_masivo.html', charlas=extraidas)

@app.route('/guardar_masivo', methods=['POST'])
def guardar_masivo():
    temas, propositos, grados, fechas, horas = request.form.getlist('tema[]'), request.form.getlist('proposito[]'), request.form.getlist('grado[]'), request.form.getlist('fecha[]'), request.form.getlist('hora[]')
    for i in range(len(temas)):
        if grados[i] and fechas[i] and horas[i]:
            db.session.add(PlanificacionDefensoria(tema_charla=temas[i], proposito=propositos[i], poblacion_objetivo=grados[i], 
                                                  fecha_programada=datetime.strptime(fechas[i], '%Y-%m-%d').date(), 
                                                  hora_programada=horas[i], usuario_id=session['usuario_id']))
    db.session.commit(); return redirect(url_for('defensoria'))

# ==========================================
# --- 7. CONTROL DE ASISTENCIA ---
# ==========================================

@app.route('/asistencia', methods=['GET', 'POST'])
def asistencia():
    if not session.get('logeado'): return redirect(url_for('login'))
    if request.method == 'POST':
        grado = Grado.query.get(request.form['grado_id'])
        v, h = int(request.form['varones'] or 0), int(request.form['hembras'] or 0)
        t = grado.total_varones + grado.total_hembras
        p = round(((v + h) / t) * 100, 1) if t > 0 else 0
        db.session.add(AsistenciaDiaria(fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(), 
                       grado_seccion=grado.nombre, matricula_total=t, varones=v, hembras=h, asistentes=v+h, 
                       porcentaje=p, usuario_id=session['usuario_id']))
        db.session.commit(); return redirect(url_for('asistencia'))
    if session.get('rol_id') in [1, 2]:
        regs = AsistenciaDiaria.query.order_by(AsistenciaDiaria.fecha.desc()).all()
    else:
        regs = AsistenciaDiaria.query.filter_by(usuario_id=session['usuario_id']).order_by(AsistenciaDiaria.fecha.desc()).all()
    return render_template('asistencia.html', registros=regs, grados=Grado.query.all(), hoy=datetime.now().strftime('%Y-%m-%d'))

@app.route('/eliminar_asistencia/<int:id>', methods=['POST'])
def eliminar_asistencia(id):
    db.session.delete(AsistenciaDiaria.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('asistencia'))

@app.route('/asistencia_personal', methods=['GET', 'POST'])
def asistencia_personal():
    if not session.get('logeado'): return redirect(url_for('login'))
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
        return redirect(url_for('asistencia_personal'))
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

@app.route('/eliminar_asistencia_personal/<int:id>')
def eliminar_asistencia_personal(id):
    if not session.get('logeado'): return redirect(url_for('login'))
    db.session.delete(AsistenciaPersonal.query.get_or_404(id))
    db.session.commit()
    return redirect(url_for('asistencia_personal'))

@app.route('/editar_asistencia_personal/<int:id>', methods=['GET', 'POST'])
def editar_asistencia_personal(id):
    if not session.get('logeado'): return redirect(url_for('login'))
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
        return redirect(url_for('asistencia_personal'))
    
    return render_template('editar_asistencia_personal.html', registro=reg)

# ==========================================
# --- 8. CONFIGURACIONES GLOBALES ---
# ==========================================

@app.route('/historial_global')
def historial_global():
    if not session.get('logeado'): return redirect(url_for('login'))
    if session.get('rol_id') not in [1, 2]:
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

@app.route('/configuracion', methods=['GET', 'POST'])
def configuracion():
    if not session.get('logeado'): return redirect(url_for('login'))
    if request.method == 'POST':
        if 'grado_num' in request.form and 'seccion_letra' in request.form:
            docente_id = request.form.get('docente_id')
            nombre_completo = f"{request.form.get('grado_num')} {request.form.get('seccion_letra')}"
            db.session.add(Grado(nombre=nombre_completo, 
                                 total_varones=int(request.form.get('m_varones') or 0),
                                 total_hembras=int(request.form.get('m_hembras') or 0), 
                                 usuario_id=docente_id))
        if 'nuevo_tema' in request.form:
            db.session.add(Tema(nombre=request.form['nuevo_tema'], usuario_id=session['usuario_id']))
        db.session.commit(); return redirect(url_for('configuracion'))
        
    docentes = Usuario.query.join(Rol).filter(Rol.nombre == 'Docente de Aula').all()
    return render_template('configuracion.html', grados=Grado.query.all(), temas=Tema.query.all(), docentes=docentes)

@app.route('/editar_grado/<int:id>', methods=['POST'])
def editar_grado(id):
    g = Grado.query.get_or_404(id)
    if 'grado_num' in request.form and 'seccion_letra' in request.form:
        g.nombre = f"{request.form.get('grado_num')} {request.form.get('seccion_letra')}"
    g.total_varones, g.total_hembras = int(request.form.get('m_varones', 0)), int(request.form.get('m_hembras', 0))
    if request.form.get('docente_id'):
        g.usuario_id = int(request.form.get('docente_id'))
    db.session.commit(); return redirect(url_for('configuracion'))

@app.route('/borrar_grado/<int:id>')
def borrar_grado(id):
    db.session.delete(Grado.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('configuracion'))

@app.route('/borrar_tema/<int:id>')
def borrar_tema(id):
    db.session.delete(Tema.query.get_or_404(id)); db.session.commit()
    return redirect(url_for('configuracion'))

# ==========================================
# --- 9. ADMIN Y ANUNCIOS ---
# ==========================================

@app.route('/admin_usuarios')
def admin_usuarios():
    if 'admin' not in session.get('permisos', ''): return "🚫 No autorizado."
    return render_template('admin.html', usuarios=Usuario.query.all(), roles=Rol.query.all())

@app.route('/cambiar_rol/<int:id>', methods=['POST'])
def cambiar_rol(id):
    if not session.get('logeado'): 
        return redirect(url_for('login'))
    
    usuario = Usuario.query.get_or_404(id)
    
    # 1. Protección del Administrador Supremo (Tú)
    if usuario.id == 1:
        flash("No puedes cambiar el rol del creador del sistema.", "error")
        return redirect(url_for('admin_usuarios'))
        
    # 2. Capturar el rol seleccionado en el menú desplegable
    nuevo_rol_id = request.form.get('rol')
    
    # Protección: Si el form está vacío o no eligieron nada, abortar
    if not nuevo_rol_id or nuevo_rol_id == '-- Elegir Rol --':
        flash("Por favor, selecciona un rol de la lista antes de autorizar.", "error")
        return redirect(url_for('admin_usuarios'))

    # 3. Guardar el nuevo rol en la base de datos (¡Muy importante!)
    usuario.rol_id = int(nuevo_rol_id)
    db.session.commit()
    
    rol_obj = Rol.query.get(usuario.rol_id)
    nombre_del_rol = rol_obj.nombre

    # 4. Enviar correo de notificación (Si no lo están devolviendo a Espera)
    if nombre_del_rol not in ['Espera', 'Pendiente']:
        try:
            msg = Message("¡Bienvenido a EduPlanner OS! - Acceso Aprobado",
                          sender="eepdanieloleary9@gmail.com",
                          recipients=[usuario.email]) 
            
            msg.html = f"""
            <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <div style="background-color: #4648d4; padding: 30px; text-align: center;">
                    <img src="https://scontent.fccs3-2.fna.fbcdn.net/v/t39.30808-6/655859736_122093139284858211_4065468023018454536_n.jpg?_nc_cat=100&ccb=1-7&_nc_sid=1d70fc&_nc_ohc=cWS418DldfAQ7kNvwGBlCkW&_nc_oc=AdosVWENtdz8RHtfJX0ahAPr28zsLb4xpgbrKs4w-h25ldNc1P83sINbQHP19QksMic&_nc_zt=23&_nc_ht=scontent.fccs3-2.fna&_nc_gid=ozBo8F8Ydo9C0PEr6jNyOw&_nc_ss=7a3a8&oh=00_Af0bu_2MBk-q9kOp_2eZ1pEZwX2ZRLb_V9XA7a17djR5-w&oe=69E19C90" alt="Logo Institucional" style="max-height: 80px; filter: drop-shadow(0px 2px 4px rgba(0,0,0,0.2));">
                </div>
                <div style="padding: 40px 30px; color: #333333; background-color: #ffffff;">
                    <h2 style="color: #111827; margin-top: 0;">¡Acceso Aprobado!</h2>
                    <p style="font-size: 16px; line-height: 1.6;">Hola <strong>{usuario.nombre_completo}</strong>,</p>
                    <p style="font-size: 16px; line-height: 1.6;">Nos complace informarte que tu cuenta ha sido verificada. Ya formas parte de <strong>EduPlanner OS</strong>.</p>
                    <div style="background-color: #f3f4f6; padding: 15px; border-radius: 8px; margin: 25px 0; text-align: center;">
                        <p style="margin: 0; font-size: 15px; color: #4b5563;">Rol asignado:</p>
                        <h3 style="margin: 5px 0 0 0; color: #4648d4; font-size: 20px;">{nombre_del_rol}</h3>
                    </div>
                    <div style="text-align: center; margin: 40px 0 20px 0;">
                        <a href="https://roboclass.pythonanywhere.com/login" style="background-color: #4648d4; color: #ffffff; padding: 14px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; display: inline-block;">Ingresar al Sistema</a>
                    </div>
                </div>
                <div style="background-color: #f9fafb; padding: 20px; text-align: center; font-size: 12px; color: #6b7280; border-top: 1px solid #e0e0e0;">
                    <p style="margin: 0;">Mensaje automático de EduPlanner OS.</p>
                </div>
            </div>
            """
            mail.send(msg)
            print("✔️ CORREO ENVIADO EXITOSAMENTE A:", usuario.email)
        except Exception as e:
            print("❌ ERROR ENVIANDO CORREO:", e)

    # 5. Mensaje de éxito en pantalla y recarga
    flash("Rol actualizado correctamente.", "success")
    return redirect(url_for('admin_usuarios'))

@app.route('/aprobar_cambio/<int:id>', methods=['POST'])
def aprobar_cambio(id):
    if 'admin' not in session.get('permisos', ''): return redirect(url_for('index'))
    usuario = Usuario.query.get_or_404(id)
    if not usuario.cargo_solicitado:
        flash('El usuario no tiene solicitudes pendientes.', 'error')
        return redirect(url_for('admin_usuarios'))
        
    MAPEO_CARGOS = {
        'Docente de Aula (1ro a 6to)': 'Docente de Aula',
        'Especialista (Robótica / Deportes)': 'Docente Especialista',
        'Defensoría Estudiantil': 'Defensoría Estudiantil',
        'Equipo Directivo / Administrativo': 'Equipo Directivo',
        'Obrero': 'Obrero',
        'Personal de Vigilancia': 'Personal de Vigilancia',
        'Personal de Cocina': 'Personal de Cocina'
    }
    
    nuevo_rol_nombre = MAPEO_CARGOS.get(usuario.cargo_solicitado)
    if not nuevo_rol_nombre:
        flash('Cargo solicitado no reconocido.', 'error')
        return redirect(url_for('admin_usuarios'))
        
    rol = Rol.query.filter_by(nombre=nuevo_rol_nombre).first()
    if not rol:
        flash(f'Rol {nuevo_rol_nombre} no existe en la base de datos.', 'error')
        return redirect(url_for('admin_usuarios'))
        
    usuario.area_trabajo = usuario.cargo_solicitado
    usuario.rol_id = rol.id
    usuario.cargo_solicitado = None
    db.session.commit()
    flash(f'Solicitud aprobada. Cargo de {usuario.nombre_completo} actualizado.', 'success')
    return redirect(url_for('admin_usuarios'))

@app.route('/rechazar_cambio/<int:id>', methods=['POST'])
def rechazar_cambio(id):
    if 'admin' not in session.get('permisos', ''): return redirect(url_for('index'))
    usuario = Usuario.query.get_or_404(id)
    usuario.cargo_solicitado = None
    db.session.commit()
    flash(f'Solicitud de cambio rechazada para {usuario.nombre_completo}.', 'success')
    return redirect(url_for('admin_usuarios'))

@app.route('/eliminar_usuario/<int:id>', methods=['POST'])
def eliminar_usuario(id):
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403
    if id != session['usuario_id']:
        usuario = Usuario.query.get_or_404(id)
        if usuario.id == 1:
            return redirect(url_for('admin_usuarios'))
        db.session.delete(usuario)
        db.session.commit()
    return redirect(url_for('admin_usuarios'))

@app.route('/anuncios', methods=['GET', 'POST'])
def anuncios():
    if not session.get('logeado'): return redirect(url_for('login'))
    if request.method == 'POST' and 'anuncios' in session.get('permisos', ''):
        db.session.add(Anuncio(titulo=request.form['titulo'], mensaje=request.form['mensaje'], autor_id=session['usuario_id']))
        db.session.commit(); return redirect(url_for('anuncios'))
    return render_template('anuncios.html', anuncios=Anuncio.query.order_by(Anuncio.fecha.desc()).all())

@app.route('/eliminar_anuncio/<int:id>', methods=['POST'])
def eliminar_anuncio(id):
    if 'admin' in session.get('permisos', ''):
        db.session.delete(Anuncio.query.get(id)); db.session.commit()
    return redirect(url_for('anuncios'))

# ==========================================
# --- 10. MÓDULO DE ESTADÍSTICA ---
# ==========================================

def generar_cedula_escolar(nro_parto, anio_nino, cedula_rep):
    # [Nro Parto] + [Últimos 2 dígitos Año del Niño] + [Cédula Madre/Representante]
    anio_str = str(anio_nino)[-2:]
    return f"{str(nro_parto)}{anio_str}{str(cedula_rep)}"

@app.route('/estadistica_global')
def estadistica_global():
    if not session.get('logeado'):
        return redirect(url_for('login'))
        
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

@app.route('/registrar_estudiante', methods=['POST'])
def registrar_estudiante():
    if not session.get('logeado'):
        return redirect(url_for('login'))

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
    
    db.session.add(estudiante)
    db.session.commit()
    
    return redirect(url_for('estadistica_global'))

@app.route('/perfil_estudiante/<int:id>')
def perfil_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('login'))
    estudiante = Estudiante.query.get_or_404(id)
    return render_template('perfil_estudiante.html', estudiante=estudiante, grados=Grado.query.all())

@app.route('/editar_estudiante/<int:id>', methods=['POST'])
def editar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('login'))
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
    return redirect(url_for('perfil_estudiante', id=est.id))

@app.route('/egresar_estudiante/<int:id>', methods=['POST'])
def egresar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('login'))
    est = Estudiante.query.get_or_404(id)
    est.estatus = 'Activo' if est.estatus == 'Egreso' else 'Egreso'
    db.session.commit()
    return redirect(url_for('perfil_estudiante', id=est.id))

@app.route('/eliminar_estudiante/<int:id>', methods=['POST'])
def eliminar_estudiante(id):
    if not session.get('logeado'): return redirect(url_for('login'))
    est = Estudiante.query.get_or_404(id)
    db.session.delete(est)
    db.session.commit()
    return redirect(url_for('estadistica_global'))

# ==========================================
# --- 11. MÓDULO MI AULA ---
# ==========================================

@app.route('/mi_aula', methods=['GET', 'POST'])
def mi_aula():
    if not session.get('logeado'): return redirect(url_for('login'))
    
    rol = session.get('nombre_rol')
    if rol not in ['Administrador Supremo', 'Equipo Directivo', 'Docente de Aula']:
        return redirect(url_for('index'))
    
    grados = []
    grado_seleccionado = None
    estudiantes = []
    asistencia_porcentaje = {}

    if rol == 'Docente de Aula':
        usuario_id = session.get('usuario_id')
        grado_seleccionado = Grado.query.filter_by(usuario_id=usuario_id).first()
        if grado_seleccionado:
            estudiantes = Estudiante.query.filter_by(grado_id=grado_seleccionado.id).order_by(Estudiante.nombre_completo.asc()).all()
    else:
        grados = Grado.query.all()
        grado_id = request.args.get('grado_id')
        if grado_id:
            grado_seleccionado = Grado.query.get(grado_id)
            if grado_seleccionado:
                estudiantes = Estudiante.query.filter_by(grado_id=grado_seleccionado.id).order_by(Estudiante.nombre_completo.asc()).all()

    for est in estudiantes:
        total_dias = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id).count() or 0
        asistencias = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id, asistio=True).count() or 0
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
            if inc.fecha.month == mes_actual and inc.categoria in datos_incidencias:
                datos_incidencias[inc.categoria] += 1
                
        treinta_dias_atras = date.today() - timedelta(days=30)
        asist_records = AsistenciaEstudiante.query.filter(
            AsistenciaEstudiante.estudiante_id.in_(estudiante_ids),
            AsistenciaEstudiante.fecha >= treinta_dias_atras
        ).all()
        
        if asist_records:
            total_dias_salon = len(asist_records)
            asistencias_positivas = sum(1 for a in asist_records if a.asistio)
            asistencia_promedio_salon = round((asistencias_positivas / total_dias_salon) * 100, 1) if total_dias_salon > 0 else 0.0
            
    total_varones = sum(1 for e in estudiantes if e.genero == 'Masculino')
    total_hembras = sum(1 for e in estudiantes if e.genero == 'Femenino')
    docente_titular = grado_seleccionado.docente_info.nombre_completo if (grado_seleccionado and grado_seleccionado.docente_info) else "Docente no asignado"

    return render_template('mi_aula.html', 
                           grado=grado_seleccionado, 
                           grados=grados, 
                           estudiantes=estudiantes, 
                           asistencia_porcentaje=asistencia_porcentaje,
                           total_matricula=total_matricula,
                           total_varones=total_varones,
                           total_hembras=total_hembras,
                           docente_titular=docente_titular,
                           datos_incidencias=datos_incidencias,
                           asistencia_promedio_salon=asistencia_promedio_salon)

@app.route('/guardar_asistencia_aula', methods=['POST'])
def guardar_asistencia_aula():
    if not session.get('logeado'): return redirect(url_for('login'))
    grado_id = request.form.get('grado_id')
    estudiantes = Estudiante.query.filter_by(grado_id=grado_id).all()
    fecha_hoy = date.today()
    
    for est in estudiantes:
        vino = request.form.get(f'asistio_{est.id}') == 'on'
        registro = AsistenciaEstudiante.query.filter_by(estudiante_id=est.id, fecha=fecha_hoy).first()
        if registro:
            registro.asistio = vino
        else:
            nuevo_registro = AsistenciaEstudiante(fecha=fecha_hoy, asistio=vino, estudiante_id=est.id)
            db.session.add(nuevo_registro)
            
    db.session.commit()
    url = url_for('mi_aula')
    if session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo'] and grado_id:
        url = url_for('mi_aula', grado_id=grado_id)
    return redirect(url)

@app.route('/agregar_incidencia', methods=['POST'])
def agregar_incidencia():
    if not session.get('logeado'): return redirect(url_for('login'))
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
        
    url = url_for('mi_aula')
    if session.get('nombre_rol') in ['Administrador Supremo', 'Equipo Directivo'] and grado_id:
        url = url_for('mi_aula', grado_id=grado_id)
    return redirect(url)

@app.route('/descargar_inscripcion_inicial/<int:grado_id>')
def descargar_inscripcion_inicial(grado_id):
    if not session.get('logeado'): return redirect(url_for('login'))
    
    grado = Grado.query.get_or_404(grado_id)
    estudiantes = Estudiante.query.filter_by(grado_id=grado.id).order_by(Estudiante.nombre_completo.asc()).all()
    
    docente_nombre = grado.docente_info.nombre_completo if grado.docente_info else "No asignado"
    
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

@app.route('/generar_enlace/<int:id>')
def generar_enlace(id):
    if not session.get('logeado'): return redirect(url_for('login'))
    
    import uuid
    estudiante = Estudiante.query.get_or_404(id)
    token = uuid.uuid4().hex
    
    nuevo_enlace = EnlaceTemporal(token=token, estudiante_id=estudiante.id, usado=False)
    db.session.add(nuevo_enlace)
    db.session.commit()
    
    url_acceso = url_for('ver_expediente', token=token, _external=True)
    mensaje = f"Aquí tienes el expediente temporal de {estudiante.nombre_completo}: {url_acceso}\n\n*Nota: Este enlace es de un solo uso y expirará después de abrirlo por razones de seguridad.*"
    
    whatsapp_url = f"https://wa.me/?text={urllib.parse.quote(mensaje)}"
    return redirect(whatsapp_url)

@app.route('/ver_expediente/<token>')
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
# --- 11. GESTIÓN DE PERSONAL ---
# ==========================================

@app.route('/gestion_personal')
def gestion_personal():
    if not session.get('logeado') or 'admin' not in session.get('permisos', []):
        return redirect(url_for('login'))
        
    personal = Usuario.query.all()
    
    # Calculate HR Document Compliance Metrics
    base_metric = lambda: {'docentes': 0, 'operativos': 0, 'directivos': 0, 'faltantes': 0}
    metrics = {
        'rif': base_metric(),
        'cv': base_metric(),
        'constancia': base_metric(),
        'voucher': base_metric(),
        'cedula': base_metric()
    }
    
    macro_docentes = ['Docente de Aula (1ro a 6to)', 'Especialista (Robótica / Deportes)', 'Defensoría Estudiantil']
    macro_directivos = ['Equipo Directivo / Administrativo', 'Administrador Supremo']
    macro_operativos = ['Obrero', 'Personal de Vigilancia', 'Personal de Cocina']
    
    for t in personal:
        area = t.area_trabajo or ''
        
        if area in macro_directivos:
            cat = 'directivos'
        elif area in macro_docentes:
            cat = 'docentes'
        elif area in macro_operativos:
            cat = 'operativos'
        else:
            cat = 'faltantes' # Fallback for unassigned or broken roles
            
        # If cat is faltantes, we still just attribute it to faltantes for charts, 
        # but to avoid breaking the dict, let's say they just add to missing unless they have it?
        # Actually if they don't belong to a category, they are outside the macros.
        # But we must count their docs. Let's just put them in 'docentes' as a safe default or ignore their loaded state?
        # Better: if cat is not identified, we just skip or put them in faltantes
        if cat != 'faltantes':
            if t.rif_path: metrics['rif'][cat] += 1
            else: metrics['rif']['faltantes'] += 1
            
            if t.curriculum_path: metrics['cv'][cat] += 1
            else: metrics['cv']['faltantes'] += 1
            
            if t.constancia_path: metrics['constancia'][cat] += 1
            else: metrics['constancia']['faltantes'] += 1
            
            if t.voucher_path: metrics['voucher'][cat] += 1
            else: metrics['voucher']['faltantes'] += 1
            
            if t.cedula_path: metrics['cedula'][cat] += 1
            else: metrics['cedula']['faltantes'] += 1
        else:
            # If they don't have a valid role, everything counts as faltante
            metrics['rif']['faltantes'] += 1
            metrics['cv']['faltantes'] += 1
            metrics['constancia']['faltantes'] += 1
            metrics['voucher']['faltantes'] += 1
            metrics['cedula']['faltantes'] += 1
            
    return render_template('gestion_personal.html', personal=personal, metrics=metrics)

@app.route('/actualizar_credenciales_mppe/<int:id>', methods=['POST'])
def actualizar_credenciales_mppe(id):
    if not session.get('logeado') or 'admin' not in session.get('permisos', []):
        return redirect(url_for('login'))
        
    usuario = Usuario.query.get_or_404(id)
    correo_mppe = request.form.get('correo_mppe')
    clave_mppe = request.form.get('clave_mppe')
    
    usuario.usuario_autogestion = correo_mppe
    usuario.clave_autogestion = clave_mppe
    
    db.session.commit()
    flash('Credenciales del MPPE actualizadas correctamente.', 'success')
    return redirect(url_for('gestion_personal'))

from werkzeug.utils import secure_filename

# ==========================================
# --- 12. PORTAL DEL TRABAJADOR ---
# ==========================================

@app.route('/portal_trabajador', methods=['GET', 'POST'])
def portal_trabajador():
    if not session.get('logeado'):
        return redirect(url_for('login'))
        
    usuario = Usuario.query.get(session['usuario_id'])
    
    if request.method == 'POST':
        tipo_documento = request.form.get('tipo_documento')
        if 'archivo' not in request.files:
            flash('No se seleccionó ningún archivo.', 'error')
            return redirect(url_for('portal_trabajador'))
            
        file = request.files['archivo']
        if file.filename == '':
            flash('No se seleccionó ningún archivo.', 'error')
            return redirect(url_for('portal_trabajador'))
            
        extension = file.filename.split('.')[-1].lower()
        if file and extension in ['pdf', 'png', 'jpg', 'jpeg']:
            # Patrón de nombre estandarizado: [tipo_documento]_[id_usuario].[extension]
            prefix = f"{tipo_documento}_{usuario.id}"
            
            # Buscar y eliminar cualquier archivo previo con el mismo prefijo
            for existing_file in os.listdir(app.config['UPLOAD_FOLDER']):
                if existing_file.startswith(prefix + "."):
                    old_path = os.path.join(app.config['UPLOAD_FOLDER'], existing_file)
                    if os.path.exists(old_path): 
                        os.remove(old_path)
            
            # Asignar nuevo nombre
            safe_filename = f"{prefix}.{extension}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
            file.save(file_path)
            
            if tipo_documento == 'voucher':
                usuario.voucher_path = safe_filename
            elif tipo_documento == 'constancia':
                usuario.constancia_path = safe_filename
            elif tipo_documento == 'curriculum':
                usuario.curriculum_path = safe_filename
            elif tipo_documento == 'rif':
                usuario.rif_path = safe_filename
            elif tipo_documento == 'cedula':
                usuario.cedula_path = safe_filename
            elif tipo_documento == 'foto_perfil':
                usuario.foto_perfil_path = safe_filename
                
            db.session.commit()
            flash('Documento guardado exitosamente en tu expediente.', 'success')
            return redirect(url_for('portal_trabajador'))
        else:
            flash('Tipo de archivo no permitido. Solo PDF, JPG o PNG.', 'error')
            return redirect(url_for('portal_trabajador'))

    return render_template('portal_trabajador.html', usuario=usuario)

from flask import send_from_directory

@app.route('/ver_documento_personal/<filename>')
def ver_documento_personal(filename):
    if not session.get('logeado'):
        return redirect(url_for('login'))
    # Verify permissions: must be admin or the owner of the document
    if 'admin' not in session.get('permisos', '') and not filename.startswith(f"{session.get('usuario_id')}_"):
        flash("Acceso denegado al documento.", "error")
        return redirect(url_for('index'))
        
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/mi_perfil', methods=['GET', 'POST'])
def mi_perfil():
    if not session.get('logeado'):
        return redirect(url_for('login'))
        
    usuario = Usuario.query.get(session['usuario_id'])
    
    if request.method == 'POST':
        nombre_completo = request.form.get('nombre_completo')
        area_trabajo = request.form.get('area_trabajo')
        cedula = request.form.get('cedula')
        file = request.files.get('foto_perfil')
        
        if nombre_completo:
            usuario.nombre_completo = nombre_completo
            session['nombre_completo'] = nombre_completo
            
        if cedula:
            usuario.cedula = cedula
            
        if area_trabajo and area_trabajo != usuario.area_trabajo and not usuario.cargo_solicitado:
            usuario.cargo_solicitado = area_trabajo
            
        if file and file.filename.split('.')[-1].lower() in ['png', 'jpg', 'jpeg']:
            import os
            extension = file.filename.split('.')[-1].lower()
            prefix = f"foto_perfil_{usuario.id}"
            # Replace existing file
            for existing_file in os.listdir(app.config['UPLOAD_FOLDER']):
                if existing_file.startswith(prefix + "."):
                    old_path = os.path.join(app.config['UPLOAD_FOLDER'], existing_file)
                    if os.path.exists(old_path): 
                        os.remove(old_path)
            
            safe_filename = f"{prefix}.{extension}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
            file.save(file_path)
            usuario.foto_perfil_path = safe_filename
            session['foto_perfil_path'] = safe_filename
            
        db.session.commit()
        flash('Perfil actualizado exitosamente.', 'success')
        return redirect(url_for('mi_perfil'))
        
    return render_template('mi_perfil.html', usuario=usuario)

if __name__ == '__main__':
    app.run(debug=True)