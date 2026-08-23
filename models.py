from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date

db = SQLAlchemy()

# ==========================================
# --- MODELOS DE BASE DE DATOS ---
# ==========================================

class Rol(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(50), unique=True, nullable=False)
    permisos = db.Column(db.String(500), nullable=False, default="")
    usuarios = db.relationship('Usuario', backref='rol_info', lazy=True)

class ConfiguracionInstitucional(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre_escuela = db.Column(db.String(150), nullable=True)
    director = db.Column(db.String(100), nullable=True)
    telefono_director = db.Column(db.String(50), nullable=True)
    correo_director = db.Column(db.String(100), nullable=True)
    codigo_estadistico = db.Column(db.String(50), nullable=True)
    codigo_dea = db.Column(db.String(50), nullable=True)
    codigo_administrativo = db.Column(db.String(50), nullable=True)
    codigo_dependencia = db.Column(db.String(50), nullable=True)
    codigo_sunagro = db.Column(db.String(50), nullable=True)
    rif_escuela = db.Column(db.String(50), nullable=True)
    rif_consejo = db.Column(db.String(50), nullable=True)
    dependencia = db.Column(db.String(50), nullable=True)
    ubicacion_geografica = db.Column(db.String(50), nullable=True)
    clase_plantel = db.Column(db.String(50), nullable=True)
    ano_fundacion = db.Column(db.String(10), nullable=True)
    telefono_escuela = db.Column(db.String(50), nullable=True)
    correo_escuela = db.Column(db.String(100), nullable=True)
    supervisora = db.Column(db.String(100), nullable=True)
    direccion = db.Column(db.Text, nullable=True)
    circuito = db.Column(db.String(50), nullable=True)

class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    nombre_completo = db.Column(db.String(100), nullable=False) 
    email = db.Column(db.String(120), unique=True, nullable=False)
    area_trabajo = db.Column(db.String(100), nullable=False)
    password = db.Column(db.String(200), nullable=False)
    rol_id = db.Column(db.Integer, db.ForeignKey('rol.id'), nullable=True)
    departamento_asignado = db.Column(db.String(50), nullable=True)
    activo = db.Column(db.Boolean, default=True)
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
    usuario = db.relationship('Usuario', backref='bitacoras')

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

grado_docente = db.Table('grado_docente',
    db.Column('grado_id', db.Integer, db.ForeignKey('grado.id'), primary_key=True),
    db.Column('usuario_id', db.Integer, db.ForeignKey('usuario.id'), primary_key=True)
)

class Grado(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(50), nullable=False)
    total_varones = db.Column(db.Integer, default=0)
    total_hembras = db.Column(db.Integer, default=0)
    docentes = db.relationship('Usuario', secondary=grado_docente, lazy='subquery',
        backref=db.backref('grados_asignados', lazy=True))

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
    # Datos Laborales (Ficha de Admisión)
    ocupacion = db.Column(db.String(100))
    lugar_direccion_trabajo = db.Column(db.Text)
    # Datos Bancarios (Ficha de Admisión)
    banco_nombre = db.Column(db.String(100))
    banco_cuenta_numero = db.Column(db.String(30))
    banco_cuenta_tipo = db.Column(db.String(20))
    banco_titular_nombre = db.Column(db.String(150))
    banco_titular_ci = db.Column(db.String(20))
    estudiantes = db.relationship('Estudiante', backref='representante_info', lazy=True)

class Estudiante(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    # Identidad y Legal
    cedula_escolar = db.Column(db.String(30), unique=True, nullable=False)
    nombre_completo = db.Column(db.String(150), nullable=False)
    nombres = db.Column(db.String(100))
    apellidos = db.Column(db.String(100))
    fecha_nacimiento = db.Column(db.Date, nullable=False)
    lugar_nacimiento = db.Column(db.String(100))
    municipio = db.Column(db.String(100))
    parroquia = db.Column(db.String(100))
    edad = db.Column(db.Integer)
    genero = db.Column(db.String(20))
    nacionalidad = db.Column(db.String(50), default='Venezolana')
    num_acta = db.Column(db.String(50))
    num_oficio = db.Column(db.String(50))
    # Contacto y Ubicación del Alumno
    direccion_alumno = db.Column(db.Text) # Dirección con punto de referencia
    telefono_habitacion = db.Column(db.String(20))
    posee_canaima = db.Column(db.Boolean, default=False)
    # Salud y Nutrición
    talla = db.Column(db.Float) # En cm (Estatura)
    peso = db.Column(db.Float) # En kg
    calzado = db.Column(db.String(10))
    talla_camisa = db.Column(db.String(50))
    talla_pantalon = db.Column(db.String(50))
    tipaje = db.Column(db.String(10)) # Grupo sanguíneo
    vacunacion_completa = db.Column(db.String(20)) # Si/No/Parcial
    alergias = db.Column(db.String(200))
    neurodivergencia = db.Column(db.Boolean, default=False) # Diversidad Funcional
    neuro_detalle = db.Column(db.String(200))
    posee_enfermedad = db.Column(db.Boolean, default=False)
    enfermedad_detalle = db.Column(db.String(200))
    toma_medicamento = db.Column(db.Boolean, default=False)
    medicamento_detalle = db.Column(db.String(200))
    alergico_medicamento = db.Column(db.Boolean, default=False)
    alergia_medicamento_detalle = db.Column(db.String(200))
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
    # Datos de la Madre
    madre_nombre = db.Column(db.String(150))
    madre_ci = db.Column(db.String(20))
    madre_ocupacion = db.Column(db.String(100))
    madre_telefono = db.Column(db.String(20))
    madre_direccion = db.Column(db.Text)
    # Datos del Padre
    padre_nombre = db.Column(db.String(150))
    padre_ci = db.Column(db.String(20))
    padre_ocupacion = db.Column(db.String(100))
    padre_telefono = db.Column(db.String(20))
    padre_direccion = db.Column(db.Text)
    # Otros y Autorizaciones
    telefono_familiar_extra = db.Column(db.String(20))
    autorizacion_odontologica = db.Column(db.Boolean, default=False)
    # Relaciones
    grado_id = db.Column(db.Integer, db.ForeignKey('grado.id'))
    grado = db.relationship('Grado', backref='estudiantes')
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
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    grado_id = db.Column(db.Integer, db.ForeignKey('grado.id'), nullable=False)
    fecha = db.Column(db.Date, nullable=False, default=date.today)
    estatus = db.Column(db.String(20), nullable=False, default='Presente')  # Presente, Ausente, Justificado
    # Relaciones
    estudiante = db.relationship('Estudiante', backref='asistencias_detalle')
    grado = db.relationship('Grado', backref='asistencias_estudiantes')

class AlertaDefensoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    fecha_emision = db.Column(db.Date, nullable=False, default=date.today)
    motivo = db.Column(db.String(300), nullable=False)
    estatus_atencion = db.Column(db.String(30), nullable=False, default='Pendiente')  # Pendiente, Contactado, Visita Domiciliaria
    semana_iso = db.Column(db.String(10), nullable=True)  # Para identificar la semana (ej: '2026-W31')
    # Relaciones
    estudiante = db.relationship('Estudiante', backref='alertas_defensoria')


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
    # Relaciones
    estudiante = db.relationship('Estudiante', backref='actas')
    defensor = db.relationship('Usuario', backref='actas_creadas')

class SolicitudEnlace(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    docente_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    motivo = db.Column(db.Text, nullable=False)
    estado = db.Column(db.String(50), default='Pendiente')
    fecha_solicitud = db.Column(db.DateTime, default=datetime.utcnow)
    token_generado = db.Column(db.String(100), nullable=True)

    docente = db.relationship('Usuario', backref='solicitudes_enlace')
    estudiante = db.relationship('Estudiante', backref='solicitudes_enlace')

class SolicitudActualizacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    docente_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    
    rep_telefono = db.Column(db.String(50))
    rep_direccion = db.Column(db.Text)
    alergias = db.Column(db.String(200))
    neuro_detalle = db.Column(db.String(200))
    
    estado = db.Column(db.String(50), default='Pendiente')
    fecha_solicitud = db.Column(db.DateTime, default=datetime.utcnow)
    
    docente = db.relationship('Usuario', backref='solicitudes_actualizacion')
    estudiante = db.relationship('Estudiante', backref='solicitudes_actualizacion')

class SolicitudDefensoria(db.Model):
    __tablename__ = 'solicitudes_defensoria'
    
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    solicitante_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False) 
    
    motivo = db.Column(db.Text, nullable=False)
    fecha_solicitud = db.Column(db.DateTime, default=datetime.utcnow)
    
    estado = db.Column(db.String(20), default='Pendiente', nullable=False)
    fecha_respuesta = db.Column(db.DateTime, nullable=True)
    
    estudiante = db.relationship('Estudiante', backref='solicitudes_visita')
    solicitante = db.relationship('Usuario', backref='solicitudes_defensoria_realizadas')

class TokenRecuperacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    token = db.Column(db.String(100), unique=True, nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    fecha_creacion = db.Column(db.DateTime, default=datetime.now)
    usado = db.Column(db.Boolean, default=False)
    usuario = db.relationship('Usuario', backref='tokens_recuperacion')

# ==========================================
# --- EVALUACIÓN DESCRIPTIVA Y BOLETINES ---
# ==========================================

class ProyectoAula(db.Model):
    """Configuración global del Proyecto de Aprendizaje por salón y momento (lapso)."""
    id = db.Column(db.Integer, primary_key=True)
    grado_id = db.Column(db.Integer, db.ForeignKey('grado.id'), nullable=False)
    momento = db.Column(db.Integer, nullable=False)  # 1, 2 o 3
    titulo_proyecto = db.Column(db.String(200), nullable=False)
    fecha_inicio = db.Column(db.Date, nullable=True)
    fecha_cierre = db.Column(db.Date, nullable=True)
    fecha_actualizacion = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    grado = db.relationship('Grado', backref='proyectos_aula')

    __table_args__ = (
        db.UniqueConstraint('grado_id', 'momento', name='uq_proyecto_aula_grado_momento'),
    )


class BancoIndicador(db.Model):
    """Repositorio de indicadores/frases pre-guardadas por el docente, por momento y nivel de logro."""
    id = db.Column(db.Integer, primary_key=True)
    docente_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    momento = db.Column(db.Integer, nullable=False)  # 1, 2 o 3
    nivel = db.Column(db.String(20), nullable=False)  # A, B, C, D, Deporte
    texto_indicador = db.Column(db.Text, nullable=False)
    fecha_creacion = db.Column(db.DateTime, default=datetime.now)

    docente = db.relationship('Usuario', backref='banco_indicadores')


class EvaluacionEstudiante(db.Model):
    """Evaluación descriptiva consolidada (nota final) de un estudiante para un momento dado."""
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    momento = db.Column(db.Integer, nullable=False)  # 1, 2 o 3
    texto_descriptivo = db.Column(db.Text)
    sugerencias = db.Column(db.Text)
    fecha_actualizacion = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    estudiante = db.relationship('Estudiante', backref='evaluaciones')

    __table_args__ = (
        db.UniqueConstraint('estudiante_id', 'momento', name='uq_evaluacion_estudiante_momento'),
    )
