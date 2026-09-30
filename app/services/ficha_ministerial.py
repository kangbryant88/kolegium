"""
Ficha ministerial del personal (RAC Nominal + RAC Beneficios).

Una sola definición de cada campo, compartida por:
- el importador de Excel (qué formato aplicar a cada celda),
- el Tablero de Nómina (pestañas e inputs del modal "Editar Detalles"),
- la ruta que guarda esa edición (validación y formato).
Para agregar un campo nuevo: columna en Usuario + entrada en CAMPOS.
"""
import re
import unicodedata
from collections import namedtuple
from datetime import date, datetime

from app.web.perfil import formatear_nombre, formatear_cargo, formatear_codigo, formatear_telefono, OPCIONES_TURNO

PESTANAS = (
    ('personales', 'Datos Personales', 'user'),
    ('laborales', 'Datos Laborales', 'briefcase'),
    ('vivienda', 'Vivienda y Tallas', 'home'),
    ('salud', 'Salud', 'heart-pulse'),
)

# tipo: texto | area | telefono | fecha | horas | turno
Campo = namedtuple('Campo', 'etiqueta pestana tipo formato sugerencias')


# ==========================================
# --- FORMATEO DE VALORES ---
# ==========================================

def _clave(texto):
    """'Soltéra ' -> 'SOLTERA' (para reconocer códigos y variantes del Excel)."""
    texto = unicodedata.normalize('NFKD', texto or '').encode('ascii', 'ignore').decode()
    return ' '.join(re.sub(r'[^A-Z0-9]+', ' ', texto.upper()).split())


def formatear_oracion(texto):
    """Texto libre: "  HIPERTENSION   ARTERIAL" -> "Hipertension arterial"."""
    texto = ' '.join((texto or '').split()).lower()
    return texto[:1].upper() + texto[1:]


def formatear_mayusculas(texto):
    """Tallas y secciones: " xl " -> "XL"."""
    return ' '.join((texto or '').split()).upper()


def _por_codigos(codigos):
    """
    Formateador que traduce códigos del Excel a un valor estándar:
    codigos = ((('S', 'SOLTER'), 'Soltero(a)'), ...) -> una clave exacta o un
    prefijo reconocido da el valor; si no, Title Case del texto original.
    """
    def formatear(texto):
        clave = _clave(texto)
        if not clave:
            return ''
        for prefijos, valor in codigos:
            if any(clave == p or (len(p) > 2 and clave.startswith(p)) for p in prefijos):
                return valor
        return formatear_nombre(texto)
    return formatear


formatear_nacionalidad = _por_codigos((
    (('V', 'VEN'), 'Venezolana'),
    (('E', 'EXT'), 'Extranjera'),
))
formatear_estado_civil = _por_codigos((
    (('S', 'SOLTER'), 'Soltero(a)'),
    (('C', 'CASAD'), 'Casado(a)'),
    (('D', 'DIVORC'), 'Divorciado(a)'),
    (('V', 'VIUD'), 'Viudo(a)'),
    (('U', 'CONCUB', 'UNION'), 'Concubino(a)'),
))
formatear_nivel_instruccion = _por_codigos((
    (('TSU', 'T S U', 'TECNICO SUPERIOR'), 'Técnico Superior Universitario'),
    (('TECNICO MEDIO',), 'Técnico Medio'),
    (('BACH', 'BACHILLER'), 'Bachiller'),
    (('LIC', 'LICENCIAD'), 'Licenciado(a)'),
    (('ESP', 'ESPECIALISTA', 'ESPECIALIZACION'), 'Especialista'),
    (('MSC', 'MAGISTER', 'MAESTRIA'), 'Magíster'),
    (('DR', 'DRA', 'DOCTOR', 'DOCTORADO'), 'Doctor(a)'),
))
formatear_si_no = _por_codigos((
    (('S', 'SI'), 'Sí'),
    (('N', 'NO', 'NINGUNA', 'NINGUNO'), 'No'),
))


def formatear_horas(valor):
    """36.0 -> '36', 4.5 -> '4.5', None -> ''."""
    return '' if valor is None else f'{valor:g}'


# ==========================================
# --- DEFINICIÓN DE LOS CAMPOS ---
# ==========================================

CAMPOS = {
    # --- Datos Personales ---
    'nacionalidad': Campo('Nacionalidad', 'personales', 'texto', formatear_nacionalidad, ('Venezolana', 'Extranjera')),
    'lugar_nacimiento': Campo('Lugar de Nacimiento', 'personales', 'texto', formatear_nombre, ()),
    'estado_civil': Campo('Estado Civil', 'personales', 'texto', formatear_estado_civil,
                          ('Soltero(a)', 'Casado(a)', 'Divorciado(a)', 'Viudo(a)', 'Concubino(a)')),
    'nivel_instruccion': Campo('Nivel de Instrucción', 'personales', 'texto', formatear_nivel_instruccion,
                               ('Bachiller', 'Técnico Medio', 'Técnico Superior Universitario', 'Licenciado(a)',
                                'Profesor(a)', 'Especialista', 'Magíster', 'Doctor(a)')),
    'profesion': Campo('Profesión', 'personales', 'texto', formatear_nombre, ()),
    'telefono': Campo('Teléfono Celular', 'personales', 'telefono', formatear_telefono, ()),
    'telefono_habitacion': Campo('Teléfono de Habitación', 'personales', 'telefono', formatear_telefono, ()),
    'telefono_oficina': Campo('Teléfono de Oficina', 'personales', 'telefono', formatear_telefono, ()),

    # --- Datos Laborales ---
    'cargo': Campo('Cargo', 'laborales', 'texto', formatear_cargo, ()),  # sugerencias: cargos ya cargados
    'codigo_rac': Campo('Código RAC', 'laborales', 'texto', formatear_codigo, ()),
    'tipo_personal': Campo('Tipo de Personal', 'laborales', 'texto', formatear_nombre,
                           ('Docente', 'Administrativo', 'Obrero')),
    'fecha_ingreso': Campo('Fecha de Ingreso', 'laborales', 'fecha', None, ()),
    'turno': Campo('Turno', 'laborales', 'turno', None, OPCIONES_TURNO),
    'situacion_trabajador': Campo('Situación del Trabajador', 'laborales', 'texto', formatear_nombre,
                                  ('Activo', 'Reposo', 'Permiso', 'Vacaciones', 'Comisión De Servicio',
                                   'Jubilado', 'Suspendido')),
    'horas_academicas': Campo('Horas Académicas', 'laborales', 'horas', None, ()),
    'horas_adm': Campo('Horas Administrativas', 'laborales', 'horas', None, ()),
    'especialidad': Campo('Especialidad', 'laborales', 'texto', formatear_nombre, ()),
    'grado_imparte': Campo('Grado que Imparte', 'laborales', 'texto', formatear_nombre,
                           ('1er Grado', '2do Grado', '3er Grado', '4to Grado', '5to Grado', '6to Grado')),
    'seccion_imparte': Campo('Sección que Imparte', 'laborales', 'texto', formatear_mayusculas, ('A', 'B', 'C', 'D')),
    'observacion': Campo('Observación', 'laborales', 'area', formatear_oracion, ()),

    # --- Vivienda y Tallas (incluye las actividades del RAC Beneficios) ---
    'tipo_vivienda': Campo('Tipo de Vivienda', 'vivienda', 'texto', formatear_nombre,
                           ('Casa', 'Apartamento', 'Quinta', 'Anexo', 'Habitación', 'Rancho')),
    'condicion_vivienda': Campo('Condición de Vivienda', 'vivienda', 'texto', formatear_nombre,
                                ('Propia', 'Alquilada', 'Prestada', 'Familiar', 'Pagándose')),
    'tipo_material': Campo('Tipo de Material', 'vivienda', 'texto', formatear_nombre,
                           ('Bloque', 'Ladrillo', 'Madera', 'Zinc', 'Bahareque', 'Adobe')),
    'talla_camisa': Campo('Talla de Camisa', 'vivienda', 'texto', formatear_mayusculas,
                          ('XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL')),
    'talla_pantalon': Campo('Talla de Pantalón', 'vivienda', 'texto', formatear_mayusculas, ()),
    'talla_zapato': Campo('Talla de Zapato', 'vivienda', 'texto', formatear_mayusculas, ()),
    'actividad_deportiva': Campo('Actividad Deportiva', 'vivienda', 'texto', formatear_oracion, ('Ninguna',)),
    'actividad_cultural': Campo('Actividad Cultural', 'vivienda', 'texto', formatear_oracion, ('Ninguna',)),

    # --- Salud ---
    'tipo_enfermedad': Campo('Tipo de Enfermedad', 'salud', 'texto', formatear_oracion, ('Ninguna',)),
    'medicamento': Campo('Medicamento', 'salud', 'texto', formatear_oracion, ('Ninguno',)),
    'posee_discapacidad': Campo('Posee Discapacidad', 'salud', 'texto', formatear_si_no, ('No', 'Sí')),
}

HORAS_MAXIMAS = 80

# Formateo en vivo del modal (app/static/js/formato_perfil.js): data-formato de cada input
_FORMATO_EN_NAVEGADOR = {formatear_nombre: 'nombre', formatear_cargo: 'cargo',
                         formatear_codigo: 'codigo', formatear_telefono: 'telefono'}
FORMATOS_JS = {nombre: _FORMATO_EN_NAVEGADOR.get(campo.formato) for nombre, campo in CAMPOS.items()}


def nombre_para_mostrar(usuario):
    """Nombre legal en Title Case: "Nombres Apellidos" si están cargados, si no
    el nombre_completo. Nunca el usuario de login."""
    legal = ' '.join(filter(None, (usuario.nombres, usuario.apellidos)))
    return formatear_nombre(legal or usuario.nombre_completo or '') or 'Sin nombre'


def valor_para_formulario(usuario, nombre):
    """Valor de la BD tal como lo espera el input del modal."""
    valor = getattr(usuario, nombre)
    if valor is None:
        return ''
    tipo = CAMPOS[nombre].tipo
    if tipo == 'fecha':
        return valor.isoformat()
    if tipo == 'horas':
        return formatear_horas(valor)
    return str(valor)


def limpiar_valor(nombre, texto, max_largo=None):
    """
    Texto del formulario -> valor para guardar (None = vacío).
    Lanza ValueError con un mensaje para el usuario si no es válido.
    """
    campo = CAMPOS[nombre]
    texto = (texto or '').strip()
    if not texto:
        return None

    if campo.tipo == 'fecha':
        try:
            fecha = datetime.strptime(texto, '%Y-%m-%d').date()
        except ValueError:
            raise ValueError(f'{campo.etiqueta}: la fecha no es válida.')
        if fecha.year < 1900 or fecha > date.today():
            raise ValueError(f'{campo.etiqueta}: no puede ser futura ni anterior a 1900.')
        return fecha

    if campo.tipo == 'horas':
        try:
            horas = float(texto.replace(',', '.'))
        except ValueError:
            raise ValueError(f'{campo.etiqueta}: debe ser un número.')
        if not 0 <= horas <= HORAS_MAXIMAS:
            raise ValueError(f'{campo.etiqueta}: debe estar entre 0 y {HORAS_MAXIMAS}.')
        return horas

    if campo.tipo == 'turno':
        if texto not in OPCIONES_TURNO:
            raise ValueError(f'{campo.etiqueta}: la opción seleccionada no es válida.')
        return texto

    valor = campo.formato(texto)
    if campo.tipo == 'telefono' and not valor:
        raise ValueError(f'{campo.etiqueta}: no es válido. Formato esperado: 0414-1234567.')
    if valor and max_largo and len(valor) > max_largo:
        raise ValueError(f'{campo.etiqueta}: máximo {max_largo} caracteres.')
    return valor or None
