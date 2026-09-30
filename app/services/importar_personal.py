"""
Importación masiva de personal desde Excel (RAC NOMINAL, Beneficios, etc.).

- Detecta sola la hoja y la fila de cabecera (RAC NOMINAL la tiene en la
  fila 0, Beneficios en la fila 4): busca la primera fila con una columna
  de cédula.
- Reconoce las columnas por alias (sin importar acentos ni mayúsculas).
- Upsert: si el usuario existe solo rellena sus campos VACÍOS; si no existe
  lo crea con usuario y contraseña = dígitos de la cédula, sin rol
  (queda "En Espera" hasta que el administrador le asigne uno).
"""
import re
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
from werkzeug.security import generate_password_hash

from app.models import db, Usuario
from app.web.perfil import formatear_nombre, formatear_cargo, formatear_codigo, formatear_cedula, formatear_telefono
from app.services.ficha_ministerial import CAMPOS as CAMPOS_FICHA, HORAS_MAXIMAS

FILAS_A_ESCANEAR = 30  # filas donde se busca la cabecera real

# Alias de cada columna. Se escriben como en el Excel (con o sin acentos):
# al cargar el módulo pasan por _normalizar_cabecera, igual que las cabeceras.
ALIAS_COLUMNAS = {
    'cedula': {'CEDULA', 'CEDULA DE IDENTIDAD', 'CEDULA IDENTIDAD', 'C I', 'CI', 'NRO CEDULA',
               'N CEDULA', 'NO CEDULA', 'NUMERO DE CEDULA', 'CEDULA N', 'C I DEL TRABAJADOR',
               'CI DEL TRABAJADOR', 'C I TRABAJADOR', 'CI TRABAJADOR', 'IDENTIFICACION',
               'DOCUMENTO DE IDENTIDAD', 'NO DE CEDULA', 'N DE CEDULA', 'NRO DE CEDULA'},
    'nacionalidad': {'NAC', 'NACIONALIDAD', 'V E'},
    'nombres': {'NOMBRES', 'NOMBRE'},
    'primer_nombre': {'PRIMER NOMBRE', '1ER NOMBRE'},
    'segundo_nombre': {'SEGUNDO NOMBRE', '2DO NOMBRE'},
    'apellidos': {'APELLIDOS', 'APELLIDO'},
    'primer_apellido': {'PRIMER APELLIDO', '1ER APELLIDO'},
    'segundo_apellido': {'SEGUNDO APELLIDO', '2DO APELLIDO'},
    'nombre_completo': {'APELLIDOS Y NOMBRES', 'NOMBRES Y APELLIDOS', 'NOMBRE COMPLETO',
                        'NOMBRE Y APELLIDO', 'APELLIDO Y NOMBRE', 'APELLIDOS NOMBRES',
                        'NOMBRES APELLIDOS', 'TRABAJADOR', 'NOMBRE DEL TRABAJADOR'},
    'fecha_nacimiento': {'FECHA DE NACIMIENTO', 'FECHA NACIMIENTO', 'F NACIMIENTO', 'FECHA NAC',
                         'F NAC', 'FEC NAC', 'NACIMIENTO'},
    'sexo': {'SEXO', 'GENERO'},
    'fecha_ingreso': {'FECHA DE INGRESO', 'FECHA INGRESO', 'F INGRESO', 'FEC INGRESO', 'INGRESO',
                      'FECHA DE INGRESO AL MPPE', 'FECHA INGRESO MPPE', 'FECHA DE INGRESO MPPE',
                      # Errores tipográficos reales de los Excel del ministerio
                      # ('FECHA DE INGRES0' con cero llega aquí ya corregido a 'O')
                      'FECHA DE INGRES', 'FECHA INGRES', 'FECHA DE INGRESOS'},
    'cargo': {'CARGO', 'DENOMINACION DEL CARGO', 'DENOMINACION CARGO', 'DENOMINACION',
              'CARGO NOMINAL', 'DESCRIPCION DEL CARGO', 'DESCRIPCION CARGO'},
    'codigo_rac': {'CODIGO RAC', 'COD RAC', 'RAC', 'CODIGO DEL RAC', 'N RAC', 'NRO RAC'},
    'turno': {'TURNO', 'TURNO QUE ATIENDE', 'TURNO ATIENDE', 'TURNO DE TRABAJO', 'TURNO LABORAL'},
    'email': {'CORREO', 'CORREO ELECTRONICO', 'EMAIL', 'E MAIL'},
    'telefono': {'TELEFONO', 'TELF', 'TLF', 'TELEFONO CELULAR', 'CELULAR', 'TELEFONO MOVIL', 'MOVIL',
                 'NRO TELEFONO', 'N TELEFONO', 'TELEFONO DE CONTACTO', 'TELEFONO CONTACTO',
                 'TLF CELULAR', 'TELF CELULAR', 'TLF MOVIL', 'TELF MOVIL'},
    # --- Ficha completa: RAC Nominal ---
    'tipo_personal': {'TIPO PERSONAL', 'TIPO DE PERSONAL', 'TIPO TRABAJADOR', 'TIPO DE TRABAJADOR'},
    'horas_academicas': {'HORAS ACADEMICAS', 'HORAS ACAD', 'HRS ACADEMICAS', 'HRS ACAD', 'H ACADEMICAS',
                         'HORAS DOCENTES', 'HORAS DE AULA'},
    'horas_adm': {'HORAS ADM', 'HORAS ADMINISTRATIVAS', 'HRS ADM', 'HRS ADMINISTRATIVAS', 'H ADM',
                  'H ADMINISTRATIVAS', 'HORAS ADMINISTRATIVA'},
    'grado_imparte': {'GRADO', 'GRADO QUE IMPARTE', 'GRADO IMPARTE', 'GRADO QUE ATIENDE', 'ANO GRADO',
                      'GRADO ANO', 'GRADO O ANO'},
    'seccion_imparte': {'SECCION', 'SECCION QUE IMPARTE', 'SECCION IMPARTE', 'SECCION QUE ATIENDE'},
    'especialidad': {'ESPECIALIDAD', 'MENCION', 'AREA DE ESPECIALIDAD'},
    'situacion_trabajador': {'SITUACION DEL TRABAJADOR', 'SITUACION TRABAJADOR', 'SITUACION',
                             'SITUACION ACTUAL', 'ESTATUS', 'ESTATUS DEL TRABAJADOR', 'ESTADO DEL TRABAJADOR'},
    'observacion': {'OBSERVACION', 'OBSERVACIONES', 'OBS'},
    # --- Ficha completa: datos personales ---
    'lugar_nacimiento': {'LUGAR DE NACIMIENTO', 'LUGAR NACIMIENTO', 'LUGAR DE NAC', 'LUGAR NAC'},
    'estado_civil': {'ESTADO CIVIL', 'EDO CIVIL', 'E CIVIL'},
    'nivel_instruccion': {'NIVEL DE INSTRUCCION', 'NIVEL INSTRUCCION', 'GRADO DE INSTRUCCION',
                          'NIVEL ACADEMICO', 'INSTRUCCION'},
    'profesion': {'PROFESION', 'PROFESION U OFICIO', 'TITULO', 'TITULO OBTENIDO'},
    'telefono_habitacion': {'TELEFONO DE HABITACION', 'TELEFONO HABITACION', 'TELF HABITACION',
                            'TLF HABITACION', 'TLF HAB', 'TELF HAB', 'TELEFONO HAB', 'TELEFONO LOCAL',
                            'TELEFONO FIJO'},
    'telefono_oficina': {'TELEFONO DE OFICINA', 'TELEFONO OFICINA', 'TELF OFICINA', 'TLF OFICINA',
                         'TLF OFIC', 'TELF OFIC', 'TELEFONO OFIC', 'TELEFONO DE TRABAJO'},
    # --- Ficha completa: RAC Beneficios ---
    'talla_camisa': {'TALLA DE CAMISA', 'TALLA CAMISA', 'CAMISA', 'TALLA DE CHEMISE', 'TALLA CHEMISE'},
    'talla_pantalon': {'TALLA DE PANTALON', 'TALLA PANTALON', 'PANTALON'},
    'talla_zapato': {'TALLA DE ZAPATO', 'TALLA ZAPATO', 'ZAPATO', 'TALLA DE CALZADO', 'TALLA CALZADO', 'CALZADO'},
    'actividad_deportiva': {'ACTIVIDAD DEPORTIVA', 'ACTIVIDADES DEPORTIVAS', 'DEPORTE', 'DEPORTES',
                            'PRACTICA ALGUN DEPORTE'},
    'actividad_cultural': {'ACTIVIDAD CULTURAL', 'ACTIVIDADES CULTURALES', 'CULTURA'},
    'tipo_vivienda': {'TIPO DE VIVIENDA', 'TIPO VIVIENDA', 'VIVIENDA'},
    'condicion_vivienda': {'CONDICION DE VIVIENDA', 'CONDICION VIVIENDA', 'CONDICION DE LA VIVIENDA',
                           'TENENCIA DE VIVIENDA', 'TENENCIA DE LA VIVIENDA'},
    'tipo_material': {'TIPO DE MATERIAL', 'TIPO MATERIAL', 'MATERIAL', 'MATERIAL DE CONSTRUCCION',
                      'MATERIAL DE LA VIVIENDA', 'TIPO DE MATERIAL DE LA VIVIENDA'},
    'tipo_enfermedad': {'TIPO DE ENFERMEDAD', 'TIPO ENFERMEDAD', 'ENFERMEDAD', 'ENFERMEDADES',
                        'PADECE ALGUNA ENFERMEDAD'},
    'medicamento': {'MEDICAMENTO', 'MEDICAMENTOS', 'TRATAMIENTO', 'MEDICAMENTO QUE CONSUME'},
    'posee_discapacidad': {'POSEE DISCAPACIDAD', 'POSEE ALGUNA DISCAPACIDAD', 'DISCAPACIDAD',
                           'TIPO DE DISCAPACIDAD'},
}

# Si ningún alias coincide exacto, se intenta por palabras contenidas en la cabecera.
REGLAS_RESPALDO = {
    'cedula': lambda h: ('CEDULA' in h.split() or h.startswith(('C I ', 'CI '))) and 'ESCOLAR' not in h
                        and 'REPRESENTANTE' not in h,
    'codigo_rac': lambda h: 'RAC' in h.split() and 'NOMINAL' not in h,
    'fecha_ingreso': lambda h: 'INGRES' in h and 'FECHA' in h,  # INGRESO / INGRES (truncado)
    'turno': lambda h: 'TURNO' in h.split(),
    'fecha_nacimiento': lambda h: 'NACIMIENTO' in h and 'LUGAR' not in h,
    'cargo': lambda h: 'CARGO' in h.split() and 'CODIGO' not in h,
    'email': lambda h: 'CORREO' in h,
    # Los teléfonos fijos van antes que el celular para que este no se los lleve
    'telefono_habitacion': lambda h: ('TELEFONO' in h or 'TLF' in h or 'TELF' in h) and 'HAB' in h,
    'telefono_oficina': lambda h: ('TELEFONO' in h or 'TLF' in h or 'TELF' in h) and 'OFIC' in h,
    'telefono': lambda h: ('TELEFONO' in h or 'CELULAR' in h) and 'HAB' not in h and 'OFIC' not in h,
    'lugar_nacimiento': lambda h: 'LUGAR' in h and 'NAC' in h,
    'estado_civil': lambda h: 'CIVIL' in h,
    'nivel_instruccion': lambda h: 'INSTRUCCION' in h,
    'profesion': lambda h: 'PROFESION' in h,
    'tipo_personal': lambda h: 'TIPO' in h and 'PERSONAL' in h,
    'horas_academicas': lambda h: ('HORAS' in h or 'HRS' in h.split()) and 'ACAD' in h,
    'horas_adm': lambda h: ('HORAS' in h or 'HRS' in h.split()) and 'ADM' in h,
    'situacion_trabajador': lambda h: 'SITUACION' in h,
    'especialidad': lambda h: 'ESPECIALIDAD' in h,
    'observacion': lambda h: 'OBSERVACION' in h,
    'talla_camisa': lambda h: 'CAMISA' in h or 'CHEMISE' in h,
    'talla_pantalon': lambda h: 'PANTALON' in h,
    'talla_zapato': lambda h: 'ZAPATO' in h or 'CALZADO' in h,
    'actividad_deportiva': lambda h: 'DEPORT' in h,
    'actividad_cultural': lambda h: 'CULTURA' in h,
    'condicion_vivienda': lambda h: 'VIVIENDA' in h and ('CONDICION' in h or 'TENENCIA' in h),
    'tipo_material': lambda h: 'MATERIAL' in h,
    'tipo_vivienda': lambda h: 'VIVIENDA' in h and 'TIPO' in h,
    'tipo_enfermedad': lambda h: 'ENFERMEDAD' in h,
    'medicamento': lambda h: 'MEDICAMENTO' in h,
    'posee_discapacidad': lambda h: 'DISCAPACIDAD' in h,
}

# Cargo del Excel -> area_trabajo de Kolegium (solo para registros nuevos)
AREAS_POR_CARGO = (
    (('COCIN', 'PROCESADOR', 'MADRE ELABORADORA'), 'Personal de Cocina'),
    (('VIGIL', 'SEGURIDAD'), 'Personal de Vigilancia'),
    (('OBRER', 'ASEADOR', 'MANTENIMIENTO', 'LIMPIEZA', 'JARDINER', 'PORTER'), 'Obrero'),
    (('DIRECTOR', 'SUBDIRECTOR', 'COORDINADOR'), 'Equipo Directivo (Dirección)'),
    (('DOCENTE', 'MAESTR', 'PROFESOR'), 'Docente de Aula (1ro a 6to)'),
    (('SECRETARI', 'ADMINISTRATIV', 'ASISTENTE', 'AUXILIAR', 'OFICINISTA'), 'Administrativo'),
)
AREA_POR_DEFECTO = 'Por Asignar'

# Campos que se rellenan si están vacíos: Muro de Contención nivel 1 + la
# ficha ministerial completa (RAC Nominal y Beneficios)
CAMPOS_ACTUALIZABLES = ('cedula', 'nombres', 'apellidos', 'fecha_nacimiento', 'sexo') + tuple(CAMPOS_FICHA)


# ==========================================
# --- NORMALIZACIÓN DE VALORES ---
# ==========================================

def _normalizar_cabecera(valor):
    """
    Cabecera del Excel -> clave comparable con ALIAS_COLUMNAS:
    '  Fecha de Ingres0\xa0' -> 'FECHA DE INGRESO'.
    - Quita acentos y caracteres invisibles (espacio duro, espacio de ancho
      cero, tabulaciones, saltos de línea) y recorta los extremos.
    - Mayúsculas y un solo espacio entre palabras; la puntuación cuenta
      como espacio ('TLF. HAB.' -> 'TLF HAB').
    - Un cero pegado a letras es una 'O' mal tecleada ('INGRES0', 'TELEF0NO').
    """
    texto = unicodedata.normalize('NFKD', str(valor)).encode('ascii', 'ignore').decode()
    texto = re.sub(r'[^A-Z0-9]+', ' ', texto.strip().upper())
    texto = re.sub(r'(?<=[A-Z])0|0(?=[A-Z])', 'O', texto)
    return ' '.join(texto.split())


def _vacio(valor):
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return str(valor).strip() == ''


def _texto(valor):
    """Celda -> str limpio. 12345678.0 -> '12345678'."""
    if _vacio(valor):
        return ''
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return ' '.join(str(valor).split())


def _cedula(valor, nacionalidad=''):
    texto = _texto(valor)
    if texto.isdigit() and nacionalidad[:1].upper() in ('V', 'E'):
        texto = nacionalidad[:1].upper() + texto
    return formatear_cedula(texto)


def _es_anio(numero):
    return 1900 <= numero <= date.today().year


def _fecha(valor):
    """
    Timestamp / datetime / serial de Excel / texto -> date (o None).
    Limpieza de los formatos sucios del ministerio:
    - '01 04 2022', '01 / 04 / 2022', '01.04.2022' -> separadores unificados a '/'.
    - Solo el año (2022 o '2022') -> 01/01/2022. Un número entre 1900 y el año
      actual es un año, no un serial de Excel (2022 como serial sería 1905).
    - Se descarta la hora: '2022-04-01 00:00:00' -> 2022-04-01.
    """
    if _vacio(valor):
        return None
    resultado = None
    if isinstance(valor, datetime):
        resultado = valor.date()
    elif isinstance(valor, date):
        resultado = valor
    elif isinstance(valor, (int, float)):
        if float(valor).is_integer() and _es_anio(int(valor)):
            resultado = date(int(valor), 1, 1)
        elif 1000 < valor < 80000:  # número de serie de Excel
            resultado = (datetime(1899, 12, 30) + timedelta(days=int(valor))).date()
    else:
        texto = re.sub(r'[\sT]+\d{1,2}:\d{2}.*$', '', _texto(valor))       # quitar la hora
        texto = re.sub(r'\s*[/.\-]\s*|\s+', '/', texto.strip())             # espacios y separadores -> '/'
        if re.fullmatch(r'\d{4}', texto) and _es_anio(int(texto)):
            resultado = date(int(texto), 1, 1)
        for formato in ('%d/%m/%Y', '%Y/%m/%d', '%d/%m/%y'):
            if resultado:
                break
            try:
                resultado = datetime.strptime(texto, formato).date()
            except ValueError:
                continue
    if resultado and 1900 <= resultado.year <= date.today().year:
        return resultado
    return None


# Número guardado como texto: '4.24328395e+09', '4,24328395E+09', '4243283950.0'
_NUMERO_COMO_TEXTO = re.compile(r'\d+(?:[.,]\d+)?[eE][+-]?\d+|\d+[.,]0+')


def _telefono(valor):
    """
    Celda de teléfono -> '0424-3283950' (o None). Pandas entrega los números
    como float (4243283950.0 o 4.24328395e+09) y pierde el 0 inicial:
    float -> int -> str, y si quedan 10 dígitos se antepone el 0.
    """
    if _vacio(valor):
        return None
    if isinstance(valor, str) and _NUMERO_COMO_TEXTO.fullmatch(valor.strip()):
        texto = valor.strip().replace(',', '.')
        mantisa, _, exponente = texto.lower().partition('e')
        # Texto científico recortado ('4.14555E+09'): los dígitos reales ya se
        # perdieron en el Excel; mejor vacío que un teléfono inventado con ceros.
        if exponente and len(re.sub(r'\D', '', mantisa)) < int(exponente) + 1:
            return None
        valor = float(texto)
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        valor = str(int(valor))
        if len(valor) == 10:
            valor = '0' + valor
    return formatear_telefono(_texto(valor))


def _sexo(valor):
    texto = _normalizar_cabecera(_texto(valor))
    if texto in ('F', 'FEM', 'FEMENINO', 'MUJER'):
        return 'Femenino'
    if texto in ('M', 'MAS', 'MASC', 'MASCULINO', 'HOMBRE'):
        return 'Masculino'
    return None


def _turno(valor):
    texto = _normalizar_cabecera(_texto(valor))
    if not texto:
        return None
    if 'INTEGRAL' in texto or ('MANANA' in texto and 'TARDE' in texto):
        return 'Integral'
    if texto.startswith('M') or 'MANANA' in texto:
        return 'Mañana'
    if texto.startswith('T') or 'TARDE' in texto:
        return 'Tarde'
    return None


def _horas(valor):
    """36 / 36.0 / '36,5' -> float (o None si no es un número razonable)."""
    if _vacio(valor):
        return None
    try:
        horas = float(valor) if isinstance(valor, (int, float)) else float(_texto(valor).replace(',', '.'))
    except ValueError:
        return None
    return horas if 0 <= horas <= HORAS_MAXIMAS else None


def _area_por_cargo(cargo):
    texto = _normalizar_cabecera(cargo)
    for claves, area in AREAS_POR_CARGO:
        if any(c in texto for c in claves):
            return area
    return AREA_POR_DEFECTO


# ==========================================
# --- LECTURA DEL EXCEL ---
# ==========================================

# Las claves de los alias se limpian con la misma función que las cabeceras
# del Excel: un alias escrito con acentos ('Cédula') nunca deja de coincidir.
ALIAS_NORMALIZADOS = {campo: {_normalizar_cabecera(a) for a in alias} for campo, alias in ALIAS_COLUMNAS.items()}


def _mapear_columnas(cabeceras):
    """{campo: índice de columna} a partir de la fila de cabecera."""
    normalizadas = [_normalizar_cabecera(h) if not _vacio(h) else '' for h in cabeceras]
    mapa = {}
    for campo, alias in ALIAS_NORMALIZADOS.items():
        for i, h in enumerate(normalizadas):
            if h in alias and i not in mapa.values():
                mapa[campo] = i
                break
    for campo, regla in REGLAS_RESPALDO.items():
        if campo in mapa:
            continue
        for i, h in enumerate(normalizadas):
            if h and i not in mapa.values() and regla(h):
                mapa[campo] = i
                break
    return mapa


def leer_excel(archivo):
    """
    Devuelve (nombre_hoja, fila_cabecera, mapa_columnas, cabeceras, filas).
    Lanza ValueError si ninguna hoja tiene una columna de cédula.
    """
    hojas = pd.read_excel(archivo, sheet_name=None, header=None, dtype=object, engine='openpyxl')
    candidatas = []  # por hoja, la fila con más celdas de texto: para explicar el error
    for nombre_hoja, df in hojas.items():
        mejor = (0, None, [])
        for fila_cabecera in range(min(FILAS_A_ESCANEAR, len(df))):
            cabeceras = list(df.iloc[fila_cabecera])
            mapa = _mapear_columnas(cabeceras)
            if 'cedula' in mapa:
                filas = df.iloc[fila_cabecera + 1:].values.tolist()
                return nombre_hoja, fila_cabecera, mapa, cabeceras, filas
            textos = [_texto(h) for h in cabeceras if isinstance(h, str) and h.strip()]
            if len(textos) > mejor[0]:
                mejor = (len(textos), fila_cabecera, textos)
        if mejor[1] is not None:
            vistas = ', '.join(f'"{t}"' for t in mejor[2][:12]) + (' …' if len(mejor[2]) > 12 else '')
            candidatas.append(f'hoja "{nombre_hoja}", fila {mejor[1] + 1}: {vistas}')
    raise ValueError('No se encontró una columna "Cédula" en las primeras '
                     f'{FILAS_A_ESCANEAR} filas de ninguna hoja del archivo. '
                     + ('Encabezados encontrados: ' + ' | '.join(candidatas) if candidatas else 'El archivo parece vacío.'))


def _datos_de_fila(fila, mapa):
    def celda(campo):
        return fila[mapa[campo]] if campo in mapa and mapa[campo] < len(fila) else None

    nombres = _texto(celda('nombres')) or ' '.join(
        filter(None, (_texto(celda('primer_nombre')), _texto(celda('segundo_nombre')))))
    apellidos = _texto(celda('apellidos')) or ' '.join(
        filter(None, (_texto(celda('primer_apellido')), _texto(celda('segundo_apellido')))))
    nombres, apellidos = formatear_nombre(nombres), formatear_nombre(apellidos)
    nombre_completo = (f'{nombres} {apellidos}'.strip()
                       or formatear_nombre(_texto(celda('nombre_completo'))))

    datos = {
        'cedula': _cedula(celda('cedula'), _texto(celda('nacionalidad'))),
        'nombres': nombres or None,
        'apellidos': apellidos or None,
        'nombre_completo': nombre_completo,
        'fecha_nacimiento': _fecha(celda('fecha_nacimiento')),
        'sexo': _sexo(celda('sexo')),
        'fecha_ingreso': _fecha(celda('fecha_ingreso')),
        'cargo': formatear_cargo(_texto(celda('cargo'))) or None,
        'codigo_rac': formatear_codigo(_texto(celda('codigo_rac'))) or None,
        'turno': _turno(celda('turno')),
        'email': _texto(celda('email')).lower() or None,
        'telefono': _telefono(celda('telefono')),
    }
    # Resto de la ficha ministerial: cada campo con su formato (ficha_ministerial.py)
    for nombre, campo in CAMPOS_FICHA.items():
        if nombre in datos:
            continue
        if campo.tipo == 'horas':
            datos[nombre] = _horas(celda(nombre))
        elif campo.tipo == 'telefono':
            datos[nombre] = _telefono(celda(nombre))
        else:
            datos[nombre] = campo.formato(_texto(celda(nombre))) or None
    return datos


# ==========================================
# --- UPSERT ---
# ==========================================

def _digitos(cedula):
    return re.sub(r'\D', '', cedula or '')


# Partículas que no identifican a nadie: no cuentan para validar un match parcial
PARTICULAS_NOMBRE = {'DE', 'DEL', 'LA', 'LAS', 'LOS', 'Y', 'DA', 'DI'}
MIN_PALABRAS_MATCH_PARCIAL = 2  # "Daissy" solo es demasiado ambiguo


def _palabras(nombre):
    """'Daissy Mercedes Reyna Bohórquez' -> {'DAISSY', 'MERCEDES', 'REYNA', 'BOHORQUEZ'}
    (sin acentos ni mayúsculas/minúsculas)."""
    return set(_normalizar_cabecera(nombre or '').split())


def buscar_usuario_existente(datos, digitos, por_cedula, por_email, sin_cedula):
    """
    Busca al usuario de una fila del Excel. Devuelve (usuario | None, vínculo).
    Orden: 1) cédula  2) correo  3) nombre completo exacto  4) match parcial.
    Los pasos 3 y 4 solo miran usuarios que aún NO tienen cédula y solo
    aceptan el resultado si hay un único candidato.
    """
    usuario = por_cedula.get(digitos)
    if usuario:
        return usuario, ''

    usuario = por_email.get(datos['email']) if datos['email'] else None
    if usuario and (not usuario.cedula or _digitos(usuario.cedula) == digitos):
        return usuario, 'vinculado por correo'

    nombre_excel = datos['nombre_completo']
    if not nombre_excel:
        return None, ''

    # 3) Nombre completo exacto
    nombre_norm = _normalizar_cabecera(nombre_excel)
    exactos = [u for u in sin_cedula if _normalizar_cabecera(u.nombre_completo) == nombre_norm]
    if len(exactos) == 1:
        return exactos[0], 'vinculado por nombre'
    if len(exactos) > 1:
        return None, ''  # homónimos: mejor crear/revisar que mezclar personas

    # 4) Match parcial: TODAS las palabras del nombre registrado ('Daissy Reyna')
    #    aparecen en el nombre legal del Excel ('Daissy Mercedes Reyna Bohorquez').
    palabras_excel = _palabras(nombre_excel)
    parciales = []
    for u in sin_cedula:
        palabras_registro = _palabras(u.nombre_completo)
        if len(palabras_registro - PARTICULAS_NOMBRE) < MIN_PALABRAS_MATCH_PARCIAL:
            continue
        if palabras_registro <= palabras_excel:
            parciales.append(u)
    if len(parciales) == 1:
        return parciales[0], f'vinculado por nombre parcial "{parciales[0].nombre_completo}"'

    return None, ''


def importar_personal(archivo, simular=False):
    """Procesa el Excel y hace el upsert. Con simular=True no guarda nada."""
    nombre_hoja, fila_cabecera, mapa, cabeceras, filas = leer_excel(archivo)

    usuarios = Usuario.query.all()
    por_cedula = {_digitos(u.cedula): u for u in usuarios if _digitos(u.cedula)}
    por_email = {u.email.lower(): u for u in usuarios if u.email}
    # Usuarios que aún no llenaron su cédula: candidatos a vincular por nombre
    sin_cedula = [u for u in usuarios if not u.cedula and u.nombre_completo]
    usernames = {u.username.lower() for u in usuarios}
    emails = set(por_email)

    resultado = {'hoja': nombre_hoja, 'fila_cabecera': fila_cabecera, 'simulacion': simular,
                 'columnas': {campo: _texto(cabeceras[i]) for campo, i in mapa.items()},
                 'actualizados': 0, 'creados': 0, 'sin_cambios': 0, 'omitidos': 0, 'detalle': []}
    vistos = set()

    def anotar(n_fila, estado, nombre, cedula, nota=''):
        resultado['detalle'].append({'fila': n_fila, 'estado': estado, 'nombre': nombre,
                                     'cedula': cedula or '', 'nota': nota})

    for i, fila in enumerate(filas):
        n_fila = fila_cabecera + i + 2  # número de fila tal como se ve en Excel
        if all(_vacio(v) for v in fila):
            continue
        try:
            datos = _datos_de_fila(fila, mapa)
        except Exception as e:  # celda con formato inesperado: se omite la fila, no el archivo
            resultado['omitidos'] += 1
            anotar(n_fila, 'Omitido', '', '', f'Error leyendo la fila: {e}')
            continue

        cedula, digitos = datos['cedula'], _digitos(datos['cedula'])
        if not cedula:
            resultado['omitidos'] += 1
            anotar(n_fila, 'Omitido', datos['nombre_completo'], _texto(fila[mapa['cedula']]),
                   'Cédula vacía o inválida')
            continue
        if digitos in vistos:
            resultado['omitidos'] += 1
            anotar(n_fila, 'Omitido', datos['nombre_completo'], cedula, 'Cédula repetida en el archivo')
            continue
        vistos.add(digitos)

        usuario, vinculo = buscar_usuario_existente(datos, digitos, por_cedula, por_email, sin_cedula)
        if usuario in sin_cedula:
            # Ya quedó vinculado a esta cédula: no puede reclamarlo otra fila
            sin_cedula.remove(usuario)

        if usuario:
            cambios = [c for c in CAMPOS_ACTUALIZABLES
                       if not getattr(usuario, c) and datos.get(c)]
            for c in cambios:
                setattr(usuario, c, datos[c])
            if not usuario.nombre_completo and datos['nombre_completo']:
                usuario.nombre_completo = datos['nombre_completo']
            por_cedula[digitos] = usuario
            if cambios:
                resultado['actualizados'] += 1
                anotar(n_fila, 'Actualizado', usuario.nombre_completo, cedula,
                       ', '.join(cambios) + (f' ({vinculo})' if vinculo else ''))
            else:
                resultado['sin_cambios'] += 1
                anotar(n_fila, 'Sin cambios', usuario.nombre_completo, cedula, 'Ya tenía los datos completos')
            continue

        # --- Nuevo registro ---
        username = digitos
        sufijo = 1
        while username in usernames:
            sufijo += 1
            username = f'{digitos}_{sufijo}'
        usernames.add(username)

        email = datos['email'] if datos['email'] and datos['email'] not in emails else f'{username}@sin-correo.kolegium'
        emails.add(email)

        nuevo = Usuario(
            username=username,
            nombre_completo=datos['nombre_completo'] or f'Personal {cedula}',
            email=email,
            area_trabajo=_area_por_cargo(datos['cargo'] or ''),
            password=generate_password_hash(digitos, method='pbkdf2:sha256'),
            activo=True,
            **{c: datos[c] for c in CAMPOS_ACTUALIZABLES},
        )
        db.session.add(nuevo)
        por_cedula[digitos] = nuevo
        resultado['creados'] += 1
        anotar(n_fila, 'Creado', nuevo.nombre_completo, cedula,
               f'Usuario: {username} · Área: {nuevo.area_trabajo}')

    if simular:
        db.session.rollback()
    else:
        db.session.commit()
    return resultado
