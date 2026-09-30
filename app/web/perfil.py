"""
Muro de Contención: obliga al personal a completar su perfil antes de usar
Kolegium, en dos niveles.

- Nivel 1 (Hard Required): datos personales básicos. Sin ellos no se entra
  a ninguna pantalla.
- Nivel 2 (Soft Required): datos ministeriales. Se insiste en cada inicio de
  sesión, pero el usuario puede posponerlos por el resto del día.
"""
import re
from datetime import date, datetime

from flask import Blueprint, render_template, request, redirect, url_for, session, flash

from app.models import db, Usuario

perfil_bp = Blueprint('perfil', __name__)

CAMPOS_NIVEL_1 = ('cedula', 'nombres', 'apellidos', 'fecha_nacimiento', 'sexo')
CAMPOS_NIVEL_2 = ('fecha_ingreso', 'cargo', 'codigo_rac', 'turno')

OPCIONES_SEXO = ('Masculino', 'Femenino')
OPCIONES_TURNO = ('Mañana', 'Tarde', 'Integral')

# Endpoints que nunca se bloquean: login/logout/recuperación, el propio muro,
# archivos estáticos y la foto de perfil que pinta la plantilla.
ENDPOINTS_LIBRES = {
    'static',
    'perfil.completar_perfil',
    'perfil.saltar_perfil',
    'ver_documento_personal',
}
# auth: login, registro, logout, en_espera, recuperar...; api/auth_api: la
# app móvil se autentica por token, no por sesión.
BLUEPRINTS_LIBRES = {'auth', 'api', 'auth_api'}


# ==========================================
# --- FORMATEO (tolerancia cero) ---
# ==========================================

def formatear_nombre(texto):
    """
    Title Case para nombres propios: "  maría  JOSÉ o'leary-pérez " ->
    "María José O'Leary-Pérez". Misma regla que `formatearNombre` en el JS
    de completar_perfil.html, para que el backend no dependa del navegador.
    """
    if not texto:
        return ''
    texto = ' '.join(texto.split()).lower()
    return re.sub(r"(^|[\s\-'])(\w)", lambda m: m.group(1) + m.group(2).upper(), texto)


def formatear_cargo(texto):
    """Title Case, pero con los niveles en romano: "docente iii" -> "Docente III"."""
    return re.sub(r'\b(Ii{0,2}|Iv|Vi{0,3}|Ix|X)\b', lambda m: m.group(1).upper(), formatear_nombre(texto))


def formatear_codigo(texto):
    """Códigos (RAC, etc.): mayúsculas y sin espacios."""
    return re.sub(r'\s+', '', texto or '').upper()


def formatear_telefono(texto):
    """
    Teléfono venezolano normalizado a "0414-1234567". Acepta "04141234567",
    "414 123 45 67", "+58 414-1234567". Devuelve None si no es válido.
    """
    digitos = re.sub(r'\D', '', texto or '')
    if len(digitos) == 12 and digitos.startswith('58'):
        digitos = '0' + digitos[2:]
    elif len(digitos) == 10 and not digitos.startswith('0'):
        digitos = '0' + digitos
    if not re.fullmatch(r'0[24]\d{9}', digitos):
        return None
    return f"{digitos[:4]}-{digitos[4:]}"


def formatear_cedula(texto):
    """
    Cédula venezolana normalizada a "V-12345678" / "E-12345678". Acepta
    "v12345678", "V 12.345.678", "12345678" (asume V). Devuelve None si
    no es una cédula válida.
    """
    limpio = re.sub(r'[\s.\-]', '', texto or '').upper()
    if limpio.isdigit():
        limpio = 'V' + limpio
    m = re.fullmatch(r'([VE])(\d{6,9})', limpio)
    if not m:
        return None
    return f"{m.group(1)}-{m.group(2)}"


def _parsear_fecha(valor):
    try:
        return datetime.strptime(valor, '%Y-%m-%d').date() if valor else None
    except ValueError:
        return None


# ==========================================
# --- ESTADO DEL PERFIL ---
# ==========================================

def nivel_1_completo(usuario):
    return all(getattr(usuario, c) for c in CAMPOS_NIVEL_1)


def nivel_2_completo(usuario):
    return all(getattr(usuario, c) for c in CAMPOS_NIVEL_2)


def _recordatorio_saltado_hoy():
    # Se guarda la fecha (no solo True) para que el pase libre caduque al
    # día siguiente aunque el navegador mantenga la sesión abierta.
    return session.get('saltar_recordatorio') == date.today().isoformat()


# ==========================================
# --- MIDDLEWARE ---
# ==========================================

@perfil_bp.before_app_request
def muro_de_contencion():
    if not session.get('logeado'):
        return None
    if request.endpoint is None or request.endpoint in ENDPOINTS_LIBRES:
        return None
    if request.blueprint in BLUEPRINTS_LIBRES:
        return None

    usuario = db.session.get(Usuario, session.get('usuario_id'))
    if usuario is None:
        session.clear()
        return redirect(url_for('auth.login'))

    # Nivel 1 incompleto: bloqueo total, sin excepciones.
    if not nivel_1_completo(usuario):
        return redirect(url_for('perfil.completar_perfil'))

    # Nivel 2 incompleto: se insiste, salvo que hoy ya haya pedido el pase.
    if not nivel_2_completo(usuario) and not _recordatorio_saltado_hoy():
        return redirect(url_for('perfil.completar_perfil'))

    return None


# ==========================================
# --- RUTAS ---
# ==========================================

@perfil_bp.route('/completar_perfil', methods=['GET', 'POST'])
def completar_perfil():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))

    usuario = db.session.get(Usuario, session['usuario_id'])
    if usuario is None:
        session.clear()
        return redirect(url_for('auth.login'))

    errores = []
    if request.method == 'POST':
        f = request.form

        # --- Nivel 1: obligatorio ---
        cedula = formatear_cedula(f.get('cedula'))
        nombres = formatear_nombre(f.get('nombres'))
        apellidos = formatear_nombre(f.get('apellidos'))
        fecha_nac = _parsear_fecha(f.get('fecha_nacimiento'))
        sexo = f.get('sexo', '')

        if not cedula:
            errores.append('La cédula no es válida. Formato esperado: V-12345678 o E-12345678.')
        elif Usuario.query.filter(Usuario.cedula == cedula, Usuario.id != usuario.id).first():
            errores.append('Esa cédula ya está registrada por otro usuario.')
        if not nombres:
            errores.append('Los nombres son obligatorios.')
        if not apellidos:
            errores.append('Los apellidos son obligatorios.')
        if not fecha_nac or fecha_nac >= date.today() or fecha_nac.year < 1920:
            errores.append('La fecha de nacimiento no es válida.')
        if sexo not in OPCIONES_SEXO:
            errores.append('Selecciona el sexo.')

        # --- Nivel 2: opcional, pero si viene, debe venir bien ---
        fecha_ing_raw = f.get('fecha_ingreso', '').strip()
        fecha_ing = _parsear_fecha(fecha_ing_raw)
        cargo = formatear_cargo(f.get('cargo'))
        codigo_rac = formatear_codigo(f.get('codigo_rac'))
        turno = f.get('turno', '')

        if fecha_ing_raw and (not fecha_ing or fecha_ing > date.today()):
            errores.append('La fecha de ingreso no es válida.')
        if turno and turno not in OPCIONES_TURNO:
            errores.append('El turno seleccionado no es válido.')

        if errores:
            for e in errores:
                flash(e, 'error')
        else:
            usuario.cedula = cedula
            usuario.nombres = nombres
            usuario.apellidos = apellidos
            usuario.fecha_nacimiento = fecha_nac
            usuario.sexo = sexo
            usuario.nombre_completo = f"{nombres} {apellidos}"
            session['nombre_completo'] = usuario.nombre_completo

            # Nivel 2: un campo vacío no borra lo que ya estaba guardado.
            if fecha_ing:
                usuario.fecha_ingreso = fecha_ing
            if cargo:
                usuario.cargo = cargo
            if codigo_rac:
                usuario.codigo_rac = codigo_rac
            if turno:
                usuario.turno = turno

            db.session.commit()

            if nivel_2_completo(usuario):
                session.pop('saltar_recordatorio', None)
                flash('¡Perfil completo! Gracias por mantener tus datos al día.', 'success')
                return redirect(url_for('index'))

            flash('Datos básicos guardados. Aún faltan tus datos ministeriales.', 'warning')
            return redirect(url_for('perfil.completar_perfil'))

    return render_template(
        'completar_perfil.html',
        usuario=usuario,
        form=request.form if errores else None,
        nivel_1_ok=nivel_1_completo(usuario),
        opciones_sexo=OPCIONES_SEXO,
        opciones_turno=OPCIONES_TURNO,
        hoy_iso=date.today().isoformat(),
    )


@perfil_bp.route('/saltar_perfil')
def saltar_perfil():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))

    # El pase libre solo existe para el Nivel 2.
    usuario = db.session.get(Usuario, session['usuario_id'])
    if usuario is None or not nivel_1_completo(usuario):
        flash('Primero debes completar tus datos personales básicos.', 'error')
        return redirect(url_for('perfil.completar_perfil'))

    session['saltar_recordatorio'] = date.today().isoformat()
    return redirect(url_for('index'))
