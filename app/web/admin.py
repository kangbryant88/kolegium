import re
from collections import Counter
from datetime import date

from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app
from flask_mail import Message
from sqlalchemy.exc import IntegrityError
from werkzeug.security import generate_password_hash
from app.models import db, Usuario, Rol
from app.web.perfil import formatear_cedula, formatear_nombre
from app.services.ficha_ministerial import (CAMPOS as CAMPOS_FICHA, PESTANAS as PESTANAS_FICHA,
                                            limpiar_valor, valor_para_formulario, formatear_horas,
                                            FORMATOS_JS)
from app.services.importar_personal import importar_personal as importar_personal_desde_excel

# Importamos 'mail' desde extensions (patrón Factory)
from app.extensions import mail

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

@admin_bp.route('/usuarios')
def admin_usuarios():
    if 'admin' not in session.get('permisos', ''): return "🚫 No autorizado."
    return render_template('admin.html', usuarios=Usuario.query.all(), roles=Rol.query.all())

@admin_bp.route('/cambiar_rol/<int:id>', methods=['POST'])
def cambiar_rol(id):
    if not session.get('logeado'): 
        return redirect(url_for('auth.login'))
    
    usuario = Usuario.query.get_or_404(id)
    
    # 1. Protección del Administrador Supremo (Tú)
    if usuario.id == 1:
        flash("No puedes cambiar el rol del creador del sistema.", "error")
        return redirect(url_for('admin.admin_usuarios'))
        
    # 2. Capturar el rol seleccionado en el menú desplegable
    nuevo_rol_id = request.form.get('rol')
    
    # Protección: Si el form está vacío o no eligieron nada, abortar
    if not nuevo_rol_id or nuevo_rol_id == '-- Elegir Rol --':
        flash("Por favor, selecciona un rol de la lista antes de autorizar.", "error")
        return redirect(url_for('admin.admin_usuarios'))

    # 3. Guardar el nuevo rol en la base de datos (¡Muy importante!)
    usuario.rol_id = int(nuevo_rol_id)
    db.session.commit()
    
    rol_obj = Rol.query.get(usuario.rol_id)
    nombre_del_rol = rol_obj.nombre

    # Auto-actualizar el área de trabajo para que coincida con el rol
    MAPEO_REVERSO_CARGOS = {
        'Docente de Aula': 'Docente de Aula (1ro a 6to)',
        'Docente Especialista': 'Especialista (Robótica / Deportes)',
        'Defensoría Estudiantil': 'Defensoría Estudiantil',
        'Equipo Directivo (Dirección)': 'Equipo Directivo (Dirección)',
        'Administrativo': 'Administrativo',
        'Administrador Supremo': 'Administrador Supremo',
        'Obrero': 'Obrero',
        'Personal de Vigilancia': 'Personal de Vigilancia',
        'Personal de Cocina': 'Personal de Cocina'
    }
    
    if nombre_del_rol in MAPEO_REVERSO_CARGOS:
        usuario.area_trabajo = MAPEO_REVERSO_CARGOS[nombre_del_rol]
        db.session.commit()

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
    return redirect(url_for('admin.admin_usuarios'))

@admin_bp.route('/aprobar_cambio/<int:id>', methods=['POST'])
def aprobar_cambio(id):
    if 'admin' not in session.get('permisos', ''): return redirect(url_for('index'))
    usuario = Usuario.query.get_or_404(id)
    if not usuario.cargo_solicitado:
        flash('El usuario no tiene solicitudes pendientes.', 'error')
        return redirect(url_for('admin.admin_usuarios'))
        
    MAPEO_CARGOS = {
        'Docente de Aula (1ro a 6to)': 'Docente de Aula',
        'Especialista (Robótica / Deportes)': 'Docente Especialista',
        'Defensoría Estudiantil': 'Defensoría Estudiantil',
        'Equipo Directivo (Dirección)': 'Equipo Directivo (Dirección)',
        'Administrativo': 'Administrativo',
        'Obrero': 'Obrero',
        'Personal de Vigilancia': 'Personal de Vigilancia',
        'Personal de Cocina': 'Personal de Cocina'
    }
    
    nuevo_rol_nombre = MAPEO_CARGOS.get(usuario.cargo_solicitado)
    if not nuevo_rol_nombre:
        flash('Cargo solicitado no reconocido.', 'error')
        return redirect(url_for('admin.admin_usuarios'))
        
    rol = Rol.query.filter_by(nombre=nuevo_rol_nombre).first()
    if not rol:
        flash(f'Rol {nuevo_rol_nombre} no existe en la base de datos.', 'error')
        return redirect(url_for('admin.admin_usuarios'))
        
    usuario.area_trabajo = usuario.cargo_solicitado
    usuario.rol_id = rol.id
    usuario.cargo_solicitado = None
    db.session.commit()
    flash(f'Solicitud aprobada. Cargo de {usuario.nombre_completo} actualizado.', 'success')
    return redirect(url_for('admin.admin_usuarios'))

@admin_bp.route('/rechazar_cambio/<int:id>', methods=['POST'])
def rechazar_cambio(id):
    if 'admin' not in session.get('permisos', ''): return redirect(url_for('index'))
    usuario = Usuario.query.get_or_404(id)
    usuario.cargo_solicitado = None
    db.session.commit()
    flash(f'Solicitud de cambio rechazada para {usuario.nombre_completo}.', 'success')
    return redirect(url_for('admin.admin_usuarios'))

@admin_bp.route('/eliminar_usuario/<int:id>', methods=['POST'])
def eliminar_usuario(id):
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403
    if id != session['usuario_id']:
        usuario = Usuario.query.get_or_404(id)
        if usuario.id == 1:
            return redirect(url_for('admin.admin_usuarios'))
        db.session.delete(usuario)
        db.session.commit()
        flash('Usuario eliminado correctamente.', 'success')
    return redirect(url_for('admin.admin_usuarios'))

@admin_bp.route('/resetear_password/<int:id>', methods=['POST'])
def resetear_password(id):
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403
    
    usuario = Usuario.query.get_or_404(id)
    
    # Protección: no resetear al creador del sistema
    if usuario.id == 1:
        flash("No puedes resetear la contraseña del creador del sistema.", "error")
        return redirect(url_for('admin.admin_usuarios'))
    
    # Establecer contraseña temporal
    clave_temporal = 'kolegium2025'
    usuario.password = generate_password_hash(clave_temporal, method='pbkdf2:sha256')
    db.session.commit()
    
    flash(f'Contraseña de {usuario.nombre_completo} restablecida a la clave temporal.', 'success')
    return redirect(url_for('admin.admin_usuarios'))

@admin_bp.route('/modificar_usuario/<int:id>', methods=['POST'])
def modificar_usuario(id):
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403
    
    usuario = Usuario.query.get_or_404(id)
    nuevo_departamento = request.form.get('departamento_asignado')
    nuevo_activo = 'activo' in request.form

    cambios_realizados = []

    if usuario.id == 1:
        flash("No puedes desactivar al creador del sistema.", "error")
        return redirect(url_for('admin.admin_usuarios'))

    # ---- Validación: TODO se revisa antes de tocar el usuario, para que un
    # error no deje la edición a medias. Campo vacío = no se modifica. ----
    def rechazar(mensaje):
        flash(mensaje, "error")
        return redirect(url_for('admin.admin_usuarios'))

    # Cédula
    cedula_raw = (request.form.get('cedula') or '').strip()
    nueva_cedula = None
    if cedula_raw:
        nueva_cedula = formatear_cedula(cedula_raw)
        if not nueva_cedula:
            return rechazar("La cédula no es válida. Formato esperado: V-12345678 o E-12345678.")
        # Se compara por dígitos: hay cédulas antiguas guardadas sin normalizar ("12345678", "V12345678")
        digitos = re.sub(r'\D', '', nueva_cedula)
        duenio = next((u for u in Usuario.query.filter(Usuario.cedula.isnot(None), Usuario.id != usuario.id)
                       if re.sub(r'\D', '', u.cedula) == digitos), None)
        if duenio:
            return rechazar(f"La cédula {nueva_cedula} ya pertenece a {duenio.nombre_completo} (usuario '{duenio.username}').")

    # Usuario (login): el login lo compara en minúsculas
    nuevo_username_raw = (request.form.get('username') or '').strip()
    nuevo_username = None
    if nuevo_username_raw:
        if ' ' in nuevo_username_raw:
            return rechazar("El nombre de usuario no puede contener espacios.")
        nuevo_username = nuevo_username_raw.lower()
        duenio = Usuario.query.filter(db.func.lower(Usuario.username) == nuevo_username,
                                      Usuario.id != usuario.id).first()
        if duenio:
            return rechazar(f"El usuario '{nuevo_username}' ya está en uso por {duenio.nombre_completo}.")

    # Correo
    nuevo_email_raw = (request.form.get('email') or '').strip()
    nuevo_email = None
    if nuevo_email_raw:
        nuevo_email = nuevo_email_raw.lower()
        if not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', nuevo_email):
            return rechazar("El correo no tiene un formato válido (ej. nombre@dominio.com).")
        duenio = Usuario.query.filter(db.func.lower(Usuario.email) == nuevo_email,
                                      Usuario.id != usuario.id).first()
        if duenio:
            return rechazar(f"El correo {nuevo_email} ya está registrado por {duenio.nombre_completo}.")

    # Contraseña (opcional). El login hace strip() de la clave ingresada,
    # así que se guarda sin espacios en los extremos para que coincida.
    nueva_contrasena = (request.form.get('nueva_contrasena') or '').strip()
    if nueva_contrasena and len(nueva_contrasena) < 6:
        return rechazar("La nueva contraseña debe tener al menos 6 caracteres.")

    # ---- Aplicar cambios ----
    if nueva_cedula and usuario.cedula != nueva_cedula:
        cambios_realizados.append(f"Cédula cambiada de '{usuario.cedula or 'Sin cédula'}' a '{nueva_cedula}'")
        usuario.cedula = nueva_cedula

    if nuevo_username and usuario.username != nuevo_username:
        cambios_realizados.append(f"Usuario cambiado de '{usuario.username}' a '{nuevo_username}'")
        usuario.username = nuevo_username

    if nuevo_email and (usuario.email or '').lower() != nuevo_email:
        cambios_realizados.append(f"Correo cambiado de '{usuario.email}' a '{nuevo_email}'")
        usuario.email = nuevo_email

    if nueva_contrasena:
        usuario.password = generate_password_hash(nueva_contrasena, method='pbkdf2:sha256')
        cambios_realizados.append("Contraseña actualizada")

    if usuario.activo != nuevo_activo:
        usuario.activo = nuevo_activo
        cambios_realizados.append(f"Estatus cambiado a {'Activo' if nuevo_activo else 'Inactivo'}")

    if nuevo_departamento is not None and ('Administrativo' in (usuario.area_trabajo or '') or 'Especialista' in (usuario.area_trabajo or '')):
        # Even if it's an empty string (Ninguno), we save it (or None)
        depto = nuevo_departamento if nuevo_departamento != "" else None
        if usuario.departamento_asignado != depto:
            old_depto = usuario.departamento_asignado or 'Ninguno'
            new_depto = depto or 'Ninguno'
            usuario.departamento_asignado = depto
            cambios_realizados.append(f"Depto cambiado de '{old_depto}' a '{new_depto}'")

    if cambios_realizados:
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            flash("No se pudo guardar: el usuario, el correo o la cédula ya están en uso por otra cuenta.", "error")
            return redirect(url_for('admin.admin_usuarios'))
        flash(" | ".join(cambios_realizados), "success")

    return redirect(url_for('admin.admin_usuarios'))

# ==========================================
# --- IMPORTACIÓN MASIVA DE PERSONAL (Excel) ---
# ==========================================

TAMANO_MAXIMO_EXCEL = 10 * 1024 * 1024  # 10 MB

@admin_bp.route('/importar_personal', methods=['GET', 'POST'])
def importar_personal():
    if not session.get('logeado'):
        return redirect(url_for('auth.login'))
    # Solo el Administrador Supremo: esta herramienta crea cuentas en lote.
    if session.get('nombre_rol') != 'Administrador Supremo':
        return '🚫 No autorizado.', 403

    if request.method == 'GET':
        return render_template('importar_personal.html', resultado=None)

    archivo = request.files.get('archivo')
    if not archivo or not archivo.filename:
        flash('Selecciona un archivo Excel (.xlsx) para importar.', 'error')
        return redirect(url_for('admin.importar_personal'))
    if not archivo.filename.lower().endswith('.xlsx'):
        flash('Formato no soportado. Guarda el archivo como Excel .xlsx e inténtalo de nuevo.', 'error')
        return redirect(url_for('admin.importar_personal'))
    if request.content_length and request.content_length > TAMANO_MAXIMO_EXCEL:
        flash('El archivo supera el límite de 10 MB.', 'error')
        return redirect(url_for('admin.importar_personal'))

    simular = request.form.get('simular') == 'on'
    try:
        resultado = importar_personal_desde_excel(archivo, simular=simular)
    except ValueError as e:
        db.session.rollback()
        flash(str(e), 'error')
        return redirect(url_for('admin.importar_personal'))
    except Exception:
        db.session.rollback()
        current_app.logger.exception('Error importando personal desde Excel')
        flash('No se pudo procesar el archivo. Verifica que sea un Excel válido y vuelve a intentarlo.', 'danger')
        return redirect(url_for('admin.importar_personal'))

    prefijo = 'Simulación (no se guardó nada): ' if simular else 'Importación completada: '
    flash(f"{prefijo}{resultado['actualizados']} usuarios actualizados y "
          f"{resultado['creados']} usuarios nuevos creados "
          f"({resultado['sin_cambios']} sin cambios, {resultado['omitidos']} filas omitidas).",
          'warning' if simular else 'success')
    return render_template('importar_personal.html', resultado=resultado)


# ==========================================
# --- GESTOR MINISTERIAL: TABLERO DE NÓMINA ---
# ==========================================

# Columnas que el tablero audita: si falta alguna, la fila se marca incompleta
CAMPOS_NOMINA = ('cedula', 'cargo', 'codigo_rac', 'fecha_ingreso', 'turno', 'telefono')
ETIQUETAS_NOMINA = {'cedula': 'Cédula', 'cargo': 'Cargo', 'codigo_rac': 'Código RAC',
                    'fecha_ingreso': 'Fecha de Ingreso', 'turno': 'Turno', 'telefono': 'Teléfono'}
SIN_CARGO = '__sin_cargo__'  # valor del filtro para quienes no tienen cargo


def nombre_para_mostrar(usuario):
    """Nombre legal en Title Case: "Nombres Apellidos" si están cargados, si no
    el nombre_completo. Nunca el usuario de login."""
    legal = ' '.join(filter(None, (usuario.nombres, usuario.apellidos)))
    return formatear_nombre(legal or usuario.nombre_completo or '') or 'Sin nombre'


@admin_bp.route('/nomina_ministerial')
def nomina_ministerial():
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403

    usuarios = sorted(Usuario.query.all(),
                      key=lambda u: ((u.apellidos or u.nombre_completo or '').lower(), (u.nombres or '').lower()))
    nombres = {u.id: nombre_para_mostrar(u) for u in usuarios}
    faltantes = {u.id: [ETIQUETAS_NOMINA[c] for c in CAMPOS_NOMINA if not getattr(u, c)] for u in usuarios}
    conteo_cargos = Counter(u.cargo for u in usuarios if u.cargo)
    cargos = sorted(conteo_cargos.items(), key=lambda c: c[0].lower())

    # Ficha completa de cada persona para rellenar el modal (se pasa como JSON)
    fichas = {u.id: {c: valor_para_formulario(u, c) for c in CAMPOS_FICHA} for u in usuarios}

    # Largo máximo de cada columna de texto -> maxlength de los inputs
    largos = {c: getattr(Usuario.__table__.c[c].type, 'length', None) for c in CAMPOS_FICHA}

    resumen = {
        'total': len(usuarios),
        'completos': sum(1 for u in usuarios if not faltantes[u.id]),
        'sin_cargo': sum(1 for u in usuarios if not u.cargo),
        'sin_rac': sum(1 for u in usuarios if not u.codigo_rac),
        'sin_telefono': sum(1 for u in usuarios if not u.telefono),
    }

    return render_template('nomina_ministerial.html',
                           usuarios=usuarios,
                           nombres=nombres,
                           faltantes=faltantes,
                           fichas=fichas,
                           cargos=cargos,
                           resumen=resumen,
                           sin_cargo=SIN_CARGO,
                           pestanas=PESTANAS_FICHA,
                           campos=CAMPOS_FICHA,
                           largos=largos,
                           formatos_js=FORMATOS_JS,
                           sugerencias_extra={'cargo': [c for c, _ in cargos]},
                           hoy_iso=date.today().isoformat())


@admin_bp.route('/nomina_ministerial/actualizar/<int:id>', methods=['POST'])
def actualizar_datos_ministeriales(id):
    """Edición desde el tablero: guarda la ficha ministerial completa (los
    campos de ficha_ministerial.CAMPOS) y nada más. Campo vacío = se borra el
    dato; un campo que no viene en el formulario no se toca."""
    if 'admin' not in session.get('permisos', ''):
        return '🚫 No autorizado.', 403

    usuario = Usuario.query.get_or_404(id)
    nombre = nombre_para_mostrar(usuario)
    f = request.form

    # Volver al tablero con los mismos filtros y la fila editada a la vista
    volver = url_for('admin.nomina_ministerial',
                     **{k: v for k, v in (('q', f.get('volver_q', '').strip()),
                                          ('cargo', f.get('volver_cargo', '')),
                                          ('incompletos', f.get('volver_incompletos', ''))) if v},
                     _anchor=f'fila-{usuario.id}')

    # ---- Validación completa antes de tocar el usuario ----
    nuevos, errores = {}, []
    for campo in CAMPOS_FICHA:
        if campo not in f:
            continue
        try:
            nuevos[campo] = limpiar_valor(campo, f.get(campo),
                                          getattr(Usuario.__table__.c[campo].type, 'length', None))
        except ValueError as e:
            errores.append(str(e))
    if errores:
        flash(f"{nombre}: no se guardó nada. " + ' | '.join(errores), 'error')
        return redirect(volver)

    # ---- Aplicar cambios ----
    def mostrar(campo, valor):
        if valor is None or valor == '':
            return 'vacío'
        tipo = CAMPOS_FICHA[campo].tipo
        if tipo == 'fecha':
            return valor.strftime('%d/%m/%Y')
        if tipo == 'horas':
            return formatear_horas(valor)
        return valor if len(valor) <= 40 else valor[:37] + '…'

    cambios = []
    for campo, valor in nuevos.items():
        actual = getattr(usuario, campo)
        if actual != valor:
            cambios.append(f"{CAMPOS_FICHA[campo].etiqueta}: '{mostrar(campo, actual)}' → '{mostrar(campo, valor)}'")
            setattr(usuario, campo, valor)

    if cambios:
        db.session.commit()
        flash(f"{nombre} — " + ' | '.join(cambios), 'success')
    else:
        flash(f"{nombre}: no hubo cambios.", 'info')
    return redirect(volver)
