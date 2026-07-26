from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from flask_mail import Message
from models import db, Usuario, Rol

# Importamos 'mail' desde extensions (patrón Factory)
from extensions import mail

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
        'Equipo Directivo / Administrativo': 'Equipo Directivo',
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
    return redirect(url_for('admin.admin_usuarios'))
