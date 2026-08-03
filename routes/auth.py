from flask import Blueprint, render_template, request, redirect, url_for, session, flash, current_app
from werkzeug.security import generate_password_hash, check_password_hash
from flask_mail import Message
from models import db, Usuario, Rol, TokenRecuperacion
from extensions import mail
from datetime import datetime, timedelta
import secrets

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username_limpio = request.form['username'].strip().lower()
        u = Usuario.query.filter_by(username=username_limpio).first()
        if u and check_password_hash(u.password, request.form['password']):
            session.update({
                'logeado': True, 'usuario_id': u.id, 'username': u.username, 
                'nombre_completo': u.nombre_completo, 'area_trabajo': u.area_trabajo,
                'rol_id': u.rol_id, 'foto_perfil_path': u.foto_perfil_path,
                'departamento_asignado': u.departamento_asignado
            })
            if u.rol_info:
                session['permisos'], session['nombre_rol'] = u.rol_info.permisos, u.rol_info.nombre
                
                # Ajustes dinámicos por departamento
                if u.rol_info.nombre == 'Docente Especialista' and u.departamento_asignado == 'CRA':
                    session['permisos'] = 'dashboard_general,planificador,asistencia,cra'
                elif u.rol_info.nombre == 'Administrativo' and u.departamento_asignado == 'Defensoría':
                    session['permisos'] = 'dashboard_general,planificador,asistencia,defensoria'
                elif u.rol_info.nombre == 'Administrativo' and u.departamento_asignado == 'Dirección':
                    # Dirección admin gets specific limited base permissions (gestion_personal is handled explicitly in templates)
                    session['permisos'] = 'dashboard_general,asistencia,planificador'

                if u.rol_info.nombre in ['Obrero', 'Personal de Cocina']:
                    return redirect(url_for('portal_trabajador'))
                return redirect(url_for('index'))
            session['permisos'], session['nombre_rol'] = '', 'En Espera'
            return redirect(url_for('auth.en_espera'))
        return render_template('login.html', error='Credenciales incorrectas.')
    return render_template('login.html')

@auth_bp.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        u_name_raw = request.form['username']
        if ' ' in u_name_raw:
            return render_template('registro.html', error="El nombre de usuario no puede contener espacios. Intente con un formato continuo como 'mariaperez' o 'mperez'.")
            
        u_name = u_name_raw.strip().lower()
        e_mail = request.form['email']
        
        if Usuario.query.filter((Usuario.username == u_name) | (Usuario.email == e_mail)).first():
            return render_template('registro.html', error="El usuario o el correo ya están registrados.")
        area_trab = request.form['area_trabajo']
        depto_asignado = request.form.get('departamento_asignado', None)
        if area_trab not in ['Administrativo', 'Especialista (Robótica / Deportes)'] or not depto_asignado:
            depto_asignado = None

        nuevo = Usuario(nombre_completo=request.form['nombre_completo'], username=u_name, email=e_mail,
                        area_trabajo=area_trab, departamento_asignado=depto_asignado,
                        password=generate_password_hash(request.form['password'], method='pbkdf2:sha256'))
        db.session.add(nuevo); db.session.commit()
        if Usuario.query.count() == 1:
            nuevo.rol_id = Rol.query.filter_by(nombre='Administrador Supremo').first().id
            db.session.commit()
        return redirect(url_for('auth.login'))
    return render_template('registro.html')

@auth_bp.route('/logout')
def logout():
    session.clear(); return redirect(url_for('index'))

@auth_bp.route('/en_espera')
def en_espera():
    return render_template('espera.html')

# ==========================================
# --- RECUPERACIÓN DE CONTRASEÑA ---
# ==========================================

@auth_bp.route('/recuperar', methods=['GET', 'POST'])
def recuperar():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        usuario = Usuario.query.filter_by(email=email).first()
        
        if usuario:
            try:
                # Invalidar tokens anteriores no usados de este usuario
                tokens_anteriores = TokenRecuperacion.query.filter_by(
                    usuario_id=usuario.id, usado=False
                ).all()
                for t in tokens_anteriores:
                    t.usado = True
                
                # Generar nuevo token seguro
                token = secrets.token_urlsafe(32)
                nuevo_token = TokenRecuperacion(
                    token=token,
                    usuario_id=usuario.id
                )
                db.session.add(nuevo_token)
                db.session.commit()
                
                # Construir enlace de recuperación
                enlace = request.host_url.rstrip('/') + url_for('auth.restablecer', token=token)
                
                # Obtener remitente asegurando que no sea None
                remitente = current_app.config.get('MAIL_USERNAME')
                if not remitente:
                    remitente = 'eepdanieloleary9@gmail.com'
                
                # Enviar correo
                msg = Message(
                    "Kolegium - Recuperación de Contraseña",
                    sender=remitente,
                    recipients=[usuario.email]
                )
                msg.html = f"""
                <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                    <div style="background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%); padding: 30px; text-align: center;">
                        <h1 style="color: #ffffff; margin: 0; font-size: 24px;">🔐 Recuperación de Contraseña</h1>
                    </div>
                    <div style="padding: 40px 30px; color: #333333; background-color: #ffffff;">
                        <p style="font-size: 16px; line-height: 1.6;">Hola <strong>{usuario.nombre_completo}</strong>,</p>
                        <p style="font-size: 16px; line-height: 1.6;">Recibimos una solicitud para restablecer tu contraseña en <strong>Kolegium</strong>.</p>
                        <p style="font-size: 16px; line-height: 1.6;">Haz clic en el siguiente botón para crear una nueva contraseña:</p>
                        <div style="text-align: center; margin: 35px 0;">
                            <a href="{enlace}" style="background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%); color: #ffffff; padding: 14px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; display: inline-block;">Restablecer Contraseña</a>
                        </div>
                        <div style="background-color: #fef3c7; padding: 12px 16px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #f59e0b;">
                            <p style="margin: 0; font-size: 14px; color: #92400e;">⏰ Este enlace expira en <strong>30 minutos</strong>. Si no solicitaste este cambio, ignora este correo.</p>
                        </div>
                        <p style="font-size: 13px; color: #9ca3af; margin-top: 20px;">Si el botón no funciona, copia y pega este enlace en tu navegador:<br><a href="{enlace}" style="color: #6366f1; word-break: break-all;">{enlace}</a></p>
                    </div>
                    <div style="background-color: #f9fafb; padding: 20px; text-align: center; font-size: 12px; color: #6b7280; border-top: 1px solid #e0e0e0;">
                        <p style="margin: 0;">Mensaje automático de Kolegium. No responder a este correo.</p>
                    </div>
                </div>
                """
                mail.send(msg)
                print(f"✔️ CORREO DE RECUPERACIÓN ENVIADO A: {usuario.email}")
                flash('Si el correo está registrado, recibirás un enlace para restablecer tu contraseña.', 'success')
                return render_template('recuperar.html', enviado=True)
            except Exception as e:
                db.session.rollback()
                print(f"❌ ERROR EN RECUPERACIÓN DE CONTRASEÑA: {e}")
                print("💡 DIAGNÓSTICO: Verifica que las variables de entorno MAIL_USER y MAIL_PASS estén configuradas en tu hosting y sean correctas.")
                import traceback
                traceback.print_exc()
                flash('Ocurrió un error en el servidor al intentar enviar el correo. Verifica la configuración en la nube.', 'error')
                return render_template('recuperar.html')
        
    return render_template('recuperar.html')

@auth_bp.route('/restablecer/<token>', methods=['GET', 'POST'])
def restablecer(token):
    # Buscar token válido
    token_obj = TokenRecuperacion.query.filter_by(token=token, usado=False).first()
    
    if not token_obj:
        flash('El enlace de recuperación no es válido o ya fue utilizado.', 'error')
        return redirect(url_for('auth.login'))
    
    # Verificar que no haya expirado (30 minutos)
    tiempo_limite = token_obj.fecha_creacion + timedelta(minutes=30)
    if datetime.now() > tiempo_limite:
        token_obj.usado = True
        db.session.commit()
        flash('El enlace de recuperación ha expirado. Solicita uno nuevo.', 'error')
        return redirect(url_for('auth.recuperar'))
    
    if request.method == 'POST':
        nueva_password = request.form.get('password', '')
        confirmar_password = request.form.get('confirmar_password', '')
        
        if len(nueva_password) < 6:
            return render_template('restablecer.html', token=token, error='La contraseña debe tener al menos 6 caracteres.')
        
        if nueva_password != confirmar_password:
            return render_template('restablecer.html', token=token, error='Las contraseñas no coinciden.')
        
        # Actualizar contraseña
        usuario = Usuario.query.get(token_obj.usuario_id)
        usuario.password = generate_password_hash(nueva_password, method='pbkdf2:sha256')
        token_obj.usado = True
        db.session.commit()
        
        flash('¡Contraseña actualizada exitosamente! Ya puedes iniciar sesión.', 'success')
        return redirect(url_for('auth.login'))
    
    return render_template('restablecer.html', token=token)

@auth_bp.route('/test-email')
def test_email_route():
    # Diagnóstico: mostrar qué valores tiene Flask-Mail configurados
    diag = "<h2>🔍 Diagnóstico de Correo</h2>"
    diag += "<table border='1' cellpadding='8' style='border-collapse:collapse;'>"
    
    mail_server = current_app.config.get('MAIL_SERVER')
    mail_port = current_app.config.get('MAIL_PORT')
    mail_tls = current_app.config.get('MAIL_USE_TLS')
    mail_user = current_app.config.get('MAIL_USERNAME')
    mail_pass = current_app.config.get('MAIL_PASSWORD')
    mail_sender = current_app.config.get('MAIL_DEFAULT_SENDER')
    
    # Mostrar valores (ocultar contraseña parcialmente)
    pass_display = '***' + mail_pass[-4:] if mail_pass and len(mail_pass) > 4 else ('(vacío)' if not mail_pass else '***')
    
    diag += f"<tr><td>MAIL_SERVER</td><td>{mail_server or '(vacío)'}</td></tr>"
    diag += f"<tr><td>MAIL_PORT</td><td>{mail_port}</td></tr>"
    diag += f"<tr><td>MAIL_USE_TLS</td><td>{mail_tls}</td></tr>"
    diag += f"<tr><td>MAIL_USERNAME</td><td>{mail_user or '(vacío)'}</td></tr>"
    diag += f"<tr><td>MAIL_PASSWORD</td><td>{pass_display}</td></tr>"
    diag += f"<tr><td>MAIL_DEFAULT_SENDER</td><td>{mail_sender or '(vacío)'}</td></tr>"
    diag += "</table><br>"
    
    # Verificar si .env fue cargado
    import os
    env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env')
    env_exists = os.path.exists(env_path)
    diag += f"<p>📁 Archivo .env en: {env_path} — {'✅ EXISTE' if env_exists else '❌ NO EXISTE'}</p>"
    
    # Intentar enviar
    try:
        from extensions import mail
        from flask_mail import Message
        msg = Message(
            "Prueba de configuración de correo",
            sender=mail_sender or mail_user or 'eepdanieloleary9@gmail.com',
            recipients=[mail_user or 'eepdanieloleary9@gmail.com']
        )
        msg.body = "Si recibes esto, el correo está funcionando correctamente en la nube."
        mail.send(msg)
        diag += "<h3 style='color:green;'>✔️ ¡Correo enviado exitosamente!</h3>"
        return diag, 200
    except Exception as e:
        import traceback
        error_info = traceback.format_exc()
        diag += f"<h3 style='color:red;'>❌ Error enviando correo:</h3>"
        diag += f"<pre>{error_info}</pre>"
        return diag, 500
