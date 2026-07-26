from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from werkzeug.security import generate_password_hash, check_password_hash
from models import db, Usuario, Rol

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
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
            return redirect(url_for('auth.en_espera'))
        return render_template('login.html', error='Credenciales incorrectas.')
    return render_template('login.html')

@auth_bp.route('/registro', methods=['GET', 'POST'])
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
        return redirect(url_for('auth.login'))
    return render_template('registro.html')

@auth_bp.route('/logout')
def logout():
    session.clear(); return redirect(url_for('index'))

@auth_bp.route('/en_espera')
def en_espera():
    return render_template('espera.html')
