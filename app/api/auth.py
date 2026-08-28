from flask import Blueprint, jsonify, request
from werkzeug.security import check_password_hash

from app.models import Usuario

auth_api_bp = Blueprint('auth_api', __name__, url_prefix='/api/v1/auth')


@auth_api_bp.route('/login', methods=['POST'])
def login():
    data = request.get_json(silent=True) or {}
    email = (data.get('email') or '').strip().lower()
    password = data.get('password') or ''

    usuario = Usuario.query.filter_by(email=email).first()

    if not usuario or not check_password_hash(usuario.password, password):
        return jsonify(error='Correo o contraseña incorrectos.'), 401

    return jsonify(
        id=usuario.id,
        nombre_completo=usuario.nombre_completo,
        username=usuario.username,
        token='token_temporal_123',
    ), 200
