"""Utilidades de seguridad para la API móvil."""
from functools import wraps

from flask import current_app, jsonify, request


def _extraer_token():
    """Lee el token de 'Authorization: Bearer <token>' o de 'X-App-Token'."""
    cabecera = request.headers.get('Authorization', '')
    if cabecera.startswith('Bearer '):
        return cabecera[len('Bearer '):].strip()
    return request.headers.get('X-App-Token', '').strip()


def token_required(f):
    """Exige el token de acceso de la app. Sin él (o inválido) -> 401 JSON."""
    @wraps(f)
    def decorada(*args, **kwargs):
        token = _extraer_token()
        esperado = current_app.config.get('APP_API_TOKEN')
        if not token or not esperado or token != esperado:
            return jsonify(error='Token de acceso inválido o ausente.'), 401
        return f(*args, **kwargs)

    return decorada
