from flask import Blueprint, jsonify

api_bp = Blueprint('api', __name__, url_prefix='/api/v1')


@api_bp.route('/ping')
def ping():
    return jsonify(status='ok')


# Registra las rutas de los submódulos de la API en api_bp.
# Import al final para evitar el import circular con `api_bp`.
from app.api import mi_aula as _mi_aula  # noqa: E402,F401
