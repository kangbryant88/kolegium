"""
update_db_proyecto_aprendizaje.py

Crea las tablas nuevas del Proyecto de Aprendizaje (PA) formal
(ProyectoAprendizaje, ProyectoArea, ProyectoEvaluacion) si aún no existen.

Al igual que update_db_evaluacion.py, estas son tablas completamente nuevas,
así que basta con db.create_all() -- SQLAlchemy solo crea las tablas que
falten y no toca las que ya existen (ProyectoAula, BancoIndicador, etc.).
"""

import os
import sys

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db
from models import ProyectoAprendizaje, ProyectoArea, ProyectoEvaluacion

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        print("Tablas 'proyecto_aprendizaje', 'proyecto_area' y 'proyecto_evaluacion' verificadas/creadas.")
