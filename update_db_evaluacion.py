"""
update_db_evaluacion.py

Crea las tablas nuevas del motor de Evaluación Descriptiva y Boletines
(ProyectoAula, BancoIndicador, EvaluacionEstudiante) si aún no existen.

A diferencia de los scripts update_db_*.py anteriores (que hacían ALTER TABLE
sobre tablas ya existentes), estas son tablas completamente nuevas, así que
basta con db.create_all() -- SQLAlchemy solo crea las tablas que falten y no
toca las que ya existen.
"""

from app import app
from models import db

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        print("Tablas 'proyecto_aula', 'banco_indicador' y 'evaluacion_estudiante' verificadas/creadas.")
