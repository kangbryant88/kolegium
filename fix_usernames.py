import os
import sys

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app
from models import db, Usuario

def arreglar_nombres_de_usuario():
    with app.app_context():
        usuarios = Usuario.query.all()
        actualizados = 0
        for u in usuarios:
            if u.username:
                # Quitamos espacios al inicio/final y forzamos minúsculas
                nuevo_username = u.username.strip().lower()
                # También quitamos los espacios internos (opcional pero recomendado)
                # nuevo_username = nuevo_username.replace(" ", "")
                
                if u.username != nuevo_username:
                    print(f"[MODIFICANDO] '{u.username}' -> '{nuevo_username}'")
                    u.username = nuevo_username
                    actualizados += 1
        
        if actualizados > 0:
            db.session.commit()
            print(f"[OK] Se actualizaron {actualizados} usuarios en la base de datos.")
        else:
            print("[INFO] Todos los nombres de usuario ya estaban normalizados. No se hicieron cambios.")

if __name__ == '__main__':
    arreglar_nombres_de_usuario()
