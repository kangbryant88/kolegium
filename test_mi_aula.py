import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import app
from flask import session

def test_mi_aula():
    app.config['TESTING'] = True
    app.config['DEBUG'] = True
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess['logeado'] = True
            sess['nombre_rol'] = 'Administrador Supremo'
            sess['usuario_id'] = 1
            sess['area_trabajo'] = 'Directivo'
        
        try:
            response = client.get('/academico/mi_aula?grado_id=1', follow_redirects=True)
            print(f"Status Code: {response.status_code}")
            if response.status_code == 500:
                print("500 ERROR CAUGHT")
                print(response.data.decode('utf-8'))
        except Exception as e:
            print("EXCEPTION CAUGHT:")
            traceback.print_exc()

if __name__ == '__main__':
    test_mi_aula()
