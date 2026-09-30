import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()
key = os.getenv('GEMINI_API_KEY')
print(f"Probando llave que termina en: {key[-4:] if key else 'NINGUNA'}")

try:
    genai.configure(api_key=key)
    # Usamos el alias que pusiste en el código
    model = genai.GenerativeModel('gemini-flash-latest')
    print("Enviando 'Hola' a Gemini...")
    response = model.generate_content("Hola, responde solo con la palabra OK")
    print("RESPUESTA EXITOSA:", response.text)
except Exception as e:
    print(f"\n!!! ERROR EXACTO DE GOOGLE !!!")
    print(f"Tipo: {type(e).__name__}")
    print(f"Mensaje: {str(e)}")
