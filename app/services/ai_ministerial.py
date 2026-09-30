"""
Gestor Ministerial - Fase 2 ("El Cerebro").

Usa Gemini para leer una petición en lenguaje natural ("mándame los docentes
fuera de aula activos") y los encabezados de un formato Excel vacío del
ministerio, y devuelve qué filtrar y qué atributo de Usuario va en cada columna.

La respuesta de la IA nunca se usa tal cual: validar_respuesta() descarta
cualquier atributo fuera de ATRIBUTOS_PERMITIDOS (p. ej. password) y cualquier
encabezado que no exista en el Excel.
"""
import json
import os

import google.generativeai as genai
from dotenv import load_dotenv

from app.services.ficha_ministerial import CAMPOS as CAMPOS_FICHA

load_dotenv()

genai.configure(api_key=os.getenv('GEMINI_API_KEY'))

# gemini-1.5-flash y 2.5-flash ya no estan disponibles; el alias "latest" sigue al Flash
# vigente. Para fijar una version: GEMINI_MODEL=... en .env
MODELO = os.getenv('GEMINI_MODEL', 'gemini-flash-latest')

PROMPT_SISTEMA = (
    'Eres el sistema experto de RRHH de una escuela. Analiza la petición del usuario '
    '(instrucciones) y la lista de encabezados de un formato Excel vacío. '
    'Debes devolver UNICAMENTE un JSON estricto.'
)

modelo = genai.GenerativeModel(MODELO, system_instruction=PROMPT_SISTEMA)

# Único vocabulario que la IA puede usar en mapeo_columnas: datos de identidad
# + la ficha ministerial completa. Nada de credenciales ni rutas de archivos.
ATRIBUTOS_IDENTIDAD = {
    'cedula': 'Cédula (V-12345678)',
    'nombres': 'Nombres',
    'apellidos': 'Apellidos',
    'nombre_completo': 'Nombres y apellidos juntos',
    'fecha_nacimiento': 'Fecha de Nacimiento',
    'sexo': 'Sexo (Masculino/Femenino)',
    'email': 'Correo electrónico',
}
ATRIBUTOS_PERMITIDOS = {**ATRIBUTOS_IDENTIDAD,
                        **{nombre: campo.etiqueta for nombre, campo in CAMPOS_FICHA.items()}}

ESTRUCTURA_RESPUESTA = """{
  "filtros": {
    "explicacion": "Breve justificación del filtro",
    "cargo_requerido": "Valor deducido del cargo a filtrar (ej. 'Docente fuera de aula', 'Obrero') o null si se piden todos",
    "solo_activos": true o false
  },
  "mapeo_columnas": {
    "NOMBRE ORIGINAL DEL ENCABEZADO EXCEL": "atributo_exacto_del_modelo_usuario",
    "OTRO ENCABEZADO": "otro_atributo"
  }
}"""


def _construir_prompt(instrucciones_usuario, lista_encabezados_excel):
    atributos = '\n'.join(f'- {nombre}: {etiqueta}' for nombre, etiqueta in ATRIBUTOS_PERMITIDOS.items())
    return f"""INSTRUCCIONES DEL USUARIO:
{instrucciones_usuario}

ENCABEZADOS DEL FORMATO EXCEL (cópialos exactamente igual como claves de mapeo_columnas):
{json.dumps(lista_encabezados_excel, ensure_ascii=False)}

ATRIBUTOS PERMITIDOS (los valores de mapeo_columnas deben ser EXACTAMENTE uno de estos nombres):
{atributos}

REGLAS:
- Si un encabezado no corresponde a ningún atributo permitido (p. ej. "N°", "Firma"), NO lo incluyas.
- No inventes atributos que no estén en la lista.
- Si el usuario no menciona un cargo, cargo_requerido es null.
- solo_activos es true salvo que el usuario pida explícitamente incluir inactivos.

Responde con este JSON y nada más:
{ESTRUCTURA_RESPUESTA}"""


def validar_respuesta(respuesta, lista_encabezados_excel):
    """Deja solo lo que Kolegium puede usar con seguridad. Devuelve {} si la
    estructura no es la esperada."""
    if not isinstance(respuesta, dict):
        return {}
    filtros = respuesta.get('filtros')
    mapeo = respuesta.get('mapeo_columnas')
    if not isinstance(filtros, dict) or not isinstance(mapeo, dict):
        return {}

    cargo = filtros.get('cargo_requerido')
    encabezados = set(lista_encabezados_excel)
    return {
        'filtros': {
            'explicacion': str(filtros.get('explicacion') or ''),
            'cargo_requerido': cargo.strip() if isinstance(cargo, str) and cargo.strip() else None,
            'solo_activos': filtros.get('solo_activos') is not False,
        },
        'mapeo_columnas': {encabezado: atributo for encabezado, atributo in mapeo.items()
                           if encabezado in encabezados and atributo in ATRIBUTOS_PERMITIDOS},
    }


def analizar_formato_ministerio(instrucciones_usuario, lista_encabezados_excel):
    """
    Devuelve {'filtros': {...}, 'mapeo_columnas': {encabezado: atributo}}
    o {} si la API falla o responde algo inutilizable.
    """
    try:
        respuesta = modelo.generate_content(
            _construir_prompt(instrucciones_usuario, lista_encabezados_excel),
            generation_config={"response_mime_type": "application/json"},
        )
        return validar_respuesta(json.loads(respuesta.text), lista_encabezados_excel)
    except Exception as e:
        print(f'[ai_ministerial] Error consultando Gemini ({MODELO}): {e}')
        return {}
