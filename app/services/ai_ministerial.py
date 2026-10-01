"""
Gestor Ministerial - Fase 2 ("El Cerebro").

Usa Gemini para leer una petición en lenguaje natural ("mándame los docentes
fuera de aula activos") y los encabezados de cada pestaña de un formato Excel
vacío del ministerio, y devuelve, pestaña por pestaña, qué personal filtrar y
qué atributo de Usuario va en cada columna.

La respuesta de la IA nunca se usa tal cual: validar_respuesta() descarta
cualquier atributo fuera de ATRIBUTOS_PERMITIDOS (p. ej. password), cualquier
pestaña que no exista en el libro y cualquier encabezado que no sea de esa pestaña.
"""
import json
import os
import time

import google.generativeai as genai
from dotenv import load_dotenv

from app.services.ficha_ministerial import CAMPOS as CAMPOS_FICHA

load_dotenv()

# REST y no gRPC: PythonAnywhere solo deja salir a internet por su proxy HTTP,
# que no deja pasar gRPC.
genai.configure(api_key=os.getenv('GEMINI_API_KEY'), transport='rest')

# La cuota diaria es por modelo: usamos Pro para no depender del limite de Flash.
# Para cambiarlo sin tocar codigo: GEMINI_MODEL=... en .env
MODELO = os.getenv('GEMINI_MODEL', 'gemini-1.5-pro')

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
  "hojas_a_procesar": {
    "NombreDePestaña1": {
      "filtros": {
        "explicacion": "Breve justificación del filtro de esta pestaña",
        "cargo_requerido": "Cargo a filtrar (ej. 'Docente fuera de aula', 'Obrero') o null si van todos",
        "solo_activos": true
      },
      "mapeo_columnas": {"ENCABEZADO EXACTO DE ESTA PESTAÑA": "atributo_exacto_del_modelo_usuario"}
    },
    "NombreDePestaña2": {
      "filtros": {"explicacion": "...", "cargo_requerido": "otro valor", "solo_activos": true},
      "mapeo_columnas": {}
    }
  }
}"""


def _construir_prompt(instrucciones_usuario, encabezados_por_hoja):
    atributos = '\n'.join(f'- {nombre}: {etiqueta}' for nombre, etiqueta in ATRIBUTOS_PERMITIDOS.items())
    return f"""INSTRUCCIONES DEL USUARIO:
{instrucciones_usuario}

PESTAÑAS DEL LIBRO EXCEL Y SUS ENCABEZADOS (nombre de pestaña -> encabezados; copia ambos exactamente igual):
{json.dumps(encabezados_por_hoja, ensure_ascii=False, indent=1)}

ATRIBUTOS PERMITIDOS (los valores de mapeo_columnas deben ser EXACTAMENTE uno de estos nombres):
{atributos}

REGLAS:
- Cada pestaña se llena con su propio personal: deduce el cargo de cada una por su nombre,
  sus encabezados y las instrucciones (ej. pestaña "OBREROS" -> cargo_requerido "Obrero").
- Omite de hojas_a_procesar las pestañas que sean portada, instrucciones o resumen, o que no necesiten datos del personal.
- En mapeo_columnas usa solo encabezados de ESA pestaña. Si un encabezado no corresponde a ningún
  atributo permitido (p. ej. "N°", "Firma"), NO lo incluyas. No inventes atributos.
- Si para una pestaña no se deduce ningún cargo, cargo_requerido es null.
- solo_activos es true salvo que el usuario pida explícitamente incluir inactivos.

Responde con este JSON y nada más:
{ESTRUCTURA_RESPUESTA}"""


def _validar_hoja(datos, encabezados):
    """Una pestaña de la respuesta -> {'filtros', 'mapeo_columnas'} seguro, o None."""
    if not isinstance(datos, dict):
        return None
    filtros = datos.get('filtros') if isinstance(datos.get('filtros'), dict) else {}
    mapeo = datos.get('mapeo_columnas')
    if not isinstance(mapeo, dict):
        return None
    mapeo = {encabezado: atributo for encabezado, atributo in mapeo.items()
             if encabezado in encabezados and atributo in ATRIBUTOS_PERMITIDOS}
    if not mapeo:
        return None  # nada que escribir en esta pestaña
    cargo = filtros.get('cargo_requerido')
    return {
        'filtros': {
            'explicacion': str(filtros.get('explicacion') or ''),
            'cargo_requerido': cargo.strip() if isinstance(cargo, str) and cargo.strip() else None,
            'solo_activos': filtros.get('solo_activos') is not False,
        },
        'mapeo_columnas': mapeo,
    }


def validar_respuesta(respuesta, encabezados_por_hoja):
    """
    Deja solo lo que Kolegium puede usar con seguridad: pestañas que existen
    en el libro, encabezados de ESA pestaña y atributos de la lista blanca.
    Devuelve {} si la estructura no es la esperada.
    """
    if not isinstance(respuesta, dict) or not isinstance(respuesta.get('hojas_a_procesar'), dict):
        return {}
    hojas = {}
    for nombre, datos in respuesta['hojas_a_procesar'].items():
        if nombre not in encabezados_por_hoja:
            continue
        hoja = _validar_hoja(datos, set(encabezados_por_hoja[nombre]))
        if hoja:
            hojas[nombre] = hoja
    return {'hojas_a_procesar': hojas}


INTENTOS_MAXIMOS = 3
ESPERA_CUOTA_SEGUNDOS = 22  # Google castiga ~19 s al pasarse del límite por minuto


def _generar_con_reintentos(prompt):
    """generate_content reintentando ante el 429 (cuota por minuto); otros errores suben de inmediato."""
    for intento in range(1, INTENTOS_MAXIMOS + 1):
        try:
            return modelo.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"},
            )
        except Exception as e:
            es_cuota = '429' in str(e) or 'TooManyRequests' in type(e).__name__ + str(e)
            if not es_cuota or intento == INTENTOS_MAXIMOS:
                raise
            print(f'[ai_ministerial] 429 de Gemini ({MODELO}), intento {intento}/{INTENTOS_MAXIMOS}; '
                  f'reintentando en {ESPERA_CUOTA_SEGUNDOS} s')
            time.sleep(ESPERA_CUOTA_SEGUNDOS)


def analizar_formato_ministerio(instrucciones_usuario, encabezados_por_hoja):
    """
    encabezados_por_hoja: {'NombreDePestaña': ['encabezado1', ...], ...}
    Devuelve {'hojas_a_procesar': {pestaña: {'filtros': {...}, 'mapeo_columnas': {...}}}}
    o, si algo falla, {'error_api': mensaje exacto, 'tipo_error': clase} (diagnóstico).
    """
    try:
        respuesta = _generar_con_reintentos(_construir_prompt(instrucciones_usuario, encabezados_por_hoja))
        texto = respuesta.text
        validado = validar_respuesta(json.loads(texto), encabezados_por_hoja)
        if not validado:
            print(f'[ai_ministerial] Estructura inesperada de Gemini ({MODELO}): {texto[:500]}')
            return {'error_api': f'Gemini ({MODELO}) respondió con una estructura inesperada: {texto[:500]}',
                    'tipo_error': 'EstructuraInvalida'}
        return validado
    except Exception as e:
        print(f'[ai_ministerial] Error consultando Gemini ({MODELO}): {type(e).__name__}: {e}')
        return {'error_api': str(e), 'tipo_error': type(e).__name__}
