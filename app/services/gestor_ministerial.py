"""
Gestor Ministerial - Fase 3 ("El Músculo").

Recibe una plantilla Excel vacía de la Zona Educativa (con una o varias
pestañas: Obreros, Docentes de Aula...) y una instrucción en lenguaje natural;
el Cerebro (ai_ministerial) decide, pestaña por pestaña, qué personal va y qué
dato va en cada columna, y aquí se rellena cada hoja con openpyxl conservando
estilos, bordes, celdas combinadas y logos.

Solo viajan a Gemini los encabezados y la instrucción: ningún dato personal.
"""
import io
import re
import unicodedata
from copy import copy
from datetime import date

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from app.models import Usuario, Rol
from app.services.ai_ministerial import analizar_formato_ministerio
from app.services.ficha_ministerial import nombre_para_mostrar

FILAS_ENCABEZADO = 25          # el encabezado se busca en las primeras 25 filas de cada hoja
MINIMO_ENCABEZADOS = 2         # una hoja con menos celdas de texto es portada o está vacía
LARGO_MAXIMO_ENCABEZADO = 60   # celdas más largas son títulos o advertencias
# Columna de numeración ("N°", "Nro."): no es un dato del trabajador, se rellena 1, 2, 3...
ENCABEZADOS_NUMERACION = {'n', 'no', 'nro', 'num', 'numero', 'item', 'n o'}
PALABRAS_VACIAS = {'de', 'del', 'la', 'las', 'los', 'el', 'y', 'en', 'para', 'con'}


class ErrorGestor(Exception):
    """Error que se le puede mostrar tal cual al administrador."""


def _normalizar(texto):
    texto = unicodedata.normalize('NFKD', str(texto or '')).encode('ascii', 'ignore').decode()
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', texto.lower()).split())


# ==========================================
# --- LECTURA DE LA PLANTILLA ---
# ==========================================

def detectar_encabezados(hoja):
    """
    Fila de encabezados = la que tenga más celdas con texto entre las primeras
    FILAS_ENCABEZADO. Devuelve (número de fila, {columna: texto}).
    """
    mejor_fila, mejores = None, {}
    for fila in hoja.iter_rows(min_row=1, max_row=FILAS_ENCABEZADO):
        textos = {}
        for celda in fila:
            if isinstance(celda, MergedCell) or not isinstance(celda.value, str):
                continue
            texto = ' '.join(celda.value.split())
            if texto and len(texto) <= LARGO_MAXIMO_ENCABEZADO:
                textos[celda.column] = texto
        if len(textos) > len(mejores):
            mejor_fila, mejores = fila[0].row, textos
    return mejor_fila, mejores


# ==========================================
# --- FILTRO DE TRABAJADORES ---
# ==========================================

def _coincide_cargo(pedido, usuario):
    """
    'Docentes fuera de aula' coincide con cargo 'Docente Fuera De Aula': cada
    palabra pedida (sin plural) debe ser el comienzo de alguna palabra del
    cargo, del tipo de personal o del área de trabajo.
    """
    palabras_pedidas = [p for p in _normalizar(pedido).split() if p not in PALABRAS_VACIAS]
    if not palabras_pedidas:
        return True
    # Manda el dato oficial del RAC; el área de Kolegium (el rol en el sistema,
    # p. ej. un obrero con área "Docente de Aula") solo cuenta si no hay cargo.
    oficiales = (usuario.cargo, usuario.tipo_personal)
    for texto in (oficiales if any(oficiales) else (usuario.area_trabajo,)):
        palabras = _normalizar(texto).split()
        if palabras and all(any(w.startswith(p[:max(4, len(p) - 2)]) for w in palabras)
                            for p in palabras_pedidas):
            return True
    return False


def filtrar_trabajadores(filtros):
    """Usuarios con rol asignado (nunca los 'En Espera'), según los filtros del Cerebro."""
    consulta = (Usuario.query.join(Rol, Usuario.rol_id == Rol.id)
                .filter(Usuario.rol_id.isnot(None), Rol.nombre != 'En Espera'))
    if filtros.get('solo_activos', True):
        consulta = consulta.filter(Usuario.activo.is_(True))
    usuarios = consulta.all()

    cargo = filtros.get('cargo_requerido')
    if cargo:
        usuarios = [u for u in usuarios if _coincide_cargo(cargo, u)]
    return sorted(usuarios, key=lambda u: ((u.apellidos or u.nombre_completo or '').lower(),
                                           (u.nombres or '').lower()))


# ==========================================
# --- INYECCIÓN ---
# ==========================================

def _valor(usuario, atributo):
    if atributo == 'nombre_completo':
        return nombre_para_mostrar(usuario)
    valor = getattr(usuario, atributo)
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)  # horas: 36 y no 36.0
    return valor if valor != '' else None


def verificar_espacio(hoja, fila_encabezado, columnas, columnas_numeracion, cantidad):
    """
    Lanza ErrorGestor si alguna celda donde se va a escribir ya tiene contenido
    (p. ej. el pie con las firmas) o es parte de una celda combinada.
    """
    inicio = fila_encabezado + 1
    for i in range(cantidad):
        for col in (*columnas, *columnas_numeracion):
            celda = hoja.cell(row=inicio + i, column=col)
            # La numeración puede venir ya escrita en la plantilla: se reescribe igual
            ocupada = celda.value not in (None, '') and col not in columnas_numeracion
            if isinstance(celda, MergedCell) or ocupada:
                raise ErrorGestor(
                    f'La pestaña "{hoja.title}" no tiene espacio para {cantidad} trabajador(es): la fila '
                    f'{inicio + i} ya tiene contenido en la columna {celda.column_letter}. Deja al menos '
                    f'{cantidad} fila(s) vacía(s) debajo de sus encabezados.')


def inyectar(hoja, fila_encabezado, columnas, columnas_numeracion, usuarios):
    """
    Escribe un trabajador por fila desde la fila siguiente al encabezado.
    Las filas nuevas copian el estilo (bordes, fuente) de la primera fila de
    datos. Llamar antes a verificar_espacio().
    """
    inicio = fila_encabezado + 1
    todas = {**columnas, **{c: None for c in columnas_numeracion}}
    estilos = {col: copy(hoja.cell(row=inicio, column=col)._style) for col in todas}
    for i, usuario in enumerate(usuarios):
        for col in todas:
            celda = hoja.cell(row=inicio + i, column=col)
            if i and not celda.has_style:
                celda._style = copy(estilos[col])
            valor = i + 1 if col in columnas_numeracion else _valor(usuario, columnas[col])
            # openpyxl pone 'yyyy-mm-dd' al asignar una fecha: se respeta el formato
            # de la plantilla y, si no tenía, se usa el venezolano
            formato_plantilla = celda.number_format
            celda.value = valor
            if isinstance(valor, date):
                celda.number_format = formato_plantilla if formato_plantilla != 'General' else 'DD/MM/YYYY'


# ==========================================
# --- ORQUESTADOR ---
# ==========================================

def leer_pestanas(libro):
    """
    {nombre de pestaña: (fila de encabezado, {columna: texto}, [columnas de N°])}
    para cada hoja visible con encabezados. Las hojas ocultas (listas de
    validación, cálculos) y las que no tienen encabezados no se tocan.
    """
    pestanas = {}
    for hoja in libro.worksheets:
        if hoja.sheet_state != 'visible':
            continue
        fila, encabezados = detectar_encabezados(hoja)
        if len(encabezados) < MINIMO_ENCABEZADOS:
            continue
        numeracion = [c for c, t in encabezados.items() if _normalizar(t) in ENCABEZADOS_NUMERACION]
        pestanas[hoja.title] = (fila, encabezados, numeracion)
    return pestanas


def generar_reporte(archivo, instrucciones):
    """
    Devuelve (BytesIO con el Excel relleno, resumen para mostrar al usuario).
    Lanza ErrorGestor con un mensaje claro si algo impide generar el reporte;
    en ese caso no se entrega ningún archivo a medias.
    """
    try:
        libro = load_workbook(archivo)
    except Exception:
        raise ErrorGestor('No se pudo abrir el archivo. Verifica que sea un Excel .xlsx válido.')

    pestanas = leer_pestanas(libro)
    if not pestanas:
        raise ErrorGestor(f'No se encontraron encabezados en las primeras {FILAS_ENCABEZADO} filas '
                          f'de ninguna pestaña visible del archivo.')

    # Solo viajan a la IA los encabezados (sin la columna N°) de cada pestaña
    para_ia = {nombre: list(dict.fromkeys(t for c, t in encabezados.items() if c not in numeracion))
               for nombre, (fila, encabezados, numeracion) in pestanas.items()}

    analisis = analizar_formato_ministerio(instrucciones, para_ia)
    if 'error_api' in analisis:
        # Diagnóstico: se muestra el fallo técnico real de Gemini, sin traducir
        raise ErrorGestor(f"Error de Gemini [{analisis.get('tipo_error', '?')}]: {analisis['error_api']}")
    if not analisis:
        raise ErrorGestor('El Cerebro (Gemini) no respondió. Puede ser la cuota de la API, la clave '
                          'o la conexión; intenta de nuevo en un minuto.')
    hojas_ia = analisis.get('hojas_a_procesar', {})
    if not hojas_ia:
        raise ErrorGestor('El Cerebro no encontró ninguna pestaña que rellenar con datos del personal. '
                          'Revisa los encabezados de la plantilla o detalla más las instrucciones.')

    # 1) Planificar todas las pestañas y verificar espacio ANTES de escribir nada
    plan = []
    for nombre, datos in hojas_ia.items():
        fila, encabezados, numeracion = pestanas[nombre]
        hoja = libro[nombre]
        mapeo = datos['mapeo_columnas']
        columnas = {c: mapeo[t] for c, t in encabezados.items() if t in mapeo}
        usuarios = filtrar_trabajadores(datos['filtros'])
        verificar_espacio(hoja, fila, columnas, numeracion, len(usuarios))
        plan.append((hoja, fila, encabezados, columnas, numeracion, usuarios, datos['filtros']))

    if not any(usuarios for *_, usuarios, _ in plan):
        raise ErrorGestor('Ningún trabajador coincide con los filtros de ninguna pestaña: ' + '; '.join(
            f'"{hoja.title}": ' + (f.get('cargo_requerido') or 'todos') + (' (solo activos)' if f['solo_activos'] else '')
            for hoja, *_, f in plan) + '.')

    # 2) Escribir
    resumen_hojas = []
    for hoja, fila, encabezados, columnas, numeracion, usuarios, filtros in plan:
        inyectar(hoja, fila, columnas, numeracion, usuarios)
        resumen_hojas.append({
            'hoja': hoja.title,
            'fila_encabezado': fila,
            'explicacion': filtros['explicacion'],
            'cargo_requerido': filtros['cargo_requerido'],
            'solo_activos': filtros['solo_activos'],
            'total': len(usuarios),
            'columnas': {encabezados[c]: atributo for c, atributo in columnas.items()},
            'sin_mapear': [t for c, t in encabezados.items() if c not in columnas and c not in numeracion],
            'numeracion': [encabezados[c] for c in numeracion],
        })

    salida = io.BytesIO()
    libro.save(salida)
    salida.seek(0)

    resumen = {
        'hojas': resumen_hojas,
        'total': sum(h['total'] for h in resumen_hojas),
        'sin_procesar': [h.title for h in libro.worksheets if h.title not in hojas_ia],
    }
    return salida, resumen
