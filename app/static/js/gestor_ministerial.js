// Gestor Ministerial: envía la plantilla con fetch para poder mostrar el
// indicador de carga mientras trabaja la IA, descargar el Excel generado y
// explicar qué decidió el Cerebro (cabecera X-Gestor-Resumen).
// Sin JavaScript el formulario se envía normal y los errores llegan como flash.

(function () {
    const form = document.getElementById('formGestor');
    const boton = document.getElementById('btnGenerar');
    const spinner = document.getElementById('spinnerGenerar');
    const icono = document.getElementById('iconoGenerar');
    const texto = document.getElementById('textoGenerar');
    const error = document.getElementById('errorGestor');
    const panelAyuda = document.getElementById('panelAyuda');
    const panelResumen = document.getElementById('panelResumen');
    const NOMBRE_ARCHIVO = 'Reporte_Ministerial_Generado.xlsx';

    function cargando(activo) {
        boton.disabled = activo;
        spinner.classList.toggle('d-none', !activo);
        icono.classList.toggle('d-none', activo);
        texto.textContent = activo ? 'Analizando y generando…' : 'Generar reporte';
    }

    function mostrarError(mensaje) {
        error.textContent = mensaje;
        error.classList.remove('d-none');
        error.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function descargar(blob) {
        const url = URL.createObjectURL(blob);
        const enlace = document.createElement('a');
        enlace.href = url;
        enlace.download = NOMBRE_ARCHIVO;
        document.body.appendChild(enlace);
        enlace.click();
        enlace.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function chip(contenido, gris) {
        const span = document.createElement('span');
        span.className = 'gestor-chip' + (gris ? ' gris' : '');
        span.append(...contenido);
        return span;
    }

    // Todo con textContent: el texto viene de la IA y no se interpreta como HTML.
    function mostrarResumen(r) {
        document.getElementById('resTotal').textContent = r.total ?? '—';
        const partes = [r.cargo_requerido ? `Cargo: ${r.cargo_requerido}` : 'Todos los cargos',
                        r.solo_activos ? 'solo activos' : 'activos e inactivos'];
        document.getElementById('resFiltro').textContent = partes.join(' · ');
        document.getElementById('resExplicacion').textContent = r.explicacion || '—';

        const columnas = document.getElementById('resColumnas');
        columnas.replaceChildren();
        (r.numeracion || []).forEach(t => columnas.append(chip([t + ' → ', Object.assign(document.createElement('code'), { textContent: '1, 2, 3…' })])));
        Object.entries(r.columnas || {}).forEach(([encabezado, atributo]) => {
            columnas.append(chip([encabezado + ' → ', Object.assign(document.createElement('code'), { textContent: atributo })]));
        });

        const sinMapear = r.sin_mapear || [];
        const contenedor = document.getElementById('resSinMapear');
        contenedor.replaceChildren(...sinMapear.map(t => chip([t], true)));
        contenedor.classList.toggle('d-none', !sinMapear.length);
        document.getElementById('resSinMapearTitulo').classList.toggle('d-none', !sinMapear.length);

        panelAyuda.classList.add('d-none');
        panelResumen.classList.remove('d-none');
    }

    form.addEventListener('submit', async evento => {
        evento.preventDefault();
        if (!form.reportValidity()) return;
        error.classList.add('d-none');
        cargando(true);
        try {
            const respuesta = await fetch(form.action, {
                method: 'POST',
                body: new FormData(form),
                headers: { 'X-Requested-With': 'fetch' },
            });
            if (!respuesta.ok) {
                const datos = await respuesta.json().catch(() => ({}));
                throw new Error(datos.error || `Error inesperado del servidor (${respuesta.status}).`);
            }
            const blob = await respuesta.blob();
            descargar(blob);
            let resumen = {};
            try { resumen = JSON.parse(decodeURIComponent(respuesta.headers.get('X-Gestor-Resumen') || '%7B%7D')); } catch (e) { }
            mostrarResumen(resumen);
        } catch (e) {
            mostrarError(e.message === 'Failed to fetch' ? 'No se pudo conectar con el servidor.' : e.message);
        } finally {
            cargando(false);
        }
    });
})();
