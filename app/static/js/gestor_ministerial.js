// Gestor Ministerial: envía la plantilla con fetch para poder mostrar el
// indicador de carga mientras trabaja la IA, descargar el Excel generado y
// explicar, pestaña por pestaña, qué decidió el Cerebro.
// Respuesta: {archivo: base64, nombre, resumen}. Sin JavaScript el formulario
// se envía normal, se descarga el archivo y los errores llegan como flash.

(function () {
    const form = document.getElementById('formGestor');
    const boton = document.getElementById('btnGenerar');
    const spinner = document.getElementById('spinnerGenerar');
    const icono = document.getElementById('iconoGenerar');
    const texto = document.getElementById('textoGenerar');
    const error = document.getElementById('errorGestor');
    const panelAyuda = document.getElementById('panelAyuda');
    const panelResumen = document.getElementById('panelResumen');

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

    function descargar(base64, nombre) {
        const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const enlace = document.createElement('a');
        enlace.href = url;
        enlace.download = nombre;
        document.body.appendChild(enlace);
        enlace.click();
        enlace.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    // Todo con textContent: el texto viene de la IA y no se interpreta como HTML.
    function el(etiqueta, clase, contenido) {
        const nodo = document.createElement(etiqueta);
        if (clase) nodo.className = clase;
        if (contenido !== undefined) nodo.append(...[].concat(contenido));
        return nodo;
    }
    function chip(encabezado, valor, gris) {
        return el('span', 'gestor-chip' + (gris ? ' gris' : ''),
                  valor === undefined ? encabezado : [encabezado + ' → ', el('code', '', valor)]);
    }

    function tarjetaHoja(h) {
        const filtro = [h.cargo_requerido ? `Cargo: ${h.cargo_requerido}` : 'Todos los cargos',
                        h.solo_activos ? 'solo activos' : 'activos e inactivos'].join(' · ');
        const tarjeta = el('div', 'gestor-hoja', [
            el('h6', '', `${h.hoja} — ${h.total} trabajador(es)`),
            el('div', 'meta', `${filtro}. Encabezados en la fila ${h.fila_encabezado}.`),
        ]);
        if (h.explicacion) tarjeta.append(el('div', 'mb-2', h.explicacion));
        const columnas = el('div', '');
        (h.numeracion || []).forEach(t => columnas.append(chip(t, '1, 2, 3…')));
        Object.entries(h.columnas || {}).forEach(([t, atributo]) => columnas.append(chip(t, atributo)));
        (h.sin_mapear || []).forEach(t => columnas.append(chip(t + ' (vacía)', undefined, true)));
        tarjeta.append(columnas);
        return tarjeta;
    }

    function mostrarResumen(r) {
        const hojas = r.hojas || [];
        document.getElementById('resTotal').textContent = r.total ?? 0;
        document.getElementById('resNumHojas').textContent = hojas.length;
        document.getElementById('resHojas').replaceChildren(...hojas.map(tarjetaHoja));

        const sinProcesar = document.getElementById('resSinProcesar');
        sinProcesar.textContent = (r.sin_procesar || []).length
            ? `Pestañas sin tocar (portadas o sin datos del personal): ${r.sin_procesar.join(', ')}.` : '';
        sinProcesar.classList.toggle('d-none', !(r.sin_procesar || []).length);

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
            const datos = await respuesta.json().catch(() => ({}));
            if (!respuesta.ok || !datos.archivo) {
                throw new Error(datos.error || `Error inesperado del servidor (${respuesta.status}).`);
            }
            descargar(datos.archivo, datos.nombre);
            mostrarResumen(datos.resumen || {});
        } catch (e) {
            mostrarError(e.message === 'Failed to fetch' ? 'No se pudo conectar con el servidor.' : e.message);
        } finally {
            cargando(false);
        }
    });
})();
