// Tablero de Nómina (Gestor Ministerial): búsqueda en tiempo real, filtro
// por cargo y modal único de edición (ficha completa en 4 pestañas).
// Los filtros viajan en la URL (?q=&cargo=&incompletos=1) para que, al
// guardar desde el modal, el tablero vuelva exactamente como estaba.

(function () {
    const buscador = document.getElementById('buscadorNomina');
    const filtroCargo = document.getElementById('filtroCargo');
    const soloIncompletos = document.getElementById('soloIncompletos');
    const contador = document.getElementById('contadorNomina');
    const vacio = document.getElementById('nominaVacia');
    const filas = Array.from(document.querySelectorAll('#tablaNomina tbody tr'));

    // "José Pérez" -> "jose perez": se busca sin acentos ni mayúsculas.
    function normalizar(texto) {
        return (texto || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
    }

    // Se precalcula una vez: texto normalizado y solo-dígitos (para cédulas
    // escritas como "12.345.678" o "V12345678").
    filas.forEach(fila => {
        const texto = normalizar(fila.dataset.busqueda);
        fila._texto = texto;
        fila._digitos = texto.replace(/\D/g, '');
    });

    function aplicarFiltros() {
        const q = normalizar(buscador.value);
        const palabras = q.split(/\s+/).filter(Boolean);
        const qDigitos = q.replace(/\D/g, '');
        const esCedula = qDigitos.length >= 4 && /^[ve]?[\d.\s-]+$/.test(q);
        const cargo = filtroCargo.value;
        const incompletos = soloIncompletos.checked;

        let visibles = 0;
        filas.forEach(fila => {
            const coincide =
                (!cargo || fila.dataset.cargo === cargo) &&
                (!incompletos || fila.dataset.incompleto === '1') &&
                (esCedula ? fila._digitos.includes(qDigitos)
                          : palabras.every(p => fila._texto.includes(p)));
            fila.hidden = !coincide;
            if (coincide) visibles++;
        });

        contador.textContent = `Mostrando ${visibles} de ${filas.length}`;
        vacio.classList.toggle('d-none', visibles > 0);
        guardarEnUrl();
    }

    function guardarEnUrl() {
        const params = new URLSearchParams();
        if (buscador.value.trim()) params.set('q', buscador.value.trim());
        if (filtroCargo.value) params.set('cargo', filtroCargo.value);
        if (soloIncompletos.checked) params.set('incompletos', '1');
        const qs = params.toString();
        history.replaceState(null, '', location.pathname + (qs ? '?' + qs : '') + location.hash);
    }

    // Restaurar filtros desde la URL (p. ej. al volver de guardar)
    const inicial = new URLSearchParams(location.search);
    buscador.value = inicial.get('q') || '';
    if (inicial.get('cargo') && [...filtroCargo.options].some(o => o.value === inicial.get('cargo'))) {
        filtroCargo.value = inicial.get('cargo');
    }
    soloIncompletos.checked = inicial.get('incompletos') === '1';

    let espera;
    buscador.addEventListener('input', () => { clearTimeout(espera); espera = setTimeout(aplicarFiltros, 120); });
    filtroCargo.addEventListener('change', aplicarFiltros);
    soloIncompletos.addEventListener('change', aplicarFiltros);
    aplicarFiltros();

    // Llevar a la vista la fila recién editada (#fila-ID)
    if (location.hash.startsWith('#fila-')) {
        const fila = document.querySelector(location.hash);
        if (fila && !fila.hidden) fila.scrollIntoView({ block: 'center' });
    }

    // ---- Modal de edición: se rellena con la ficha (JSON) de la fila pulsada ----
    const fichas = JSON.parse(document.getElementById('fichasNomina').textContent);
    const modal = document.getElementById('modalNomina');
    const form = document.getElementById('formNomina');
    const campos = form.querySelectorAll('[data-campo]');
    const primeraPestana = document.querySelector('#fichaTabs .nav-link');

    modal.addEventListener('show.bs.modal', evento => {
        const d = evento.relatedTarget.dataset;
        const ficha = fichas[d.id] || {};
        form.action = d.url;
        document.getElementById('nominaNombre').textContent = d.nombre;
        document.getElementById('nominaCedula').textContent = d.cedula;
        campos.forEach(input => { input.value = ficha[input.name] ?? ''; });

        document.getElementById('volverQ').value = buscador.value.trim();
        document.getElementById('volverCargo').value = filtroCargo.value;
        document.getElementById('volverIncompletos').value = soloIncompletos.checked ? '1' : '';

        // Siempre abrir en "Datos Personales"
        bootstrap.Tab.getOrCreateInstance(primeraPestana).show();
        modal.querySelector('.modal-body').scrollTop = 0;
    });
})();
