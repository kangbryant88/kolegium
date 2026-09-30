// Formateo estricto de datos del personal (tolerancia cero a la mala escritura).
// Mismas reglas que formatear_nombre / _cargo / _codigo / _cedula / _telefono en
// app/web/perfil.py; el backend vuelve a aplicarlas al guardar.
// Uso: <input data-formato="nombre|cargo|codigo|cedula|telefono">

function formatearNombre(texto, final) {
    let t = texto.toLowerCase();
    t = final ? t.split(/\s+/).filter(Boolean).join(' ') : t.replace(/\s{2,}/g, ' ').replace(/^\s+/, '');
    return t.replace(/(^|[\s\-'])(\p{L})/gu, (m, sep, letra) => sep + letra.toUpperCase());
}
function formatearCargo(texto, final) {
    return formatearNombre(texto, final).replace(/\b(Ii{0,2}|Iv|Vi{0,3}|Ix|X)\b/g, r => r.toUpperCase());
}
function formatearCodigo(texto) {
    return texto.replace(/\s+/g, '').toUpperCase();
}
function formatearCedula(texto, final) {
    let t = texto.replace(/[\s.]/g, '').toUpperCase();
    if (!final) return t;
    t = t.replace(/-/g, '');
    if (/^\d+$/.test(t)) t = 'V' + t;
    const m = t.match(/^([VE])(\d{6,9})$/);
    return m ? m[1] + '-' + m[2] : t;
}
function formatearTelefono(texto, final) {
    if (!final) return texto.replace(/[^\d+\-\s()]/g, '');
    let d = texto.replace(/\D/g, '');
    if (d.length === 12 && d.startsWith('58')) d = '0' + d.slice(2);
    else if (d.length === 10 && !d.startsWith('0')) d = '0' + d;
    return /^0[24]\d{9}$/.test(d) ? d.slice(0, 4) + '-' + d.slice(4) : texto.trim();
}
const FORMATOS = { nombre: formatearNombre, cargo: formatearCargo, codigo: formatearCodigo, cedula: formatearCedula, telefono: formatearTelefono };

function aplicarFormato(input, final) {
    const fn = FORMATOS[input.dataset.formato];
    const antes = input.value;
    const despues = fn(antes, final);
    if (antes === despues) return;
    // Conservar la posición del cursor mientras se escribe.
    const pos = input.selectionStart - (antes.length - despues.length);
    input.value = despues;
    if (!final && document.activeElement === input) input.setSelectionRange(pos, pos);
}

document.querySelectorAll('[data-formato]').forEach(input => {
    input.addEventListener('input', () => aplicarFormato(input, false));
    input.addEventListener('blur', () => aplicarFormato(input, true));
    aplicarFormato(input, true);
    // Formato definitivo justo antes de enviar, aunque no haya habido blur.
    if (input.form) input.form.addEventListener('submit', () => aplicarFormato(input, true));
});
