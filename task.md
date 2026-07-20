# Implementación del Portal del Trabajador

- [x] Modificar el modelo `Usuario` en `app.py` añadiendo `voucher_path` y `constancia_path`.
- [x] Configurar la carpeta de subida `UPLOAD_FOLDER` segura (`uploads/personal`) en `app.py`.
- [x] Ejecutar la alteración SQLite para añadir las nuevas columnas sin borrar la tabla.
- [x] Crear la ruta `/portal_trabajador` con soporte `GET` y `POST` para recibir archivos.
- [x] Añadir un botón en el menú principal (`base.html`) para que el trabajador pueda acceder al "Portal del Trabajador".
- [x] Crear el frontend `portal_trabajador.html`:
  - [x] Paso 1: Mostrar Credenciales con botón de Copiar al portapapeles.
  - [x] Paso 1: Añadir enlace target="_blank" hacia `https://autogestion.mppe.gob.ve/iniciar_sesion`.
  - [x] Paso 2: Crear formulario simple de subida de archivos (Voucher / Constancia).
