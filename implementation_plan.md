# Implementación del Portal del Trabajador (Asistente y Bóveda Documental)

Debido a las restricciones de reCAPTCHA en el sitio web del Ministerio, cambiaremos el enfoque para crear un Asistente de Acceso Manual y una Bóveda Documental donde el propio trabajador descargará su documento en el sitio oficial y lo cargará en nuestro sistema.

## User Review Required

> [!WARNING]
> **Seguridad en el Acceso al Portal:** En las instrucciones mencionas *"si el trabajador ingresa su correo y coincide... muéstrale un panel"*. Para evitar que cualquier persona ingrese el correo de un compañero y vea sus credenciales privadas de MPPE, mi propuesta técnica es requerir que el usuario **inicie sesión** en el sistema normalmente (con su contraseña). Una vez dentro del sistema, podrá acceder a la ruta `/portal_trabajador` para ver su bóveda y credenciales. ¿Estás de acuerdo con este enfoque de seguridad?

## Proposed Changes

### Base de Datos (`app.py`)
- **[MODIFY]** `app.py`: Añadir dos nuevas columnas al modelo `Usuario` para almacenar las rutas de los archivos subidos: `voucher_path` (String) y `constancia_path` (String).
- **[MODIFY]** `app.py`: Configurar la variable `app.config['UPLOAD_FOLDER']` apuntando al directorio seguro `uploads/personal`.

### Rutas del Backend (`app.py`)
- **[NEW]** `@app.route('/portal_trabajador', methods=['GET', 'POST'])`: 
  - **GET:** Mostrará la interfaz con las credenciales, el botón de copiar, y el enlace al MPPE.
  - **POST:** Recibirá el archivo subido mediante un formulario, validará su extensión segura (PDF, JPG, PNG), lo guardará en la carpeta `uploads/personal` usando un nombre seguro que incluya el ID del usuario para evitar colisiones, y actualizará la ruta en el registro de la base de datos del usuario actual.

### Interfaz del Portal
- **[NEW]** `templates/portal_trabajador.html`: Plantilla que contendrá:
  - **Paso 1:** Tarjeta con credenciales de MPPE (`usuario_autogestion` y `clave_autogestion`) y un botón "Copiar Clave" con JavaScript integrado. Botón "Abrir Portal del Ministerio" con `target="_blank"`.
  - **Paso 2:** Sección "Resguarda tus Documentos" con un formulario `<form enctype="multipart/form-data">`. Contendrá un `<select>` para elegir el tipo de documento (Voucher o Constancia), un `<input type="file">` y un botón de submit.

## Verification Plan

### Automated/Manual Verification
1. Ingresar al sistema con una cuenta de Docente/Trabajador.
2. Navegar al nuevo `/portal_trabajador`.
3. Validar el funcionamiento del botón "Copiar Clave" (JavaScript).
4. Subir un archivo de prueba (Voucher) y verificar:
   - Que el archivo se guarda físicamente en la carpeta `uploads/personal`.
   - Que la base de datos se actualiza con la ruta del archivo.
   - Que el portal muestra un indicador visual de que el archivo ya fue subido exitosamente.
