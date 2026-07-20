# 🚀 Portal del Trabajador y Bóveda Documental

Siguiendo tu validación de seguridad, he implementado exitosamente el **Portal del Trabajador** enfocado en el Asistente Manual y Bóveda Documental. El nuevo enfoque es sumamente seguro, intuitivo y guía al docente paso a paso.

## 1. Ajustes en Backend y Seguridad (app.py)
- **Restricción de Acceso:** La ruta `/portal_trabajador` requiere que la persona esté validada e iniciada con su sesión regular en el sistema (`session['logeado']`). Esto garantiza que nadie más pueda acceder a sus claves de autogestión.
- **Bóveda Documental:** Configuré un directorio seguro oculto (`uploads/personal`) en el servidor donde se guardarán físicamente los PDFs o imágenes.
- **Estructura de Base de Datos:** Agregué con éxito las columnas `voucher_path` y `constancia_path` al modelo del trabajador sin afectar sus datos actuales, permitiendo que la base de datos sepa exactamente dónde encontrar el archivo subido de cada persona.
- **Nomenclatura Segura:** Al subir un archivo, el sistema lo renombra automáticamente (Ej. `3_voucher_recibo.pdf`) anteponiendo el ID del usuario para evitar colisiones si dos personas suben un archivo llamado igual.

## 2. Nueva Interfaz: Mi Portal (portal_trabajador.html)

He diseñado una vista dividida en dos grandes tarjetas (Paso 1 y Paso 2) para guiar la experiencia del usuario de forma amigable.

### Paso 1: Asistente MPPE (Izquierda)
- Se muestran sus credenciales actuales extraídas de su perfil.
- **Botón Inteligente de Copiado:** Al presionar "Copiar" al lado de la contraseña, un pequeño script de JavaScript (como pediste) copia la clave al portapapeles y cambia visualmente a color verde con un "¡Copiado!" por 2 segundos.
- Un botón inmenso y claro permite abrir el "Portal del Ministerio" en una nueva pestaña.

### Paso 2: Resguarda tus Documentos (Derecha)
- Formulario de subida de archivos protegido (sólo permite PDF, JPG o PNG).
- El usuario selecciona qué está subiendo (Voucher o Constancia).
- **Indicadores de Estado:** Abajo del formulario, el trabajador puede ver el estado en tiempo real de su expediente. Si ya subió un voucher, este panel se ilumina en verde con una insignia de "Subido"; de lo contrario, aparece en gris marcado como "Pendiente".

> [!TIP]
> **Prueba Recomendada:**
> Inicia sesión con cualquier usuario. Notarás un nuevo enlace en la barra lateral llamado **Mi Portal (MPPE)**. Accede allí, revisa el flujo del Paso 1 probando el botón "Copiar" y realiza la prueba final subiendo un archivo de prueba en el Paso 2 para ver cómo el sistema lo aprueba y marca como "Subido".
