# Renovaciones VivoLife — MiBlue Medical

Dashboard interactivo para el seguimiento de renovaciones anuales de suscripciones VivoLife.
Usado por CAS y Clínica.

---

## Cómo publicar en GitHub Pages (paso a paso)

### 1. Crear el repositorio en GitHub
1. Ve a [github.com](https://github.com) e inicia sesión
2. Haz clic en **New** (botón verde)
3. Nombre del repo: `renovaciones-vivolife`
4. Selecciona **Public**
5. Haz clic en **Create repository**

### 2. Subir el archivo
1. En la página del repo recién creado, haz clic en **uploading an existing file**
2. Arrastra el archivo `index.html` a la zona de carga
3. Escribe el mensaje: `Dashboard inicial de renovaciones`
4. Haz clic en **Commit changes**

### 3. Activar GitHub Pages
1. Ve a **Settings** (en el repositorio)
2. En el menú izquierdo selecciona **Pages**
3. En "Branch" selecciona **main** y carpeta **/ (root)**
4. Haz clic en **Save**
5. En ~2 minutos tendrás la URL: `https://tu-usuario.github.io/renovaciones-vivolife/`

### 4. Compartir el link
Comparte esa URL con el equipo de CAS y Clínica. El dashboard se abre directamente en el navegador, sin instalar nada.

---

## Cómo funciona el tablero

### Código de colores por prioridad
| Color | Prioridad | Acción | Responsable |
|-------|-----------|--------|-------------|
| 🔴 Rojo pulsante | VENCIDA | Acción inmediata | **CAS** |
| 🟠 Naranja | URGENTE (≤7 días) | Contactar esta semana | **CAS** |
| 🟡 Amarillo | PRÓXIMA (≤30 días) | Iniciar gestión | **CAS + Clínica** |
| 🔵 Azul | MEDIA (≤60 días) | En seguimiento | Asignado |
| 🟢 Verde | BAJA (>60 días) | Normal | Asignado |
| 🟣 Morado | RENOVADA | Completada | — |

### Funciones principales
- **Agregar** nuevas suscripciones con el botón amarillo
- **Renovar** con el botón verde ✓ (extiende automáticamente 1 año)
- **Editar** cualquier campo con el botón de lápiz
- **Filtrar** por estado, asignado o buscar por nombre
- **Exportar CSV** para reportes en Excel
- **Metas** ajustables para CAS y Clínica (botón Metas)

### Los datos se guardan localmente
Los datos se guardan en el navegador (localStorage). Para que todos vean los mismos datos,
una persona debe ser la administradora y exportar el CSV periódicamente, o contactar a
soporte para habilitar una hoja de Google Sheets como base de datos compartida.

---

## Planes VivoLife incluidos
- VivoLife Básico
- VivoLife Plus
- VivoLife Premium
- VivoLife Familiar
- VivoLife Empresarial
