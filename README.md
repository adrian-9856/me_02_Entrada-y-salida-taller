# Sistema de Control de Horas - KoboToolbox + Google Sheets

Sistema automatizado para el control y reporte de horas laborales importando datos desde KoboToolbox hacia Google Sheets.

## Características

- Importación automática de datos desde KoboToolbox
- Generación de reportes por día, semana, mes o rango de fechas
- Cálculo automático de horas trabajadas
- Diferenciación de tipos de jornada (Normal, Terapia, Computación, Permiso)
- Formato profesional con firmas y resumen general
- Regla de validación: solo se cuentan entradas AM (antes de 12:00)

## Configuración Inicial

### 1. Crear una nueva hoja de cálculo en Google Sheets

1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Dale un nombre descriptivo (ej: "Control de Horas - Mi Empresa")

### 2. Importar el código

1. En tu hoja de cálculo, ve a **Extensiones > Apps Script**
2. Borra todo el código de ejemplo que aparece
3. Copia y pega el contenido del archivo `Code.gs` de este repositorio
4. Guarda el proyecto (Ctrl+S o Cmd+S)

### 3. Configurar el archivo de manifiesto

1. En el editor de Apps Script, ve a **Configuración del proyecto** (icono de engranaje ⚙️)
2. Marca la casilla **"Mostrar archivo de manifiesto 'appsscript.json' en el editor"**
3. En el panel izquierdo, verás aparecer el archivo `appsscript.json`
4. Haz clic en `appsscript.json` y reemplaza su contenido con:

```json
{
  "timeZone": "America/Guatemala",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp"
  ]
}
```

5. Guarda los cambios (Ctrl+S o Cmd+S)

### 4. Configurar variables de empresa

En el archivo `Code.gs`, modifica las siguientes variables al inicio del archivo:

```javascript
var NOMBRE_EMPRESA = "Tu Empresa";
var NIT_EMPRESA = "NIT: XXXXXXXXX-X";
var DIRECCION = "Tu Dirección";
var HORAS_JORNADA_NORMAL = 7; // Horas de trabajo por día
var URL_KOBO = "TU_URL_DE_EXPORTACION_KOBO";
```

#### Obtener la URL de KoboToolbox:

1. Entra a tu cuenta de KoboToolbox
2. Ve al proyecto que contiene tus datos de entrada/salida
3. Click en **DATA > Downloads**
4. Selecciona **Export type: CSV**
5. Click en **Advanced options**
6. Copia la URL que aparece en "Direct link to CSV file"
7. Pega esa URL en la variable `URL_KOBO`

### 5. Autorizar el script

1. Recarga la hoja de Google Sheets (F5)
2. Verás aparecer un nuevo menú: **⏱️ Control de Horas**
3. La primera vez que uses cualquier opción del menú, Google pedirá autorización
4. Click en **Revisar permisos**
5. Selecciona tu cuenta de Google
6. Click en **Configuración avanzada**
7. Click en **Ir a [nombre del proyecto] (no seguro)**
8. Click en **Permitir**

## Uso del Sistema

### Menú Principal

Una vez configurado, verás el menú **⏱️ Control de Horas** con las siguientes opciones:

#### 🔄 Actualizar desde Kobo
Importa o actualiza los datos desde KoboToolbox manualmente.

#### 📅 Reporte por Día
Genera un reporte de las horas trabajadas en una fecha específica.
- Formato: `dd/mm/yyyy` (ej: 15/12/2024)

#### 📆 Reporte por Semana
Genera un reporte de una semana completa (7 días).
- Formato: `dd/mm/yyyy` (fecha de inicio de la semana)

#### 🗓️ Reporte por Mes
Genera un reporte de todo un mes.
- Formato: `mm/yyyy` (ej: 12/2024)

#### 📊 Reporte por Rango
Genera un reporte entre dos fechas específicas.
- Requiere fecha inicio y fecha fin
- Formato: `dd/mm/yyyy`

#### 📋 Reporte Completo
Genera un reporte con TODOS los registros disponibles.

### Actualización Automática (Opcional)

Para configurar la actualización automática cada hora:

1. En el editor de Apps Script, busca la función `configurarActualizacionAutomatica()`
2. Ejecuta esta función manualmente (click en ▶️ Run)
3. Autoriza si es necesario

Esto creará un trigger que importará datos desde KoboToolbox cada hora automáticamente.

## Estructura de Datos

El sistema espera las siguientes columnas en los datos de KoboToolbox:

- `start` - Fecha/hora de inicio
- `end` - Fecha/hora de fin
- `Participante` - Nombre del empleado
- `Ingreso / Egreso/Ingreso` - Marcador de ingreso
- `Ingreso / Egreso/Egreso` - Marcador de egreso
- `Ingreso / Egreso/Terapia (se paga el 100%)` - Tipo especial
- `Ingreso / Egreso/Computación (se paga el 50%)` - Tipo especial
- `Ingreso / Egreso/Permiso` - Tipo especial

## Reglas de Cálculo

1. **Solo entradas AM:** El sistema SOLO procesa registros cuya hora de entrada sea antes de las 12:00 (AM). Las entradas PM se ignoran automáticamente.

2. **Cálculo de horas:**
   - Con ingreso y egreso: calcula la diferencia real
   - Solo con ingreso: asume la jornada normal configurada (default: 7 horas)

3. **Porcentajes de pago:**
   - Normal/Ingreso/Egreso/Terapia: 100%
   - Computación: 50%
   - Permiso: 0%

## Formato de Reporte

Cada reporte generado incluye:

- Detalle por empleado con todas sus entradas/salidas
- Subtotales por empleado
- Resumen general con totales
- Leyendas y notas importantes
- Sección de firmas
- Fecha y hora de generación

Los reportes se crean en nuevas hojas con el nombre `Reporte_[timestamp]`.

## Solución de Problemas

### Error: "Los permisos especificados no son suficientes"

Asegúrate de haber configurado correctamente el archivo `appsscript.json` con todos los permisos necesarios (ver paso 3).

### No aparece el menú "Control de Horas"

1. Recarga la hoja de cálculo (F5)
2. Verifica que el código esté guardado en el editor de Apps Script
3. Espera unos segundos después de recargar

### Error al importar datos de KoboToolbox

1. Verifica que la URL_KOBO sea correcta
2. Asegúrate de que la exportación de KoboToolbox esté configurada como CSV
3. Verifica que tengas acceso a internet

### Los datos no se importan correctamente

1. Revisa que las columnas en KoboToolbox coincidan con las esperadas
2. Verifica el delimitador CSV (el sistema usa punto y coma `;`)

## Soporte

Para problemas o sugerencias, contacta al administrador del sistema o revisa la documentación de:
- [Google Apps Script](https://developers.google.com/apps-script)
- [KoboToolbox](https://support.kobotoolbox.org/)

---

**Versión:** 1.0
**Última actualización:** Diciembre 2024
