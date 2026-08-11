# Manual Técnico - Cuaderno Excel Avanzado

**Proyecto:** Cuaderno_Excel_Avanzado  
**Autor Principal:** `@migueljerico`  
**Documentación y Análisis:** QwenCloud / Asistente Técnico  
**Versión del Documento:** 2.0.0  
**Fecha de Actualización:** Agosto de 2026  

---

## 1. Arquitectura General del Sistema

El proyecto **Cuaderno_Excel_Avanzado** está estructurado bajo un patrón de arquitectura por capas enfocado en el procesamiento analítico, automatización de hojas de cálculo, modelos de datos y canalizaciones ETL (Extract, Transform, Load) embebidas. 

El siguiente diagrama en bloque ASCII ilustra el flujo de información y la separación de responsabilidades dentro de la solución:

```
└── Capa de Presentación (Interfaz de Usuario)
    │   ├── Hojas de Control y KPI Dashboards
    │   ├── Formularios de Usuario (UserForms VBA / Controles ActiveX)
    │   └── Vistas de Reportes e Interacción Dinámica
    │
    └── Capa de Lógica de Negocio y Automatización
        │   ├── Módulos VBA (Procedimientos, Funciones UDF, Eventos)
        │   ├── Consultas Power Query (Transformación y Limpieza ETL)
        │   └── Motor de Cálculo (Fórmulas de Matriz Dinámica y Medidas DAX)
        │
        └── Capa de Datos y Conectividad
            ├── Modelo de Datos Interno (Power Pivot / Data Model)
            ├── Hojas de Almacenamiento Estructurado (ListObjects / Tablas)
            └── Conectores a Fuentes Externas (Archivos CSV/TXT, SQL Server, APIs REST)
```

### Descripción de las Capas:

1. **Capa de Presentación:** Muestra la información consolidada al usuario final mediante tableros dinámicos, segmentadores de datos, gráficos avanzados y componentes de interfaz para la ejecución de comandos.
2. **Capa de Lógica:** Procesa las reglas de negocio, valida datos ingresados, ejecuta cálculos analíticos complejos mediante macros de código VBA y motor de expresiones avanzadas.
3. **Capa de Datos:** Gestiona la persistencia de la información en estructuras eficientes de tablas integradas (`ListObjects`), repositorios relacionales internos y conexiones externas.

---

## 2. Descripción de Módulos y Componentes Principales

A continuación se detallan los módulos lógicos y estructurales que componen el ecosistema técnico del proyecto:

| Módulo / Archivo | Responsabilidad Principal | Componentes y Funciones Clave |
| :--- | :--- | :--- |
| `mod_Main_Automation.bas` | Coordinación del flujo principal de ejecución de macros y tareas programadas. | `Sub IniciarProcesamiento()`, `Sub LimpiarBuffers()`, `Sub GenerarReporteFinal()` |
| `mod_Data_ETL.bas` | Automatización de la extracción, formateo e ingesta de datos externos. | `Function ImportarCSV(ruta As String) As Boolean`, `Sub RefrescarPowerQuery()` |
| `mod_Custom_UDF.bas` | Biblioteca de Funciones Definidas por el Usuario (UDF) para cálculos no nativos. | `Function CALCULAR_IRPF_AVANZADO()`, `Function BUSCAR_REGEX()` |
| `cls_Logger.cls` | Módulo de clase para el control de logs de auditoría y manejo de excepciones en tiempo de ejecución. | `Sub LogError(errNum As Long, desc As String)`, `Sub LogInfo(msg As String)` |
| `PowerQuery_Queries/` | Definiciones M de Power Query para la transformación de modelos de datos. | Consultas de unificación, despivotado de columnas y normalización de texto. |
| `Templates / Sheets` | Plantillas base estructuradas para captura y representación gráfica de KPIs. | Hoja `Dashboards`, Hoja `DataStore`, Hoja `Config`. |

---

## 3. Interfaces, Macros y Funciones Documentadas

El proyecto expone procedimientos y funciones (UDF / API de VBA) para la interacción programática e inter-hoja.

### Tabla de Automatizaciones y Funciones Exportadas (UDF / Procedimientos)

| Método / Función | Tipo | Descripción | Parámetros |
| :--- | :--- | :--- | :--- |
| `Ejecutar_ETL_Consolidado` | Macro / Sub | Ejecuta el flujo completo de refresco de datos, limpieza de tablas temporales y recalculo del modelo. | *Ninguno* |
| `CALCULAR_PROYECCION_FINANCIERA` | UDF (Function) | Calcula el retorno proyectado aplicando tasa variable y ajuste de inflación según matriz de datos. | `monto_base` (Double), `tasa_anual` (Double), `periodos` (Integer) |
| `Exportar_Reporte_PDF` | Macro / Sub | Genera y exporta la hoja activa o el tablero dinámico seleccionado en formato PDF con marca de agua. | `ruta_destino` (String), `nombre_archivo` (String) |
| `BUSCAR_RECURSIVO_TABLA` | UDF (Function) | Realiza búsquedas multinivel sobre estructuras `ListObject` ignorando celdas vacías o con error. | `rango_tabla` (Range), `criterio` (Variant), `columna_retorno` (Integer) |

---

## 4. Variables de Entorno y Configuración

Para garantizar la correcta portabilidad y ejecución del proyecto entre distintos entornos (desarrollo, pruebas, producción), se utilizan parámetros de configuración almacenados en una hoja dedicada (`Config`) o leídos del sistema:

| Variable / Parámetro | Valor de Ejemplo | Obligatoria | Descripción |
| :--- | :--- | :--- | :--- |
| `ENV_MODE` | `PROD` | Sí | Define el entorno de ejecución (`DEV`, `QA`, `PROD`) para ajustar el detalle de logs. |
| `PATH_DATA_INPUT` | `C:\Datos\Entrada\` | Sí | Ruta local o de red donde se leen los archivos fuente (.csv, .xlsx). |
| `PATH_EXPORT_PDF` | `C:\Reportes\2026\` | No | Directorio por defecto para la exportación de informes PDF. |
| `DB_CONNECTION_STRING` | `Provider=SQLOLEDB;Data Source=...` | No | Cadena de conexión OLEDB/ODBC para ingesta directa desde base de datos. |
| `LOG_LEVEL` | `DEBUG` / `ERROR` | Sí | Nivel de severidad para el registro de eventos en `cls_Logger`. |
| `AUTO_REFRESH_ON_OPEN` | `TRUE` | Sí | Booleano que determina si el libro actualiza las consultas Power Query al abrirse. |

---

## 5. Guía de Despliegue y Ejecución Paso a Paso

Siga estas instrucciones para desplegar e inicializar el libro técnico en un entorno cliente de forma segura:

### Prerrequisitos de Software
* **Sistema Operativo:** Windows 10/11 (Recomendado para soporte completo de automatizaciones COM y VBA).
* **Plataforma:** Microsoft Excel 365 o Microsoft Excel 2021 (64-bit) con soporte para Power Query y Power Pivot.
* **Seguridad:** Configuración de macros en *Habilitar macros con notificación* o firma digital activa.

### Paso 1: Clonar / Descargar el Repositorio
```bash
git clone https://github.com/migueljerico/Cuaderno_Excel_Avanzado.git
cd Cuaderno_Excel_Avanzado
```

### Paso 2: Unblock de Seguridad de Windows (Si aplica)
Al descargar archivos `.xlsm` o `.xlsb` de la red, Windows puede bloquear su ejecución por directivas de MotW (Mark of the Web):
1. Clic derecho sobre el archivo ejecutable del libro Excel.
2. Seleccionar **Propiedades**.
3. En la pestaña **General**, marcar la casilla **Desbloquear** (Unblock) y presionar **Aplicar**.

### Paso 3: Configuración Inicial de Hojas y Rutas
1. Abrir el archivo `Cuaderno_Excel_Avanzado.xlsm`.
2. Navegar a la hoja `Config`.
3. Validar y actualizar la variable `PATH_DATA_INPUT` apuntando a las carpetas locales de su estación de trabajo.

### Paso 4: Habilitar Entorno de Automatización (VBA)
1. Presionar `ALT + F11` para abrir el Editor de Visual Basic for Applications.
2. Ir a **Herramientas (Tools)** > **Referencias (References)**.
3. Verificar que las siguientes librerías estén marcadas correctamente:
   * *Visual Basic For Applications*
   * *Microsoft Excel 16.0 Object Library*
   * *Microsoft Scripting Runtime* (Para soporte de `FileSystemObject`)
   * *Microsoft ActiveX Data Objects 6.1 Library* (Si se utilizan conexiones OLEDB)

### Paso 5: Ejecución del Flujo Principal
1. En la hoja `Dashboards`, presione el botón de comando **"Iniciar Procesamiento Avanzado"**.
2. O presione `ALT + F8`, seleccione la macro `mod_Main_Automation.IniciarProcesamiento` y haga clic en **Ejecutar**.

---

## 6. Limitaciones Conocidas y Posibles Mejoras Futuras

### Limitaciones Conocidas
1. **Límite de Filas de Excel:** El procesamiento directo en hojas físicas está limitado a 1,048,576 filas (mitigado si se utiliza el Modelo de Datos interno mediante Power Pivot).
2. **Monohilo en VBA:** La ejecución de macros en VBA es de un solo hilo (*single-threaded*), lo que puede congelar temporalmente la interfaz de usuario durante la carga masiva de datos.
3. **Compatibilidad Multiplataforma:** Algunas funciones avanzadas de VBA basadas en Win32 API o llamadas a DLLs no son compatibles con Microsoft Excel para macOS.

### Posibles Mejoras Futuras
* **Integración con Python (Excel-Python Net):** Migrar módulos pesados de transformación de datos (`mod_Data_ETL`) a scripts de Python utilizando `pandas` y `openpyxl` / `xlwings`.
* **Automatización CI/CD de Macros:** Implementar librerías como *vba-developer* o *Rubberduck* para control de versiones modular de código `.bas` y `.cls` directamente en Git.
* **Integración con Power BI / Cloud:** Migrar las capas de datos estáticas a tableros de Power BI Service con actualización programada en la nube.

<p align="center">Creado por <a href="https://github.com/migueljerico">@migueljerico</a> y documentado por Google Gemini (gemini-3.6-flash) desde la App Asistente de IA · 2026</p>