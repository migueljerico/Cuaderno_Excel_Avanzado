# 📊 Cuaderno de Excel Avanzado

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-1255CC?style=for-the-badge&logo=visual-basic&logoColor=white)
![Estado](https://img.shields.io/badge/Estado-Publicado_2026-brightgreen?style=for-the-badge)
![Licencia](https://img.shields.io/badge/Licencia-MIT-orange?style=for-the-badge)

*Recopilación técnica, modelos analíticos, fórmulas complejas y automatizaciones VBA para el dominio profesional de Microsoft Excel.*

---

## 🔗 Acceso / Demo

Puedes acceder al repositorio directamente en GitHub para explorar los libros de trabajo, scripts y plantillas:

- **Repositorio oficial:** [github.com/migueljerico/Cuaderno_Excel_Avanzado](https://github.com/migueljerico/Cuaderno_Excel_Avanzado)
- **Descargar última versión (2026):** [ZIP del Proyecto](https://github.com/migueljerico/Cuaderno_Excel_Avanzado/archive/refs/heads/main.zip)

---

## 📋 Descripción

**Cuaderno_Excel_Avanzado** es un recurso estructurado orientando a profesionales, analistas de datos y desarrolladores de hojas de cálculo que buscan llevar sus habilidades de Microsoft Excel al siguiente nivel. Este proyecto recopila casos de uso reales, estructuras de datos optimizadas y guías paso a paso para dominar las capacidades analíticas de la herramienta.

El proyecto resuelve la problemática habitual del manejo ineficiente de grandes volúmenes de datos en hojas de cálculo, proporcionando patrones de diseño de modelos financieros, automatización de tareas repetitivas mediante macros VBA y técnicas avanzadas de transformación de datos mediante Power Query y Power Pivot.

Diseñado con un enfoque práctico y corporativo, el cuaderno abarca desde la formulación matricial dinámica hasta la creación de cuadros de mando (dashboards) ejecutivos e interactivos, adecuados para el entorno empresarial actual en 2026.

---

## ✨ Funcionalidades

| Funcionalidad | Descripción |
| :--- | :--- |
| **Formulación Avanzada y Dinámica** | Implementación de `BUSCARX` (`XLOOKUP`), `INDICE/COINCIDIR`, y matrices dinámicas (`FILTRAR`, `UNICOS`, `ORDENAR`, `SELECCIONARCOLS`). |
| **Automatización con VBA / Macros** | Módulos de código para automatizar generación de reportes, envío masivo de correos y validación de datos. |
| **Modelado de Datos y DAX** | Creación de esquemas en estrella mediante Power Pivot e implementación de medidas calculadas con lenguaje DAX. |
| **Transformación ETL (Power Query)** | Limpieza, unificación y estructuración automatizada de fuentes de datos heterogéneas (CSV, SQL, Web, JSON). |
| **Dashboards e Indicadores (KPIs)** | Diseños de tableros ejecutivos con segmentadores de datos, gráficos dinámicos y controles de formulario. |
| **Auditoría y Optimización** | Buenas prácticas para reducir el tamaño de archivos, optimizar tiempos de cálculo y rastrear dependencias complejas. |

---

## ⚙️ Instalación

Para utilizar los cuadernos de trabajo y ejecutar los scripts de automatización en tu equipo local, sigue estos pasos:

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/migueljerico/Cuaderno_Excel_Avanzado.git
   cd Cuaderno_Excel_Avanzado
   ```

2. **Requisitos previos del sistema:**
   - Microsoft Excel 365, Office 2021 o superior (edición 2026 recomendada).
   - Habilitar la pestaña **Programador / Desarrollador** en Excel:
     - Ir a `Archivo` > `Opciones` > `Personalizar cinta de opciones` > Marcar **Programador**.

3. **Configuración de Macros y Seguridad:**
   - En Excel, dirígete a `Archivo` > `Opciones` > `Centro de confianza` > `Configuración del Centro de confianza`.
   - En **Configuración de macros**, selecciona *Habilitar macros VBA* (o habilitar con notificación) para permitir la ejecución de las automatizaciones incluidas.

---

## 🚀 Uso

### Ejemplo 1: Automatización VBA para Limpieza de Datos

Puedes importar el módulo `LimpiezaDatos.bas` en el editor de VBA (`ALT + F11`) para ejecutar la siguiente rutina optimizada:

```vba
Public Sub LimpiarYEstructurarDatos()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Eliminar filas en blanco y formatear encabezados
    With ws.UsedRange
        .RemoveDuplicates Columns:=Array(1), Header:=xlYes
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(33, 115, 70) ' Verde Excel
        .Rows(1).Font.Color = RGB(255, 255, 255)
        .Columns.AutoFit
    End With
    
    MsgBox "Proceso de limpieza completado con éxito.", vbInformation, "Excel Avanzado 2026"

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error en Proceso"
    Resume CleanExit
End Sub
```

### Ejemplo 2: Consulta de Matriz Dinámica para Reportes Filtro

Uso de fórmulas dinámicas en Excel para la extracción en tiempo real de registros sin macros:

```excel
=ORDENAR(FILTRAR(TablaVentas[Cliente]:TablaVentas[Monto]; (TablaVentas[Region]="LATAM") * (TablaVentas[Monto]>5000); "Sin Resultados"); 2; -1)
```

---

## 📁 Estructura del proyecto

```text
Cuaderno_Excel_Avanzado/
├── README.md
├── docs/
│   ├── Guia_Formulas_Matriciales.pdf
│   └── Manual_PowerQuery_ETL.pdf
├── macros/
│   ├── LimpiezaDatos.bas
│   ├── GeneradorReportes.bas
│   └── ExportadorPDF.bas
├── modelos/
│   ├── Dashboard_Ejecutivo.xlsx
│   ├── Modelo_Financiero_Proyectado.xlsx
│   └── Analisis_Ventas_PowerPivot.xlsx
└── plantillas/
    ├── Plantilla_ETL_PowerQuery.xlsx
    └── Plantilla_Formulario_VBA.xlsm
```

---

## 🛠️ Tecnologías

| Herramienta | Versión / Detalle | Uso en el proyecto |
| :--- | :--- | :--- |
| **Microsoft Excel** | versión 365 / 2026 | Entorno base para modelos de datos, dashboards y hojas de cálculo |
| **VBA (Visual Basic for Applications)** | v7.1 | Automatización de flujos de trabajo, eventos e interfaces de usuario |
| **Power Query (Engine M)** | Integrado en Excel | Conexión, transformación y carga (ETL) de orígenes de datos |
| **Power Pivot (DAX)** | Integrado en Excel | Modelado de datos relacionales y métricas analíticas de rendimiento |

---

## 📚 Contexto formativo o motivación del proyecto

Este repositorio fue creado por **[@migueljerico](https://github.com/migueljerico)** con el propósito de servir como bitácora de aprendizaje y cuaderno de referencia técnica para el manejo profesional de Microsoft Excel. 

El contenido recopila soluciones prácticas a problemas complejos de inteligencia de negocios a nivel departamental, sirviendo tanto de guía formativa como de biblioteca de código reusable (módulos `.bas` y modelos `.xlsx`) para proyectos de análisis de datos e ingeniería financiera.

<p align="center">Creado por <a href="https://github.com/migueljerico">@migueljerico</a> y documentado por Google Gemini (gemini-3.6-flash) desde la App Asistente de IA · 2026</p>