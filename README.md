# 📊 Análisis de Asistencia · Austral Pack S.A.

Herramienta para analizar informes de asistencia mensual y generar un reporte ejecutivo en Excel con gráficos y métricas por área.

---

## Funcionalidades

- **Permisos y vacaciones** — días por tipo (con/sin goce, licencias médicas) por persona y área
- **Atrasos de entrada** — minutos de retraso, promedio por área, persona con mayor atraso
- **Horas asignadas vs. asistidas** — diferencia real de horas trabajadas por persona y área
- **Minutos perdidos** — déficit de horas respecto al turno asignado
- **Índice de criticidad (0–10)** — ranking de áreas según frecuencia y duración de atrasos
- **Dashboard con 5 gráficos** — visualización ejecutiva lista para presentar

El reporte de salida es un archivo `.xlsx` con 6 hojas:

| Hoja | Contenido |
|---|---|
| Dashboard | KPIs + gráficos de barras |
| Resumen Ejecutivo | Ranking de criticidad por área |
| Permisos | Detalle por persona |
| Atrasos | Resumen por área + listado de personas |
| Horas Asig vs Asist | Diferencias por persona y área |
| Minutos Perdidos | Déficit por área |

---

## Archivos de entrada

Se requieren tres archivos Excel. El RUT se usa como clave de cruce entre ellos.

| Archivo | Descripción |
|---|---|
| `informe-asistencia-general.xlsx` | Permisos, vacaciones y licencias médicas por persona |
| `informe-asistencia-semanal.xlsx` | Registros diarios de marcación: entrada/salida turno vs. real |
| `Nomina para turnos.xlsx` | Fuente de verdad de identidad: RUT, nombre completo, cargo y área |

> La nómina es la fuente de verdad para nombres y áreas. Si un RUT no aparece en la nómina, esa persona no tendrá nombre ni área asignada en el reporte.

---

## Instalación (solo la primera vez)

```bash
pip install -r requirements.txt
```

---

## Uso

Doble clic en **`iniciar.bat`** — se abre la app en el navegador automáticamente.

O desde la terminal:

```bash
streamlit run app.py
```

---

## Parámetros configurables

En la sección superior de `analisis_asistencia.py`:

| Parámetro | Valor por defecto | Descripción |
|---|---|---|
| `PESO_FRECUENCIA` | `0.5` | Peso de frecuencia en el índice de criticidad |
| `PESO_DURACION` | `0.5` | Peso de duración en el índice de criticidad |
| `UMBRAL_ATRASO_MAX` | `480` min | Atrasos mayores se consideran marcación anómala y se excluyen |
