# assemble

![GitHub release](https://img.shields.io/github/v/release/linhermx/assemble)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)

Herramienta interna para calcular la **capacidad de producción de hornos** a partir de un **checklist estándar** y un **inventario estándar**, desarrollada en Python.

El sistema compara los materiales requeridos para un horno contra la existencia actual y responde de forma clara:

- cuántos hornos completos pueden salir hoy
- qué material bloquea el siguiente horno
- qué faltantes existen para completar el siguiente horno
- qué faltantes existen para alcanzar un objetivo de `N` hornos

Incluye:

- aplicación **Windows (.exe)** para usuarios no técnicos
- interfaz gráfica (GUI)
- actualización automática mediante launcher
- uso por línea de comandos (CLI) para usuarios técnicos
- generación de reporte en Excel y log de ejecución

---

## Características

- cálculo de capacidad de producción con inventario actual
- identificación del **material más limitante**
- simulación de faltantes para el **siguiente horno**
- simulación de faltantes para un **objetivo de N hornos**
- exclusión de materiales no limitantes con `incluir_en_capacidad = NO`
- soporte para consumos fraccionales en unidades como `Metro`
- agregación automática de materiales repetidos en el checklist
- validación de columnas requeridas y detección de incidencias
- exportación de resultados a **Excel**
- launcher con actualización automática por **GitHub Releases**

---

## Uso en Windows (Recomendado)

### Descarga

1. Ir a **Releases**:
   https://github.com/linhermx/assemble/releases
2. Descargar:
   **`assemble_launcher.exe`**

> No necesitas instalar Python ni dependencias.

---

### Primer uso

1. Ejecuta `assemble_launcher.exe`
2. El launcher:
   - revisa si hay una versión más reciente
   - pregunta si deseas actualizar
3. Acepta y el sistema se actualiza automáticamente

Después se abre la aplicación principal.

---

### Uso de la aplicación

1. Selecciona el **checklist estándar**
2. Selecciona el **inventario estándar**
3. Selecciona la **carpeta de salida**
4. Define el **objetivo de hornos**
5. Haz clic en **Analizar capacidad**

Salidas generadas:

- `reporte_capacidad_hornos.xlsx`
- `run_log.txt`

---

## Formato de archivos Excel

La aplicación espera dos archivos Excel con hojas y columnas específicas.

### Checklist estándar

Hoja requerida: `checklist`

| Columna | Descripción |
|---|---|
| `modelo_nombre` | Nombre del modelo de horno |
| `seccion` | Área o subconjunto al que pertenece el material |
| `descripcion_material` | Nombre del material |
| `unidad_consumo` | Unidad base del consumo |
| `cantidad_por_horno` | Cantidad requerida para fabricar 1 horno |
| `incluir_en_capacidad` | `SI` o `NO` para definir si limita la capacidad |
| `observaciones` | Campo libre opcional |

### Inventario estándar

Hoja requerida: `inventario`

| Columna | Descripción |
|---|---|
| `descripcion_material` | Nombre del material |
| `unidad` | Unidad base de existencia |
| `existencia` | Existencia disponible |

### Reglas de captura

- `descripcion_material` debe coincidir entre checklist e inventario
- `unidad_consumo` y `unidad` deben coincidir para el mismo material
- `cantidad_por_horno` y `existencia` aceptan enteros y decimales
- para materiales fraccionales se usa la unidad declarada en el archivo
- ejemplo: `0.68` en `Metro` significa `0.68 metros`
- ejemplo: `3.5` en `Metro` significa `3.5 metros`

---

## Reglas de negocio implementadas

- si falta aunque sea un material obligatorio, no sale otro horno completo
- la capacidad final se toma del material con menor capacidad por existencia
- si un material aparece varias veces en el checklist, primero se suma su consumo total por horno
- los materiales con `incluir_en_capacidad = NO` quedan fuera del cálculo
- si un material obligatorio no existe en inventario, se toma `existencia = 0` y se reporta como incidencia
- si checklist e inventario usan unidades distintas para el mismo material, se reporta incidencia
- el cálculo de faltantes para objetivo usa la cantidad total requerida contra la existencia actual

---

## Uso técnico / desarrolladores (CLI)

### Requisitos

- Python **3.10+**
- Windows

Dependencias principales:

- `pandas`
- `openpyxl`
- `xlsxwriter`
- `requests`

---

### Instalación

Se recomienda usar un entorno virtual (`venv`).

```powershell
git clone https://github.com/linhermx/assemble.git
cd assemble

.\scripts\setup_venv.ps1
.\venv\Scripts\Activate.ps1
```

O manualmente:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

---

### Ejecutar GUI en desarrollo

```powershell
.\scripts\run_gui.ps1
```

O directamente:

```powershell
python .\assemble_gui.py
```

### Ejecutar launcher en desarrollo

```powershell
python .\assemble_launcher.py
```

Si no existe una release publicada, el launcher cae localmente a `assemble_gui.py`.

---

### Uso básico por CLI

```powershell
python .\assemble.py `
  --checklist "examples\goliat_premium_standardized\checklist_goliat_premium_estandar.xlsx" `
  --inventory "examples\goliat_premium_standardized\inventario_global_estandar.xlsx" `
  --outdir "salida" `
  --target 10 `
  --overwrite
```

---

## Parámetros CLI

| Parámetro | Descripción |
|---|---|
| `--checklist` | Ruta al checklist estándar |
| `--inventory` | Ruta al inventario estándar |
| `--outdir` | Carpeta de salida. Default: `salida` |
| `--target` | Objetivo de hornos a simular. Default: `1` |
| `--overwrite` | Sobrescribe el reporte si ya existe |

---

## Reporte generado

El archivo `reporte_capacidad_hornos.xlsx` incluye estas hojas:

- `Resumen`
- `Checklist agregado`
- `Limitantes`
- `Faltantes siguiente`
- `Faltantes objetivo`
- `Excluidos`
- `Detalle checklist`
- `Incidencias`

El archivo `run_log.txt` resume:

- modelo analizado
- hornos completos posibles hoy
- siguiente horno objetivo
- material más limitante
- faltantes para el siguiente horno
- faltantes para el objetivo
- incidencias detectadas

---

## Flujo recomendado

1. Preparar el checklist estándar
2. Preparar el inventario estándar
3. Ejecutar la aplicación
4. Revisar el resumen principal
5. Analizar faltantes para el siguiente horno
6. Simular faltantes para el objetivo
7. Descargar y compartir el reporte

---

## Build para Windows

```powershell
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
```

Los binarios se generan en `dist/`:

- `assemble_windows.exe`
- `assemble_launcher.exe`

Ambos scripts crean `.\venv` automáticamente si todavía no existe.
