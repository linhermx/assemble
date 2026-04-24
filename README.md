# assemble

Herramienta de escritorio para calcular:

- cuantos hornos completos pueden armarse con el inventario actual
- que material detiene la produccion del siguiente horno
- que piezas faltan para completar el siguiente horno
- que material falta para alcanzar `N` hornos objetivo

Usa el mismo stack base de `barcodes_from_csv_ean13`:

- Python
- `tkinter` para GUI de escritorio
- paquete en `src/`
- launcher con auto-actualizacion por GitHub Releases
- CLI para uso tecnico
- scripts de build para Windows con `PyInstaller`

## Uso en Windows sin Python

Igual que en `barcodes_from_csv_ean13`, para usuarios finales la ruta correcta es:

- `assemble_launcher.exe`
- `assemble_windows.exe`

En ese escenario:

- no se necesita instalar Python
- no se necesita crear `venv`
- el launcher descarga/abre la app publicada

El `venv` se usa solo para desarrollo local y para construir los ejecutables.

## Formato estandar de archivos

### Checklist estandar

Hoja: `checklist`

Columnas requeridas:

- `modelo_nombre`
- `seccion`
- `descripcion_material`
- `unidad_consumo`
- `cantidad_por_horno`
- `incluir_en_capacidad`
- `observaciones`

### Inventario estandar

Hoja: `inventario`

Columnas requeridas:

- `descripcion_material`
- `unidad`
- `existencia`

## Reglas de negocio implementadas

- Si falta aunque sea un material obligatorio, no sale otro horno completo.
- El calculo se hace por material agregado.
- Si un material aparece varias veces en el checklist, primero se suma su consumo total por horno.
- Los materiales con `incluir_en_capacidad = NO` no limitan la capacidad.
- Para materiales fraccionales se usa la unidad base declarada en el archivo.
  - Ejemplo: `0.68` en `Metro` significa `0.68 metros`.
  - Ejemplo: `3.5` en `Metro` significa `3.5 metros`.

## Ejecutar GUI

```bash
pip install -r requirements.txt
python assemble_gui.py
```

## Desarrollo local con venv

```powershell
.\scripts\setup_venv.ps1
.\venv\Scripts\Activate.ps1
python .\assemble_gui.py
```

O en una sola instrucción:

```powershell
.\scripts\run_gui.ps1
```

## Ejecutar launcher

```bash
python assemble_launcher.py
```

Si no existe release publicada, el launcher cae localmente a `assemble_gui.py`.

## Ejecutar CLI

```bash
python assemble.py ^
  --checklist "examples\\goliat_premium_standardized\\checklist_goliat_premium_estandar.xlsx" ^
  --inventory "examples\\goliat_premium_standardized\\inventario_global_estandar.xlsx" ^
  --outdir "salida" ^
  --target 10
```

## Salidas generadas

La herramienta crea:

- `reporte_capacidad_hornos.xlsx`
- `run_log.txt`

Y responde estas preguntas:

- cuantos hornos completos salen hoy
- que material bloquea el siguiente horno
- que materiales faltan para completar el siguiente horno
- que faltantes existen para llegar a `N` hornos objetivo

## Build para Windows

```bash
.\scripts\build_windows.ps1
.\scripts\build_launcher.ps1
```

Ambos scripts crean `.\venv` automáticamente si todavía no existe.
