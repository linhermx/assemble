from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


CHECKLIST_SHEET = "checklist"
INVENTORY_SHEET = "inventario"

REQUIRED_CHECKLIST_COLUMNS = [
    "modelo_nombre",
    "seccion",
    "descripcion_material",
    "unidad_consumo",
    "cantidad_por_horno",
    "incluir_en_capacidad",
    "observaciones",
]

REQUIRED_INVENTORY_COLUMNS = [
    "descripcion_material",
    "unidad",
    "existencia",
]


@dataclass
class RunIssue:
    level: str
    message: str


@dataclass
class RunResult:
    model_name: str
    target_ovens: int
    current_capacity: int | None
    next_oven_target: int | None
    bottleneck_material: str
    issues: list[RunIssue]
    summary_frame: pd.DataFrame
    aggregated_checklist_frame: pd.DataFrame
    limiting_frame: pd.DataFrame
    next_shortages_frame: pd.DataFrame
    target_shortages_frame: pd.DataFrame
    excluded_frame: pd.DataFrame
    detail_frame: pd.DataFrame
    report_file: Path
    log_file: Path
    output_dir: Path


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    return " ".join(text.strip().upper().split())


def as_float(value: object) -> float | None:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_yes_no(value: object) -> str:
    text = normalize_text(value)
    if text in {"SI", "NO"}:
        return text
    return ""


def load_table(path: str | Path, sheet_name: str) -> pd.DataFrame:
    workbook = pd.read_excel(path, sheet_name=None)
    if sheet_name in workbook:
        frame = workbook[sheet_name]
    else:
        frame = next(iter(workbook.values()))
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    frame = frame.dropna(how="all").reset_index(drop=True)
    return frame


def validate_required_columns(frame: pd.DataFrame, required_columns: list[str], label: str) -> list[RunIssue]:
    missing = [column for column in required_columns if column not in frame.columns]
    if not missing:
        return []
    return [RunIssue("error", f"{label}: faltan columnas obligatorias: {missing}")]


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    counter = 2
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def write_report(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        for sheet_name, frame in sheets.items():
            export = frame.copy()
            export.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            for column_index, column_name in enumerate(export.columns):
                max_length = max(
                    len(str(column_name)),
                    export[column_name].astype(str).map(len).max() if not export.empty else 0,
                )
                worksheet.set_column(column_index, column_index, min(max_length + 2, 42))


def aggregate_checklist(checklist: pd.DataFrame, issues: list[RunIssue]) -> tuple[pd.DataFrame, pd.DataFrame]:
    frame = checklist.copy()
    frame["fila_origen"] = frame.index + 2
    frame["descripcion_material"] = frame["descripcion_material"].astype(str).str.strip()
    frame["unidad_consumo"] = frame["unidad_consumo"].astype(str).str.strip()
    frame["modelo_nombre"] = frame["modelo_nombre"].astype(str).str.strip()
    frame["seccion"] = frame["seccion"].astype(str).str.strip()
    frame["observaciones"] = frame["observaciones"].fillna("").astype(str).str.strip()
    frame["incluir_en_capacidad"] = frame["incluir_en_capacidad"].apply(parse_yes_no)
    frame["cantidad_por_horno_num"] = frame["cantidad_por_horno"].apply(as_float)
    frame["descripcion_normalizada"] = frame["descripcion_material"].apply(normalize_text)
    frame["unidad_normalizada"] = frame["unidad_consumo"].apply(normalize_text)

    invalid_yes_no = frame[frame["incluir_en_capacidad"] == ""]
    for _, row in invalid_yes_no.iterrows():
        issues.append(
            RunIssue(
                "error",
                f"Checklist fila {int(row['fila_origen'])}: 'incluir_en_capacidad' debe ser SI o NO.",
            )
        )

    invalid_qty = frame[frame["cantidad_por_horno_num"].isna()]
    for _, row in invalid_qty.iterrows():
        issues.append(
            RunIssue(
                "error",
                f"Checklist fila {int(row['fila_origen'])}: 'cantidad_por_horno' no es numérica.",
            )
        )

    model_names = [value for value in frame["modelo_nombre"].dropna().unique().tolist() if str(value).strip()]
    if len(model_names) > 1:
        issues.append(
            RunIssue(
                "warning",
                f"El checklist contiene varios modelos: {model_names}. Se usará el primero para el resumen.",
            )
        )

    detail_frame = frame[
        [
            "fila_origen",
            "modelo_nombre",
            "seccion",
            "descripcion_material",
            "unidad_consumo",
            "cantidad_por_horno_num",
            "incluir_en_capacidad",
            "observaciones",
        ]
    ].rename(
        columns={
            "fila_origen": "Fila origen",
            "modelo_nombre": "Modelo",
            "seccion": "Sección",
            "descripcion_material": "Material",
            "unidad_consumo": "Unidad",
            "cantidad_por_horno_num": "Cantidad por horno",
            "incluir_en_capacidad": "Incluye en capacidad",
            "observaciones": "Observaciones",
        }
    )

    aggregated = (
        frame.groupby(["descripcion_normalizada", "unidad_normalizada", "incluir_en_capacidad"], dropna=False, as_index=False)
        .agg(
            {
                "modelo_nombre": "first",
                "descripcion_material": "first",
                "unidad_consumo": "first",
                "cantidad_por_horno_num": "sum",
                "seccion": lambda values: " | ".join(sorted({value for value in values if value})),
                "observaciones": lambda values: " | ".join([value for value in values if value]),
            }
        )
    )

    duplicated_units = aggregated.groupby("descripcion_normalizada")["unidad_normalizada"].nunique()
    problematic = duplicated_units[duplicated_units > 1].index.tolist()
    for description in problematic:
        subset = aggregated[aggregated["descripcion_normalizada"] == description]
        issues.append(
            RunIssue(
                "error",
                "El checklist usa diferentes unidades para el mismo material: "
                + ", ".join(subset["descripcion_material"].tolist()),
            )
        )

    aggregated = aggregated.rename(
        columns={
            "modelo_nombre": "modelo_nombre",
            "descripcion_material": "descripcion_material",
            "unidad_consumo": "unidad_consumo",
            "cantidad_por_horno_num": "cantidad_por_horno",
            "seccion": "secciones",
            "observaciones": "observaciones",
        }
    )
    return aggregated, detail_frame


def aggregate_inventory(inventory: pd.DataFrame, issues: list[RunIssue]) -> pd.DataFrame:
    frame = inventory.copy()
    frame["fila_origen"] = frame.index + 2
    frame["descripcion_material"] = frame["descripcion_material"].astype(str).str.strip()
    frame["unidad"] = frame["unidad"].astype(str).str.strip()
    frame["existencia_num"] = frame["existencia"].apply(as_float)
    frame["descripcion_normalizada"] = frame["descripcion_material"].apply(normalize_text)
    frame["unidad_normalizada"] = frame["unidad"].apply(normalize_text)

    invalid_qty = frame[frame["existencia_num"].isna()]
    for _, row in invalid_qty.iterrows():
        issues.append(
            RunIssue(
                "error",
                f"Inventario fila {int(row['fila_origen'])}: 'existencia' no es numérica.",
            )
        )

    aggregated = (
        frame.groupby(["descripcion_normalizada", "unidad_normalizada"], dropna=False, as_index=False)
        .agg(
            {
                "descripcion_material": "first",
                "unidad": "first",
                "existencia_num": "sum",
            }
        )
    )

    duplicated_units = aggregated.groupby("descripcion_normalizada")["unidad_normalizada"].nunique()
    problematic = duplicated_units[duplicated_units > 1].index.tolist()
    for description in problematic:
        subset = aggregated[aggregated["descripcion_normalizada"] == description]
        issues.append(
            RunIssue(
                "error",
                "El inventario usa diferentes unidades para el mismo material: "
                + ", ".join(subset["descripcion_material"].tolist()),
            )
        )

    aggregated = aggregated.rename(
        columns={
            "descripcion_material": "descripcion_material",
            "unidad": "unidad",
            "existencia_num": "existencia",
        }
    )
    return aggregated


def calculate_capacity(
    checklist_path: str | Path,
    inventory_path: str | Path,
    output_dir: str | Path,
    target_ovens: int,
    overwrite: bool = True,
) -> RunResult:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_file, log_file = build_output_paths(output_dir=output_dir, overwrite=overwrite)

    issues: list[RunIssue] = []
    checklist_raw = load_table(checklist_path, CHECKLIST_SHEET)
    inventory_raw = load_table(inventory_path, INVENTORY_SHEET)

    issues.extend(validate_required_columns(checklist_raw, REQUIRED_CHECKLIST_COLUMNS, "Checklist"))
    issues.extend(validate_required_columns(inventory_raw, REQUIRED_INVENTORY_COLUMNS, "Inventario"))
    if any(issue.level == "error" for issue in issues):
        return _empty_result(output_dir, target_ovens, issues, report_file=report_file, log_file=log_file)

    checklist_agg, detail_frame = aggregate_checklist(checklist_raw, issues)
    inventory_agg = aggregate_inventory(inventory_raw, issues)

    included = checklist_agg[checklist_agg["incluir_en_capacidad"] == "SI"].copy()
    excluded = checklist_agg[checklist_agg["incluir_en_capacidad"] == "NO"].copy()

    inventory_lookup = {
        (str(row["descripcion_normalizada"]), str(row["unidad_normalizada"])): row
        for _, row in inventory_agg.iterrows()
    }
    inventory_by_description: dict[str, list[dict[str, object]]] = {}
    for _, row in inventory_agg.iterrows():
        inventory_by_description.setdefault(str(row["descripcion_normalizada"]), []).append(row.to_dict())

    rows = []
    for _, row in included.iterrows():
        desc_key = str(row["descripcion_normalizada"])
        unit_key = str(row["unidad_normalizada"])
        lookup = inventory_lookup.get((desc_key, unit_key))
        inventory_description = ""
        inventory_unit = ""
        inventory_existence = 0.0

        if lookup is None:
            alternatives = inventory_by_description.get(desc_key, [])
            if alternatives:
                inventory_unit = str(alternatives[0]["unidad"])
                issues.append(
                    RunIssue(
                        "error",
                        f"Unidad incompatible para {row['descripcion_material']}: checklist={row['unidad_consumo']} inventario={inventory_unit}.",
                    )
                )
            else:
                issues.append(
                    RunIssue(
                        "warning",
                        f"No se encontró en inventario el material obligatorio: {row['descripcion_material']}. Se toma existencia = 0.",
                    )
                )
        else:
            inventory_description = str(lookup["descripcion_material"])
            inventory_unit = str(lookup["unidad"])
            inventory_existence = float(lookup["existencia"] or 0.0)

        quantity_per_oven = float(row["cantidad_por_horno"] or 0.0)
        capacity_by_material = (
            math.floor(inventory_existence / quantity_per_oven) if quantity_per_oven > 0 else None
        )

        rows.append(
            {
                "Modelo": row["modelo_nombre"],
                "Secciones": row["secciones"],
                "Material": row["descripcion_material"],
                "Unidad": row["unidad_consumo"],
                "Cantidad por horno": quantity_per_oven,
                "Existencia actual": inventory_existence,
                "Capacidad por material": capacity_by_material,
                "Observaciones": row["observaciones"],
            }
        )

    aggregated_checklist_frame = pd.DataFrame(rows)

    capacity = None
    if not aggregated_checklist_frame.empty:
        valid_capacities = aggregated_checklist_frame["Capacidad por material"].dropna()
        if not valid_capacities.empty:
            capacity = int(valid_capacities.min())

    fatal_errors = [issue for issue in issues if issue.level == "error"]
    if fatal_errors:
        capacity = None

    next_target = None if capacity is None else capacity + 1

    limiting_frame = aggregated_checklist_frame.sort_values(
        by=["Capacidad por material", "Existencia actual", "Material"],
        ascending=[True, True, True],
    ).reset_index(drop=True)

    bottleneck_material = ""
    if capacity is not None and not limiting_frame.empty:
        item = limiting_frame.iloc[0]
        bottleneck_material = (
            f"{item['Material']} | requiere {item['Cantidad por horno']} {item['Unidad']} | "
            f"hay {item['Existencia actual']} {item['Unidad']}"
        )

    next_shortages_frame = _build_shortage_frame(aggregated_checklist_frame, next_target, "Faltante para siguiente horno")
    target_shortages_frame = _build_shortage_frame(aggregated_checklist_frame, target_ovens, "Faltante para objetivo")

    excluded_frame = excluded[
        ["modelo_nombre", "secciones", "descripcion_material", "unidad_consumo", "cantidad_por_horno", "observaciones"]
    ].rename(
        columns={
            "modelo_nombre": "Modelo",
            "secciones": "Secciones",
            "descripcion_material": "Material",
            "unidad_consumo": "Unidad",
            "cantidad_por_horno": "Cantidad por horno",
            "observaciones": "Observaciones",
        }
    ).reset_index(drop=True)

    model_name = (
        str(checklist_raw["modelo_nombre"].dropna().iloc[0]).strip()
        if "modelo_nombre" in checklist_raw.columns and not checklist_raw["modelo_nombre"].dropna().empty
        else "MODELO_SIN_NOMBRE"
    )

    summary_frame = pd.DataFrame(
        [
            {"Indicador": "Modelo", "Valor": model_name},
            {"Indicador": "Hornos completos posibles hoy", "Valor": "No determinado" if capacity is None else capacity},
            {"Indicador": "Siguiente horno objetivo", "Valor": "No determinado" if next_target is None else next_target},
            {"Indicador": "Material más limitante", "Valor": bottleneck_material or "N/A"},
            {"Indicador": "Materiales faltantes para el siguiente horno", "Valor": len(next_shortages_frame)},
            {"Indicador": f"Materiales faltantes para {target_ovens} horno(s)", "Valor": len(target_shortages_frame)},
            {"Indicador": "Materiales excluidos del cálculo", "Valor": len(excluded_frame)},
            {"Indicador": "Incidencias detectadas", "Valor": len(issues)},
        ]
    )

    issues_frame = pd.DataFrame(
        [{"Nivel": issue.level.upper(), "Mensaje": issue.message} for issue in issues]
    )

    write_report(
        report_file,
        {
            "Resumen": summary_frame,
            "Checklist agregado": aggregated_checklist_frame,
            "Limitantes": limiting_frame,
            "Faltantes siguiente": next_shortages_frame,
            "Faltantes objetivo": target_shortages_frame,
            "Excluidos": excluded_frame,
            "Detalle checklist": detail_frame,
            "Incidencias": issues_frame,
        },
    )

    with log_file.open("w", encoding="utf-8") as handle:
        handle.write(f"Modelo: {model_name}\n")
        handle.write(f"Hornos completos posibles hoy: {'No determinado' if capacity is None else capacity}\n")
        handle.write(f"Siguiente horno objetivo: {'No determinado' if next_target is None else next_target}\n")
        handle.write(f"Material más limitante: {bottleneck_material or 'N/A'}\n")
        handle.write(f"Materiales faltantes para el siguiente horno: {len(next_shortages_frame)}\n")
        handle.write(f"Materiales faltantes para {target_ovens} horno(s): {len(target_shortages_frame)}\n")
        handle.write(f"Excluidos: {len(excluded_frame)}\n")
        handle.write(f"Reporte: {report_file}\n")
        handle.write(f"Incidencias: {len(issues)}\n\n")
        for issue in issues:
            handle.write(f"[{issue.level.upper()}] {issue.message}\n")

    return RunResult(
        model_name=model_name,
        target_ovens=int(target_ovens),
        current_capacity=capacity,
        next_oven_target=next_target,
        bottleneck_material=bottleneck_material,
        issues=issues,
        summary_frame=summary_frame,
        aggregated_checklist_frame=aggregated_checklist_frame,
        limiting_frame=limiting_frame,
        next_shortages_frame=next_shortages_frame,
        target_shortages_frame=target_shortages_frame,
        excluded_frame=excluded_frame,
        detail_frame=detail_frame,
        report_file=report_file,
        log_file=log_file,
        output_dir=output_dir,
    )


def _build_shortage_frame(aggregated_frame: pd.DataFrame, target: int | None, shortage_column_name: str) -> pd.DataFrame:
    if target is None:
        return pd.DataFrame(
            columns=[
                "Material",
                "Unidad",
                "Cantidad por horno",
                "Existencia actual",
                "Requerido total",
                shortage_column_name,
            ]
        )

    frame = aggregated_frame.copy()
    frame["Requerido total"] = frame["Cantidad por horno"] * int(target)
    frame[shortage_column_name] = (frame["Requerido total"] - frame["Existencia actual"]).clip(lower=0)
    frame = frame[frame[shortage_column_name] > 0][
        [
            "Material",
            "Unidad",
            "Cantidad por horno",
            "Existencia actual",
            "Requerido total",
            shortage_column_name,
        ]
    ].sort_values(by=[shortage_column_name, "Material"], ascending=[False, True]).reset_index(drop=True)
    return frame


def build_output_paths(output_dir: Path, overwrite: bool) -> tuple[Path, Path]:
    report_file = output_dir / "reporte_capacidad_hornos.xlsx"
    log_file = output_dir / "run_log.txt"
    if overwrite:
        return report_file, log_file
    return unique_path(report_file), unique_path(log_file)


def _empty_result(
    output_dir: Path,
    target_ovens: int,
    issues: list[RunIssue],
    *,
    report_file: Path,
    log_file: Path,
) -> RunResult:
    summary = pd.DataFrame(
        [
            {"Indicador": "Modelo", "Valor": "MODELO_SIN_NOMBRE"},
            {"Indicador": "Hornos completos posibles hoy", "Valor": "No determinado"},
            {"Indicador": "Siguiente horno objetivo", "Valor": "No determinado"},
            {"Indicador": "Material más limitante", "Valor": "N/A"},
            {"Indicador": "Incidencias detectadas", "Valor": len(issues)},
        ]
    )
    issues_frame = pd.DataFrame(
        [{"Nivel": issue.level.upper(), "Mensaje": issue.message} for issue in issues]
    )
    write_report(
        report_file,
        {
            "Resumen": summary,
            "Incidencias": issues_frame,
        },
    )
    with log_file.open("w", encoding="utf-8") as handle:
        handle.write("Modelo: MODELO_SIN_NOMBRE\n")
        handle.write("Hornos completos posibles hoy: No determinado\n")
        handle.write("Siguiente horno objetivo: No determinado\n")
        handle.write(f"Incidencias: {len(issues)}\n\n")
        for issue in issues:
            handle.write(f"[{issue.level.upper()}] {issue.message}\n")

    empty = pd.DataFrame()
    return RunResult(
        model_name="MODELO_SIN_NOMBRE",
        target_ovens=int(target_ovens),
        current_capacity=None,
        next_oven_target=None,
        bottleneck_material="",
        issues=issues,
        summary_frame=summary,
        aggregated_checklist_frame=empty,
        limiting_frame=empty,
        next_shortages_frame=empty,
        target_shortages_frame=empty,
        excluded_frame=empty,
        detail_frame=empty,
        report_file=report_file,
        log_file=log_file,
        output_dir=output_dir,
    )
