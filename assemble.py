from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if SRC.exists() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from assemble.core import calculate_capacity


def main():
    parser = argparse.ArgumentParser(
        description="Calcula la capacidad de producción de hornos usando checklist e inventario estándar."
    )
    parser.add_argument("--checklist", required=True, help="Ruta al checklist estándar")
    parser.add_argument("--inventory", required=True, help="Ruta al inventario estándar")
    parser.add_argument("--outdir", default="salida", help="Carpeta de salida")
    parser.add_argument("--target", type=int, default=1, help="Objetivo de hornos a simular")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescribir si el reporte ya existe")
    args = parser.parse_args()

    result = calculate_capacity(
        checklist_path=Path(args.checklist),
        inventory_path=Path(args.inventory),
        output_dir=Path(args.outdir),
        target_ovens=args.target,
        overwrite=args.overwrite,
    )

    print(f"Modelo: {result.model_name}")
    print(f"Hornos completos posibles hoy: {'No determinado' if result.current_capacity is None else result.current_capacity}")
    print(f"Material más limitante: {result.bottleneck_material or 'N/A'}")
    print(f"Materiales faltantes para el siguiente horno: {len(result.next_shortages_frame)}")
    print(f"Materiales faltantes para {result.target_ovens} horno(s): {len(result.target_shortages_frame)}")
    print(f"Reporte: {result.report_file}")
    if result.issues:
        print("\nIncidencias:")
        for issue in result.issues:
            print(f"- [{issue.level.upper()}] {issue.message}")


if __name__ == "__main__":
    main()
