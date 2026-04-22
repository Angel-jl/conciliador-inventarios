# -*- coding: utf-8 -*-
"""
Conciliación de inventarios por columna clave (UPC, SKU, EAN, etc.)

Compara dos fuentes de inventario (hojas de un mismo Excel o archivos distintos)
e identifica diferencias, faltantes y sobrantes por unidad de comparación.

Uso:
    python conciliar_inventarios.py                  # Modo interactivo (recomendado)
    python conciliar_inventarios.py --config          # Muestra ejemplo de configuración
"""

import os
import sys
import argparse
import unicodedata
import pandas as pd


# ─────────────────────────────────────────────
#  CONFIGURACIÓN — edita aquí si prefieres no
#  usar el modo interactivo
# ─────────────────────────────────────────────
# Deja en None para que el script lo pida en pantalla.

GEN_PATH   = None   # Ruta al archivo de inventario GENERAL  (Excel o CSV)
CLI_PATH   = None   # Ruta al archivo de inventario CLIENTES (puede ser el mismo)
GEN_SHEET  = None   # Nombre de la hoja del inventario GENERAL   (None = primera hoja)
CLI_SHEET  = None   # Nombre de la hoja del inventario CLIENTES  (None = primera hoja)

# Columna clave de comparación (UPC, SKU, EAN, Código, etc.)
FORCED_KEY_COL      = None   # Ej: "UPC"  — None para detección automática

# Columna de cantidad en cada fuente
FORCED_TOTAL_COL_GEN = None  # Ej: "TOTAL FISICO"
FORCED_TOTAL_COL_CLI = None  # Ej: "TOTAL FISICO"

# Ruta del archivo de salida
OUTPUT_PATH = None  # None = se guarda en la misma carpeta del script
# ─────────────────────────────────────────────


# ══════════════════════════════════════════════
#  UTILIDADES DE TEXTO Y DETECCIÓN DE COLUMNAS
# ══════════════════════════════════════════════

def strip_accents_lower(s: str) -> str:
    """Normaliza texto: quita acentos y convierte a minúsculas."""
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def find_col(df: pd.DataFrame, candidates: list) -> str:
    """
    Busca la primera columna cuyo nombre normalizado coincida con algún candidato.
    Tolera diferencias de acentos y mayúsculas.
    """
    norm_map = {col: strip_accents_lower(col) for col in df.columns}
    cand_norm = [strip_accents_lower(c) for c in candidates]
    for original, normalized in norm_map.items():
        if normalized in cand_norm:
            return original
    raise KeyError(
        f"\n  ✗ No se encontró ninguna de estas columnas: {candidates}"
        f"\n  Columnas disponibles: {list(df.columns)}"
        f"\n  → Especifica el nombre exacto con la opción FORCED_KEY_COL o en el modo interactivo."
    )


def normalize_key(series: pd.Series) -> pd.Series:
    """Limpia la columna clave: quita espacios, elimina nulos y vacíos."""
    s = series.astype(str).str.strip().str.replace(r"\s+", "", regex=True)
    return s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def normalize_qty(series: pd.Series) -> pd.Series:
    """Convierte cantidades a entero, reemplaza NaN por 0."""
    return pd.to_numeric(series, errors="coerce").fillna(0).astype("int64")


# ══════════════════════════════════════════════
#  MODO INTERACTIVO
# ══════════════════════════════════════════════

def ask(prompt: str, default: str = None) -> str:
    """Pide un valor al usuario. Si hay default y el usuario presiona Enter, usa el default."""
    display = f"{prompt}"
    if default:
        display += f" [{default}]"
    display += ": "
    resp = input(display).strip()
    return resp if resp else (default or "")


def listar_hojas(path: str) -> list:
    """Devuelve las hojas disponibles de un Excel."""
    try:
        xl = pd.ExcelFile(path)
        return xl.sheet_names
    except Exception:
        return []


def modo_interactivo() -> dict:
    """
    Guía al usuario paso a paso para configurar la conciliación
    sin editar el código.
    """
    print("\n" + "═" * 60)
    print("  CONCILIADOR DE INVENTARIOS — Modo Interactivo")
    print("═" * 60)
    print("  Compara dos fuentes de inventario y detecta diferencias.")
    print("  Presiona Enter para aceptar el valor entre [corchetes].\n")

    cfg = {}

    # ── Archivo GENERAL ──────────────────────────────────────
    print("─── FUENTE 1: Inventario General (almacén/físico) ───")
    while True:
        cfg["gen_path"] = ask("  Ruta del archivo (.xlsx, .xlsm, .csv)")
        if os.path.exists(cfg["gen_path"]):
            break
        print(f"  ✗ Archivo no encontrado. Intenta de nuevo.")

    hojas_gen = listar_hojas(cfg["gen_path"])
    if hojas_gen:
        print(f"  Hojas disponibles: {hojas_gen}")
        cfg["gen_sheet"] = ask("  Nombre de la hoja", default=hojas_gen[0])
    else:
        cfg["gen_sheet"] = None  # CSV — sin hojas

    # ── Archivo CLIENTES ─────────────────────────────────────
    print("\n─── FUENTE 2: Inventario Clientes (sistema/partidas) ───")
    mismo = ask("  ¿Es el mismo archivo que la Fuente 1? (s/n)", default="n").lower()

    if mismo == "s":
        cfg["cli_path"] = cfg["gen_path"]
        hojas_cli = hojas_gen
    else:
        while True:
            cfg["cli_path"] = ask("  Ruta del archivo (.xlsx, .xlsm, .csv)")
            if os.path.exists(cfg["cli_path"]):
                break
            print(f"  ✗ Archivo no encontrado. Intenta de nuevo.")
        hojas_cli = listar_hojas(cfg["cli_path"])

    if hojas_cli:
        print(f"  Hojas disponibles: {hojas_cli}")
        default_cli = hojas_cli[1] if len(hojas_cli) > 1 else hojas_cli[0]
        cfg["cli_sheet"] = ask("  Nombre de la hoja", default=default_cli)
    else:
        cfg["cli_sheet"] = None

    # ── Columna clave ─────────────────────────────────────────
    print("\n─── COLUMNA CLAVE DE COMPARACIÓN ───")
    print("  (La columna que identifica cada producto: UPC, SKU, EAN, Código, etc.)")
    cfg["key_col"] = ask("  Nombre de la columna clave", default="UPC")

    # ── Columna de cantidad ───────────────────────────────────
    print("\n─── COLUMNA DE CANTIDAD ───")
    cfg["total_col_gen"] = ask("  Nombre de la columna de cantidad en Fuente 1", default="TOTAL FISICO")
    cfg["total_col_cli"] = ask("  Nombre de la columna de cantidad en Fuente 2", default="TOTAL FISICO")

    # ── Salida ────────────────────────────────────────────────
    print("\n─── ARCHIVO DE SALIDA ───")
    default_out = os.path.join(os.path.dirname(cfg["gen_path"]), "conciliacion_resultado.xlsx")
    cfg["output_path"] = ask("  Ruta del Excel de salida", default=default_out)

    print("\n" + "═" * 60)
    return cfg


# ══════════════════════════════════════════════
#  CARGA Y PROCESAMIENTO
# ══════════════════════════════════════════════

def load_and_prepare(path: str, sheet, key_col_forced, total_col_forced, label: str) -> pd.DataFrame:
    """
    Carga un archivo de inventario (Excel o CSV), detecta columnas y agrupa por clave.
    """
    print(f"\n  Cargando {label}...")

    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    else:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)

    print(f"  ✓ {len(df):,} filas cargadas | Columnas: {list(df.columns)}")

    # Detectar columna clave
    if key_col_forced and key_col_forced in df.columns:
        key_col = key_col_forced
    else:
        key_col = find_col(df, candidates=[
            key_col_forced or "UPC",
            "upc", "ean", "sku", "codigo", "código", "clave", "item"
        ])

    # Detectar columna de cantidad
    if total_col_forced and total_col_forced in df.columns:
        total_col = total_col_forced
    else:
        total_col = find_col(df, candidates=[
            total_col_forced or "TOTAL FISICO",
            "total fisico", "total físico", "total", "cantidad",
            "qty", "existencia", "pieza", "piezas", "unidades"
        ])

    print(f"  Columna clave detectada   : '{key_col}'")
    print(f"  Columna cantidad detectada: '{total_col}'")

    df = df[[key_col, total_col]].copy()
    df.rename(columns={key_col: "CLAVE", total_col: "TOTAL"}, inplace=True)
    df["CLAVE"] = normalize_key(df["CLAVE"])
    df = df.dropna(subset=["CLAVE"])
    df["TOTAL"] = normalize_qty(df["TOTAL"])

    grouped = (
        df.groupby("CLAVE", as_index=False)["TOTAL"]
        .sum()
    )
    print(f"  Claves únicas agrupadas   : {len(grouped):,}")
    return grouped


def conciliacion(gen: pd.DataFrame, cli: pd.DataFrame) -> pd.DataFrame:
    """
    Hace el merge outer por CLAVE y calcula diferencias y estado.
    DIFERENCIA = GENERAL - CLIENTES
    """
    final = pd.merge(
        gen.rename(columns={"TOTAL": "QTY GENERAL"}),
        cli.rename(columns={"TOTAL": "QTY CLIENTES"}),
        on="CLAVE",
        how="outer",
    ).fillna(0)

    final["QTY GENERAL"]  = final["QTY GENERAL"].astype("int64")
    final["QTY CLIENTES"] = final["QTY CLIENTES"].astype("int64")
    final["DIFERENCIA"]   = final["QTY GENERAL"] - final["QTY CLIENTES"]
    final["ESTADO"] = final["DIFERENCIA"].apply(
        lambda x: "✅ OK" if x == 0 else ("⚠️ FALTANTE" if x > 0 else "🔴 SOBRANTE")
    )

    return final.sort_values(["ESTADO", "CLAVE"]).reset_index(drop=True)


def build_non_crossed(gen: pd.DataFrame, cli: pd.DataFrame):
    """Separa los registros que solo existen en una fuente."""
    set_gen = set(gen["CLAVE"])
    set_cli = set(cli["CLAVE"])

    solo_gen = gen[gen["CLAVE"].isin(set_gen - set_cli)].rename(
        columns={"TOTAL": "QTY GENERAL"}
    ).sort_values("CLAVE").reset_index(drop=True)

    solo_cli = cli[cli["CLAVE"].isin(set_cli - set_gen)].rename(
        columns={"TOTAL": "QTY CLIENTES"}
    ).sort_values("CLAVE").reset_index(drop=True)

    return solo_gen, solo_cli


def to_excel(final, gen, cli, solo_gen, solo_cli, output_path: str):
    """Exporta todos los resultados a un Excel con hojas organizadas."""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final.to_excel(writer, index=False, sheet_name="Conciliacion")
        gen.rename(columns={"TOTAL": "QTY GENERAL"}).sort_values("CLAVE").to_excel(
            writer, index=False, sheet_name="General Agrupado"
        )
        cli.rename(columns={"TOTAL": "QTY CLIENTES"}).sort_values("CLAVE").to_excel(
            writer, index=False, sheet_name="Clientes Agrupado"
        )
        solo_gen.to_excel(writer, index=False, sheet_name="Solo en General")
        solo_cli.to_excel(writer, index=False, sheet_name="Solo en Clientes")

        resumen = pd.DataFrame({
            "Métrica": [
                "Claves en General",
                "Claves en Clientes",
                "Claves cruzadas (en ambos)",
                "Claves solo en General",
                "Claves solo en Clientes",
                "Registros OK (sin diferencia)",
                "Registros con FALTANTE (General > Clientes)",
                "Registros con SOBRANTE (Clientes > General)",
                "Suma total de DIFERENCIA",
            ],
            "Valor": [
                gen["CLAVE"].nunique(),
                cli["CLAVE"].nunique(),
                len(set(gen["CLAVE"]) & set(cli["CLAVE"])),
                len(solo_gen),
                len(solo_cli),
                len(final[final["ESTADO"] == "✅ OK"]),
                len(final[final["ESTADO"] == "⚠️ FALTANTE"]),
                len(final[final["ESTADO"] == "🔴 SOBRANTE"]),
                int(final["DIFERENCIA"].sum()),
            ],
        })
        resumen.to_excel(writer, index=False, sheet_name="Resumen")


# ══════════════════════════════════════════════
#  PUNTO DE ENTRADA
# ══════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Conciliador de inventarios por columna clave.",
        add_help=True
    )
    parser.add_argument("--config", action="store_true",
                        help="Muestra un ejemplo de configuración y sale.")
    args = parser.parse_args()

    if args.config:
        print("""
Ejemplo de configuración hardcodeada (edita las variables al inicio del script):

    GEN_PATH   = r"C:\\ruta\\a\\inventario_general.xlsx"
    CLI_PATH   = r"C:\\ruta\\a\\inventario_clientes.xlsx"
    GEN_SHEET  = "Hoja1"
    CLI_SHEET  = "Hoja2"
    FORCED_KEY_COL       = "UPC"
    FORCED_TOTAL_COL_GEN = "TOTAL FISICO"
    FORCED_TOTAL_COL_CLI = "TOTAL FISICO"
    OUTPUT_PATH = r"C:\\ruta\\salida\\conciliacion_resultado.xlsx"
""")
        sys.exit(0)

    # Decidir si usar modo interactivo o variables hardcodeadas
    if GEN_PATH is None:
        cfg = modo_interactivo()
        gen_path        = cfg["gen_path"]
        cli_path        = cfg["cli_path"]
        gen_sheet       = cfg["gen_sheet"]
        cli_sheet       = cfg["cli_sheet"]
        key_col         = cfg["key_col"]
        total_col_gen   = cfg["total_col_gen"]
        total_col_cli   = cfg["total_col_cli"]
        output_path     = cfg["output_path"]
    else:
        gen_path        = GEN_PATH
        cli_path        = CLI_PATH
        gen_sheet       = GEN_SHEET
        cli_sheet       = CLI_SHEET
        key_col         = FORCED_KEY_COL
        total_col_gen   = FORCED_TOTAL_COL_GEN
        total_col_cli   = FORCED_TOTAL_COL_CLI
        output_path     = OUTPUT_PATH or "conciliacion_resultado.xlsx"

    # Pipeline
    gen = load_and_prepare(gen_path, gen_sheet, key_col, total_col_gen, "Inventario General")
    cli = load_and_prepare(cli_path, cli_sheet, key_col, total_col_cli, "Inventario Clientes")

    final = conciliacion(gen, cli)
    solo_gen, solo_cli = build_non_crossed(gen, cli)

    to_excel(final, gen, cli, solo_gen, solo_cli, output_path)

    print("\n" + "═" * 60)
    print("  RESULTADO DE CONCILIACIÓN")
    print("═" * 60)
    print(f"  Claves en General        : {gen['CLAVE'].nunique():,}")
    print(f"  Claves en Clientes       : {cli['CLAVE'].nunique():,}")
    print(f"  Claves cruzadas          : {len(set(gen['CLAVE']) & set(cli['CLAVE'])):,}")
    print(f"  Solo en General          : {len(solo_gen):,}")
    print(f"  Solo en Clientes         : {len(solo_cli):,}")
    print(f"  ✅ OK (sin diferencia)   : {len(final[final['ESTADO'] == '✅ OK']):,}")
    print(f"  ⚠️  FALTANTE             : {len(final[final['ESTADO'] == '⚠️ FALTANTE']):,}")
    print(f"  🔴 SOBRANTE             : {len(final[final['ESTADO'] == '🔴 SOBRANTE']):,}")
    print(f"  Suma total DIFERENCIA    : {int(final['DIFERENCIA'].sum()):,}")
    print(f"\n  Archivo generado: {output_path}")
    print("═" * 60)


if __name__ == "__main__":
    main()
