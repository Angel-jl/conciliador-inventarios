# 🔄 Conciliador de Inventarios — Por Columna Clave

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-150458?style=flat-square&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Excel](https://img.shields.io/badge/Output-Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)]()
[![Retail](https://img.shields.io/badge/Industria-Retail_/_Logística-orange?style=flat-square)]()

> Compara dos fuentes de inventario (almacén vs sistema, físico vs ERP, cliente A vs cliente B)  
> e identifica diferencias, faltantes y sobrantes por unidad de comparación (UPC, SKU, EAN, etc.).

---

## 📌 Contexto del Problema

En operaciones de Retail y logística es común tener **múltiples versiones del inventario**:
el inventario físico del almacén, el sistema del cliente, el ERP, el reporte del transportista, etc.

Cruzar estas fuentes manualmente en Excel —especialmente a nivel talla o UPC— es lento,
propenso a errores y no escala cuando hay miles de referencias.

Este script automatiza ese proceso completamente.

---

## ✨ ¿Qué hace?

1. **Lee dos fuentes de inventario** (hojas de un mismo Excel, archivos distintos, o CSVs)
2. **Agrupa por columna clave** (UPC, SKU, EAN, Código — tú eliges cuál)
3. **Cruza ambas fuentes** y calcula la diferencia por registro
4. **Clasifica** cada registro como OK, FALTANTE o SOBRANTE
5. **Exporta un Excel** con 6 hojas organizadas listas para revisar o cargar al ERP

---

## 🗂️ Archivos de Salida

El script genera un único Excel con las siguientes hojas:

| Hoja | Contenido |
|---|---|
| **Conciliacion** | Vista completa: CLAVE / QTY GENERAL / QTY CLIENTES / DIFERENCIA / ESTADO |
| **General Agrupado** | Inventario de Fuente 1 agrupado por clave |
| **Clientes Agrupado** | Inventario de Fuente 2 agrupado por clave |
| **Solo en General** | Claves que existen en Fuente 1 pero no en Fuente 2 |
| **Solo en Clientes** | Claves que existen en Fuente 2 pero no en Fuente 1 |
| **Resumen** | Métricas globales de la conciliación |

---

## ▶️ Cómo Usar

### Requisitos
```bash
pip install pandas openpyxl
```

### Modo Interactivo (recomendado — sin tocar el código)

```bash
python conciliar_inventarios.py
```

El script te guía paso a paso:

```
══════════════════════════════════════════════════════════
  CONCILIADOR DE INVENTARIOS — Modo Interactivo
══════════════════════════════════════════════════════════

─── FUENTE 1: Inventario General (almacén/físico) ───
  Ruta del archivo (.xlsx, .xlsm, .csv): C:\datos\inventario_general.xlsx
  Hojas disponibles: ['INV FISICO', 'Resumen']
  Nombre de la hoja [INV FISICO]: INV FISICO

─── FUENTE 2: Inventario Clientes (sistema/partidas) ───
  ¿Es el mismo archivo que la Fuente 1? (s/n) [n]: s
  Hojas disponibles: ['INV FISICO', 'INV CLIENTES', 'Resumen']
  Nombre de la hoja [INV CLIENTES]: INV CLIENTES

─── COLUMNA CLAVE DE COMPARACIÓN ───
  Nombre de la columna clave [UPC]: UPC

─── COLUMNA DE CANTIDAD ───
  Nombre de la columna de cantidad en Fuente 1 [TOTAL FISICO]: TOTAL FISICO
  Nombre de la columna de cantidad en Fuente 2 [TOTAL FISICO]: TOTAL FISICO
```

### Modo Configuración (para automatizar o ejecutar en batch)

Edita las variables al inicio del script:

```python
GEN_PATH   = r"C:\ruta\inventario_general.xlsx"
CLI_PATH   = r"C:\ruta\inventario_clientes.xlsx"
GEN_SHEET  = "INV FISICO"
CLI_SHEET  = "INV CLIENTES"
FORCED_KEY_COL       = "UPC"
FORCED_TOTAL_COL_GEN = "TOTAL FISICO"
FORCED_TOTAL_COL_CLI = "TOTAL FISICO"
OUTPUT_PATH = r"C:\ruta\salida\conciliacion_resultado.xlsx"
```

---

## 🔍 Detección Automática de Columnas

Si no especificas los nombres exactos, el script busca columnas por sinónimos comunes:

| Tipo | Nombres que reconoce automáticamente |
|---|---|
| Columna clave | `UPC`, `EAN`, `SKU`, `Codigo`, `Código`, `Clave`, `Item` |
| Columna cantidad | `TOTAL FISICO`, `Total`, `Cantidad`, `QTY`, `Existencia`, `Piezas`, `Unidades` |

Acepta variaciones de acentos y mayúsculas/minúsculas.

---

## 📊 Lógica de Conciliación

```
DIFERENCIA = QTY GENERAL − QTY CLIENTES

✅ OK       → DIFERENCIA = 0  (inventarios coinciden)
⚠️ FALTANTE → DIFERENCIA > 0  (el general tiene más que el cliente)
🔴 SOBRANTE → DIFERENCIA < 0  (el cliente tiene más que el general)
```

---

## 👤 Autor

**Ángel** — Account Manager & Data Analytics Specialist  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Conectar-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/angel-jimenez-28696b255)
[![GitHub](https://img.shields.io/badge/GitHub-Angel--jl-181717?style=flat-square&logo=github&logoColor=white)](https://github.com/Angel-jl)

---

## 📄 Licencia

MIT License — Libre para uso y adaptación con atribución.
