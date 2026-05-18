# AcSapfir

AutoCAD 2020 plugin for exporting geometry to Sapfir (САПФИР).

## Requirements

- AutoCAD 2020
- Sapfir (САПФИР 2018+)
- .NET Framework 4.8

## Installation

1. Build `AcSapfir.dll` (Debug or Release)
2. In AutoCAD, run `NETLOAD` and select `bin\Debug\AcSapfir.dll`
3. Use the commands listed below

## Commands

| Command | Description |
|---|---|
| `Sapfir_AXES` | Create coordinate axes from lines (numeric names) |
| `Sapfir_AXES_WIZ` | Axes wizard: auto-group parallel lines, 4 naming rules |
| `Sapfir_Point_LOADS` | Create point loads from points |
| `Sapfir_Line_LOADS` | Create uniform line loads from lines |
| `Sapfir_Line_LOADS2` | Create trapezoidal line loads from lines |
| `Sapfir_SLABS` | Create slabs from polylines |
| `Sapfir_SLAB_Holes` | Create slab holes from polylines |
| `Sapfir_Found_SLABS` | Create foundation slabs from polylines |
| `Sapfir_WALLS` | Create walls from lines |
| `Sapfir_WALL_EDIT` | Edit existing walls (update geometry from XData) |
| `Sapfir_DOORS` | Create doors in walls from lines |
| `Sapfir_WINDOWS` | Create windows in walls from lines |
| `Sapfir_COLUMNS` | Create columns from polylines (auto/manual angle) |
| `Sapfir_BEAMS` | Create beams from lines with h/b section |
| `Sapfir_LINES` | Create building lines from AutoCAD lines |
| `Sapfir_STOREYS` | Create storeys from MText labels |
| `Sapfir_STOREYS_WIZ` | Storey generator wizard (modal, with templates) |

## Features

- **Multi-tab support** — commands work with the active Sapfir tab via `GetActiveSapfirView()`
- **XData tracking** — created Sapfir IDs are stored back to AutoCAD entities for later editing
- **Material selector** — modal dialog to pick material from project before creating elements
- **Console logging** — all commands output created IDs to AutoCAD console (F2)

## Coordinate Scaling

All geometry is scaled from millimeters (AutoCAD) to meters (Sapfir) with a factor of 0.001.

Storey elevation uses absolute mark in meters: `newStorey.Parameter["M_LEVEL"] = elevation_m`.

## Build

```bash
msbuild AcSapfir.csproj /t:Build /p:Configuration=Debug
```

Output: `bin\Debug\AcSapfir.dll`
