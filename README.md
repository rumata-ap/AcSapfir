# AcSapfir

AutoCAD 2020 plugin for exporting geometry to Sapfir (САПФИР).

## Requirements

- AutoCAD 2020
- Sapfir (САПФИР 2018+)
- .NET Framework 4.8

## Installation

1. Build `AcSapfir.dll` (Debug or Release)
2. In AutoCAD, run `NETLOAD` and select the DLL
3. Use the commands listed below

## Commands

| Command | Description |
|---|---|
| `Sapfir_AXES` | Create coordinate axes from lines |
| `Sapfir_Point_LOADS` | Create point loads from points |
| `Sapfir_Line_LOADS` | Create uniform line loads from lines |
| `Sapfir_Line_LOADS2` | Create trapezoidal line loads from lines |
| `Sapfir_SLABS` | Create slabs from polylines |
| `Sapfir_SLAB_Holes` | Create slab holes from polylines |
| `Sapfir_Found_SLABS` | Create foundation slabs from polylines |
| `Sapfir_WALLS` | Create walls from lines |
| `Sapfir_DOORS` | Create doors in walls from lines |
| `Sapfir_WINDOWS` | Create windows in walls from lines |
| `Sapfir_COLUMNS` | Create columns from polylines | 
| `Sapfir_STOREYS` | Create storeys from MText labels |

## Coordinate Scaling

All geometry is scaled from millimeters (AutoCAD) to meters (Sapfir) with a factor of 0.001.

Storey elevation uses Z-axis: `SetPosition(0, 0, Y*0.001)` where Y is the absolute elevation in mm.

## Build

```bash
msbuild AcSapfir.csproj /t:Build /p:Configuration=Debug
```

Output: `bin\Debug\AcSapfir.dll`
