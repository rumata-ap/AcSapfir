# AGENTS.md - AcSapfir

AutoCAD 2020 plugin for exporting geometry to Sapfir.

## Tech Stack
- **Language:** C# (.NET Framework 4.7)
- **Host:** AutoCAD 2020
- **Dependencies:**
  - AutoCAD API: `accoremgd.dll`, `acdbmgd.dll`, `acmgd.dll`
  - COM Libraries: `SapfirLib`, `MathServLib`

## Architecture
- `SapfirDrafter.cs`: Contains the main AutoCAD commands (`[CommandMethod]`) and logic for creating Sapfir models.
- `AcUtilites.cs`: Helper class for selecting and iterating over AutoCAD entities using transactions.

## Developer Notes
- **Build Requirements:** Requires AutoCAD 2020 installed. DLL references in `.csproj` use absolute paths to `C:\Program Files\Autodesk\AutoCAD 2020\`.
- **Commands:**
  - `Sapfir_AXES`: Create axes.
  - `Sapfir_Point_LOADS`: Create point loads.
  - `Sapfir_Line_LOADS` / `Sapfir_Line_LOADS2`: Create line loads.
  - `Sapfir_SLABS`: Create slabs.
  - `Sapfir_SLAB_Holes`: Create slab holes.
  - `Sapfir_Found_SLABS`: Create foundation slabs.
  - `Sapfir_WALLS`: Create walls.
  - `Sapfir_DOORS`: Create doors.
  - `Sapfir_STOREYS`: Create storeys from selected MText (Y = elevation in mm, text = name).
- **Coordinate Scaling:** Geometry is scaled by `0.001` when passed to `SapfirLib` (mm to m conversion). Storey elevation uses Z-axis in Sapfir: `SetPosition(0, 0, Y*0.001)`.
- **Reference Docs:**
  - AutoCAD .NET API: https://help.autodesk.com/view/OARX/2020/ENU/?guid=GUID-C3F3C736-40CF-44A0-9210-55F6A939B6F2
