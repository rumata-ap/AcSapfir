# Sapfir Columns & Windows Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `Sapfir_COLUMNS` (колонны из полилиний AutoCAD с повторным использованием сечений) and `Sapfir_WINDOWS` (окна, аналог дверей) commands with per-call active-tab resolution.

**Architecture:** Two new `[CommandMethod]` methods in `SapfirDrafter.cs`. `Sapfir_COLUMNS` uses `ResolveActiveContext()`, extracts vertices from polylines, computes centroids, normalizes and compares multi-contours for section reuse, creates `TM_COLUMN` models with storey-level height. `Sapfir_WINDOWS` mirrors `Sapfir_DOORS` with `TM_WINDOW` type. A shared helper `NormalizeContour` extracts centroid and normalized side lengths for geometric comparison. A `GetColumnSections` helper collects existing column multi-contours on the active storey for reuse matching.

**Tech Stack:** C# .NET Framework 4.8, AutoCAD 2020 API, SapfirLib COM interop

---

## File Structure

| File | Action | Responsibility |
|---|---|---|
| `SapfirDrafter.cs` | Modify | Two new commands, section reuse helpers, contour normalization |

---

### Task 1: Column section geometry helper — normalize contour and compare

**Files:**
- Modify: `SapfirDrafter.cs` — add helpers inside class before `ResolveActiveContext`

- [ ] **Step 1: Add `NormalizeContour` helper**

Add the following method inside `SapfirDrafter` class, before `ResolveActiveContext()`:

```csharp
double[] NormalizeContour(object[] vertices)
{
    int n = vertices.Length / 3;
    if (n < 3) return new double[n];

    double cx = 0, cy = 0;
    for (int i = 0; i < n; i++)
    {
        cx += (double)vertices[3 * i];
        cy += (double)vertices[3 * i + 1];
    }
    cx /= n;
    cy /= n;

    double[] sides = new double[n];
    double maxSide = 0;
    for (int i = 0; i < n; i++)
    {
        int j = (i + 1) % n;
        double dx = (double)vertices[3 * j] - (double)vertices[3 * i];
        double dy = (double)vertices[3 * j + 1] - (double)vertices[3 * i + 1];
        sides[i] = Math.Sqrt(dx * dx + dy * dy);
        if (sides[i] > maxSide) maxSide = sides[i];
    }

    if (maxSide > 1e-12)
        for (int i = 0; i < n; i++) sides[i] /= maxSide;

    return sides;
}
```

- [ ] **Step 2: Add `ContoursMatch` helper**

```csharp
bool ContoursMatch(double[] a, double[] b)
{
    if (a.Length != b.Length) return false;

    double epsilon = 0.02;

    for (int shift = 0; shift < a.Length; shift++)
    {
        bool match = true;
        for (int i = 0; i < a.Length; i++)
        {
            int bi = (i + shift) % a.Length;
            if (Math.Abs(a[i] - b[bi]) > epsilon) { match = false; break; }
        }
        if (match) return true;
    }

    return false;
}
```

- [ ] **Step 3: Verify build compiles**

```bash
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" AcSapfir.csproj /t:Build /p:Configuration=Debug
```
Expected: 0 errors.


### Task 2: Get existing column sections and find match helper

**Files:**
- Modify: `SapfirDrafter.cs` — add helpers after `ContoursMatch`

- [ ] **Step 1: Add `GetColumnSections` method**

```csharp
Dictionary<int, double[]> GetColumnSections()
{
    var sections = new Dictionary<int, double[]>();
    for (int i = 0; i < storeySpf.CountModel; i++)
    {
        var model = storeySpf.GetModelByIndex(i);
        if (model.TypeModel != (int)ModelsTypes.TM_COLUMN) continue;

        try
        {
            var mc = model.GetMultiContour();
            if (mc == null) continue;

            var contour = mc.GetContour();
            if (contour == null || contour.CountLine == 0) continue;

            var pl = contour.GetLine(0);
            if (pl == null) continue;

            object ptsBuf = new object[0];
            pl.GetPoints(ref ptsBuf);
            object[] pts = (object[])ptsBuf;
            if (pts != null && pts.Length >= 9)
            {
                sections[model.ID] = NormalizeContour(pts);
            }
        }
        catch { }
    }
    return sections;
}
```

- [ ] **Step 2: Add `FindMatchingSection` method**

```csharp
int FindMatchingSection(object[] targetVertices, Dictionary<int, double[]> sections)
{
    double[] targetNorm = NormalizeContour(targetVertices);
    foreach (var kvp in sections)
    {
        if (ContoursMatch(targetNorm, kvp.Value))
            return kvp.Key;
    }
    return 0;
}
```

- [ ] **Step 3: Build and verify**

```bash
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" AcSapfir.csproj /t:Build /p:Configuration=Debug
```
Expected: 0 errors.


### Task 3: Add `Sapfir_COLUMNS` command

**Files:**
- Modify: `SapfirDrafter.cs` — add before `Sapfir_test` method

- [ ] **Step 1: Add the command method**

```csharp
[CommandMethod("Sapfir_COLUMNS", CommandFlags.UsePickSet)]
public void Sapfir_COLUMNS()
{
    ResolveActiveContext();

    ObjectId[] objIds = AcUtilites.Selection();
    if (objIds == null) return;

    double storeyHeight = 3.0;
    try
    {
        object levelObj = storeySpf.Parameter["M_LEVEL"];
        if (levelObj != null && double.TryParse(levelObj.ToString(),
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double h))
            storeyHeight = h;
    }
    catch { }

    var sections = GetColumnSections();

    Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
    {
        foreach (ObjectId acObjId in objIds)
        {
            Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
            if (!(acEnt is Polyline polyline)) continue;

            int n = polyline.NumberOfVertices;
            if (n < 2) continue;

            double cx = 0, cy = 0;
            object[] verts = new object[n * 3];
            for (int i = 0; i < n; i++)
            {
                Point2d pt = polyline.GetPoint2dAt(i);
                verts[3 * i] = pt.X * 0.001;
                verts[3 * i + 1] = pt.Y * 0.001;
                verts[3 * i + 2] = 0.0;
                cx += pt.X * 0.001;
                cy += pt.Y * 0.001;
            }
            cx /= n;
            cy /= n;

            int matchingId = FindMatchingSection(verts, sections);

            if (matchingId == 0)
            {
                AutoModel sectionCol = storeySpf.NewModel((int)ModelsTypes.TM_COLUMN);
                sectionCol.SetPosition(cx, cy, 0);
                var multiCont = sectionCol.GetMultiContour();
                if (multiCont != null)
                {
                    var cont = multiCont.GetContour();
                    if (cont != null)
                    {
                        var pl = cont.NewPolyLine();
                        pl.SetPoints(verts);
                    }
                }
                sectionCol.Parameter["M_HEIGHT"] = storeyHeight;
                sectionCol.RegenModel();
                matchingId = sectionCol.ID;
                sections[matchingId] = NormalizeContour(verts);
            }

            AutoModel column = storeySpf.NewModel((int)ModelsTypes.TM_COLUMN);
            column.SetPosition(cx, cy, 0);
            if (matchingId > 0)
            {
                try
                {
                    var srcModel = storeySpf.GetModelByID(matchingId);
                    var srcMc = srcModel.GetMultiContour();
                    if (srcMc != null)
                    {
                        var dstMc = column.GetMultiContour();
                        if (dstMc != null)
                        {
                            var srcCont = srcMc.GetContour();
                            if (srcCont != null && srcCont.CountLine > 0)
                            {
                                var srcPl = srcCont.GetLine(0);
                                object srcPts = new object[0];
                                srcPl.GetPoints(ref srcPts);
                                var dstCont = dstMc.GetContour();
                                if (dstCont != null)
                                {
                                    var dstPl = dstCont.NewPolyLine();
                                    dstPl.SetPoints((object[])srcPts);
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            column.Parameter["M_HEIGHT"] = storeyHeight;
            column.RegenModel();
        }
        acTrans.Commit();
    }
}
```

- [ ] **Step 2: Build**

```bash
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" AcSapfir.csproj /t:Build /p:Configuration=Debug
```
Expected: 0 errors.


### Task 4: Add `Sapfir_WINDOWS` command

**Files:**
- Modify: `SapfirDrafter.cs` — add before `Sapfir_test` method (after `Sapfir_COLUMNS`)

- [ ] **Step 1: Add `GetWallsForStorey` helper (reuses existing logic)**

Existing `GetWalls(out int count)` already works — no new method needed. The command will call it directly.

- [ ] **Step 2: Add the command (clone of `Sapfir_doors` with `TM_WINDOW` type)**

```csharp
[CommandMethod("Sapfir_WINDOWS", CommandFlags.UsePickSet)]
public void Sapfir_WINDOWS()
{
    ResolveActiveContext();
    Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
    PromptDoubleResult result = acDocEd.GetDouble("Введите высоту окна в метрах: ");

    if (result.Status != PromptStatus.OK) return;

    ObjectId[] objIds = AcUtilites.Selection();
    if (objIds == null) return;

    List<int> wallIds = GetWalls(out int cnt);
    if (cnt == 0) return;

    Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
    {
        foreach (ObjectId acObjId in objIds)
        {
            Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
            if (!(acEnt is Line line)) continue;

            foreach (int item in wallIds)
            {
                paramModelSpf = storeySpf.GetModelByID(item);
                polylineSpf = paramModelSpf.GetAxisLine();
                AutoLine spfLine = polylineSpf.GetLine(0);
                buf1 = new object[6];
                spfLine.GetPoints(ref buf1);
                object[] crd = (object[])buf1;

                Line2d wallLine = new Line2d(
                    new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000),
                    new Point2d((double)crd[3] * 1000, (double)crd[4] * 1000));
                Line2d winLine = new Line2d(
                    new Point2d(line.StartPoint.X, line.StartPoint.Y),
                    new Point2d(line.EndPoint.X, line.EndPoint.Y));
                Tolerance tol = new Tolerance(1E-8, 1E-8);

                if (wallLine.IsColinearTo(winLine, tol))
                {
                    AutoModel modelWindow = paramModelSpf.NewHole((int)ModelsTypes.TM_WINDOW);
                    modelWindow.Parameter["B"] = line.Length * 0.001;
                    modelWindow.Parameter["H"] = result.Value;
                    Point2d wallStart = new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000);
                    Point2d winStart = new Point2d(line.StartPoint.X, line.StartPoint.Y);
                    Point2d winEnd = new Point2d(line.EndPoint.X, line.EndPoint.Y);
                    modelWindow.Parameter["M_POS"] = Math.Min(
                        wallStart.GetDistanceTo(winStart),
                        wallStart.GetDistanceTo(winEnd)) * 0.001;
                }
            }
        }
        acTrans.Commit();
    }
}
```

- [ ] **Step 3: Build final**

```bash
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" AcSapfir.csproj /t:Build /p:Configuration=Debug
```
Expected: 0 errors.


### Task 5: Final verification

- [ ] **Step 1: Verify all commands are present**

```bash
Select-String -Path SapfirDrafter.cs -Pattern 'CommandMethod' | ForEach-Object { $_.Line.Trim() }
```
Expected output includes:
- `[CommandMethod("Sapfir_COLUMNS", CommandFlags.UsePickSet)]`
- `[CommandMethod("Sapfir_WINDOWS", CommandFlags.UsePickSet)]`

- [ ] **Step 2: Clean build one more time**

```bash
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" AcSapfir.csproj /t:Build /p:Configuration=Debug
```
Expected: 0 errors.

