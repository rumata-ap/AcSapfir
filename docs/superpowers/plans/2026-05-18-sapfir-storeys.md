# Sapfir_STOREYS Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add AutoCAD command `Sapfir_STOREYS` that creates a set of building storeys in Sapfir from selected MText objects, using Y-coordinate as elevation and text content as storey name.

**Architecture:** New method `GetMTextsData` in `AcUtilites.cs` extracts (name, Y) pairs from MText. New method `CreateStoreys` and command `Sapfir_STOREYS` in `SapfirDrafter.cs` clear existing storeys, sort by elevation, and create new storeys with `SetPosition(0, 0, Z)` where Z = Y * 0.001.

**Tech Stack:** C# .NET Framework 4.7, AutoCAD 2020 API, SapfirLib COM

---

### Task 1: Add `GetMTextsData` method to `AcUtilites.cs`

**Files:**
- Modify: `AcUtilites.cs`

- [ ] **Step 1: Add using directive for MText**

Add `using Autodesk.AutoCAD.DatabaseServices;` if not already present (it is already imported). The `MText` class is in `Autodesk.AutoCAD.DatabaseServices`.

- [ ] **Step 2: Add `GetMTextsData` method**

Add the following method to the `AcUtilites` class, after the existing `ActionOnPolylines` methods (around line 396):

```csharp
public static List<Tuple<string, double>> GetMTextsData(ObjectId[] objIds)
{
    List<Tuple<string, double>> result = new List<Tuple<string, double>>();
    Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

    if (objIds != null)
    {
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            try
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    if (acEnt is MText mtext)
                    {
                        string text = mtext.PlainText();
                        double y = mtext.Location.Y;
                        result.Add(new Tuple<string, double>(text, y));
                    }
                }
                acTrans.Commit();
            }
            finally
            {
                acTrans.Dispose();
            }
        }
    }

    return result;
}
```

Key points:
- Uses `MText.PlainText()` to get clean text without formatting codes
- Uses `MText.Location.Y` for the Y insertion point coordinate (in mm)
- Opens entities `ForRead` since we only extract data, not modify
- Returns `List<Tuple<string, double>>` — pairs of (name, elevation)

- [ ] **Step 3: Build to verify compilation**

Run: `msbuild AcSapfir.csproj /p:Configuration=Debug /t:Build`
Expected: Build succeeds with no errors. If `MText` type is not recognized, check that `acdbmgd.dll` reference is present.

- [ ] **Step 4: Commit**

```bash
git add AcUtilites.cs
git commit -m "feat: add GetMTextsData method to extract MText data for storeys"
```

---

### Task 2: Add `CreateStoreys` method and `Sapfir_STOREYS` command to `SapfirDrafter.cs`

**Files:**
- Modify: `SapfirDrafter.cs`

- [ ] **Step 1: Add `using` for `System.Linq`** (already present)

Verify `using System.Linq;` is in the using block at the top of `SapfirDrafter.cs`. It is already there.

- [ ] **Step 2: Add `CreateStoreys` method**

Add this method to the `SapfirDrafter` class, before the `[CommandMethod("Sapfir_AXES")]` method (around line 300):

```csharp
void CreateStoreys(List<Tuple<string, double>> storeyData)
{
    while (projSpf.CountStorey > 0)
    {
        projSpf.DelStoreyByIndex(0);
    }

    var sorted = storeyData.OrderBy(s => s.Item2).ToList();

    foreach (var item in sorted)
    {
        string name = item.Item1;
        double y = item.Item2;
        AutoStorey newStorey = projSpf.NewStorey(name);
        newStorey.SetPosition(0, 0, y * 0.001);
    }
}
```

Key points:
- Deletes all existing storeys first by repeatedly removing index 0
- Sorts by Y ascending (lowest elevation first)
- Creates each storey with `NewStorey(name)` and `SetPosition(0, 0, Z)` where Z = Y_mm * 0.001 (conversion to meters)
- Does NOT update `storeySpf` — user selects active storey manually in Sapfir

- [ ] **Step 3: Add `Sapfir_STOREYS` command method**

Add this command method to the `SapfirDrafter` class, after the existing `Sapfir_test` stub (around line 427, before the closing brace of the class — but note there are duplicate `Sapfir_doors` and `Sapfir_test` methods, add before the first duplicate):

```csharp
[CommandMethod("Sapfir_STOREYS", CommandFlags.UsePickSet)]
public void Sapfir_STOREYS()
{
    ObjectId[] objIds = AcUtilites.Selection();
    List<Tuple<string, double>> data = AcUtilites.GetMTextsData(objIds);
    if (data.Count > 0)
    {
        CreateStoreys(data);
    }
}
```

Place this after the first `Sapfir_test()` (line ~415) and before the duplicate `Sapfir_doors()` (line ~417).

- [ ] **Step 4: Build to verify compilation**

Run: `msbuild AcSapfir.csproj /p:Configuration=Debug /t:Build`
Expected: Build succeeds with no errors.

- [ ] **Step 5: Commit**

```bash
git add SapfirDrafter.cs
git commit -m "feat: add Sapfir_STOREYS command and CreateStoreys method"
```

---

### Task 3: Remove duplicate method declarations in `SapfirDrafter.cs`

**Files:**
- Modify: `SapfirDrafter.cs`

Looking at the source, there are duplicate declarations:
- `Sapfir_doors()` appears twice (lines ~402 and ~418)
- `Sapfir_test()` appears twice (lines ~413 and ~425)

- [ ] **Step 1: Identify and remove duplicates**

Remove the second (duplicate) block at lines 417-429, which contains:
```csharp
[CommandMethod("Sapfir_DOORS", CommandFlags.UsePickSet)]
public void Sapfir_doors()
{
    Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
    PromptDoubleResult result = acDocEd.GetDouble("Введите высоту проема в метрах: ");
    if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateDoors, GetWalls(out int cnt), result.Value);     
}

[CommandMethod("Sapfir_test", CommandFlags.UsePickSet)]
public void Sapfir_test()
{
    
}
```

Keep only the first occurrences (lines ~402-415).

- [ ] **Step 2: Build to verify**

Run: `msbuild AcSapfir.csproj /p:Configuration=Debug /t:Build`
Expected: Build succeeds. The duplicate `CommandMethod` attributes would cause runtime issues with AutoCAD.

- [ ] **Step 3: Commit**

```bash
git add SapfirDrafter.cs
git commit -m "fix: remove duplicate Sapfir_doors and Sapfir_test method declarations"
```

---

### Task 4: Final verification

- [ ] **Step 1: Full rebuild and manual review**

Run: `msbuild AcSapfir.csproj /p:Configuration=Debug /t:Rebuild`
Expected: Build succeeds.

- [ ] **Step 2: Review key design decisions against spec**

Verify in the built code:
1. `GetMTextsData` uses `MText.PlainText()` (not `.Text`)
2. `CreateStoreys` uses `SetPosition(0, 0, y * 0.001)` — Z coordinate, mm→m
3. `CreateStoreys` sorts by Y ascending
4. `CreateStoreys` deletes all existing storeys before creating new ones
5. `Sapfir_STOREYS` command uses `CommandFlags.UsePickSet`
6. No duplicate method declarations remain

- [ ] **Step 3: Final commit if any fixes needed**

```bash
git add -A
git commit -m "chore: final verification and cleanup for Sapfir_STOREYS feature"
```