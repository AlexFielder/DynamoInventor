# DynamoInventor vase sample

Adapted from the supplied `DynamoSampleWorkflow-vase.dyn`, using the node serialization and
assembly-resolution conventions in the supplied `Inventor2027-Test.dyn`.

Target: AlexFielder/DynamoInventor, `master` commit
`c8f4276c880fa4e8466b5c40ba454aa4a940d338`, Inventor 2027 and Dynamo Core 4.2.1.5887.

## What is included

- `DynamoInventor-Vase.dyn`: original Dynamo geometry plus a native Inventor output node.
- `DynamoInventor-Vase-Preview.dyn`: the original geometry workflow prepared for Dynamo 4.2;
  works without the new C# node and creates no objects in Inventor.
- `src/Libraries/Inventor/DSInventorNodes/Features/InventorVase.cs` (relative to the package root):
  a new, hand-written Zero-Touch node. The existing SDK-style project automatically includes it.

The original vase graph has **no Revit nodes or Revit API dependency**. Its 15 nodes generate four
circles, move three of them vertically, collect them and use `Surface.ByLoft`. The conversion adds
an Inventor output because a Dynamo preview surface alone does not create an Inventor feature.
The reference test's saved Inventor bindings are deliberately not copied.

## Install the native output once

1. Close the DynamoInventor host window/process before rebuilding its DLLs.
2. Extract the package into your DynamoInventor repository root, retaining the `src` and `samples`
   directory structure. The new source goes into the existing `InventorLibrary` project; no wrapper
   regeneration, new NuGet package, Python engine or project-file edit is needed.
3. From that repository root, build/deploy using the project's normal development script:

   ```powershell
   pwsh -File .\scripts\deploy-dev-addin.ps1
   ```

4. Open a **standard part (.ipt)** in Inventor and make it the active document.
5. Launch DynamoInventor from the ribbon. Open `samples\Vase\DynamoInventor-Vase.dyn`.
6. Click **Run**. The graph intentionally opens in **Manual** mode so opening a file does not
   immediately change the active Inventor document. Use Zoom All in either application if needed.

Opening the native graph before rebuilding will leave `InventorVase.ByCircles` unresolved.
The preview-only graph can be opened immediately with your existing build.

## Expected result

The Dynamo background preview shows the original loft. The native output creates:

- Four offset work planes and four circular sketches in the active part.
- Eight user length parameters: a radius and Z height for each section. Their `DV_<id>_` prefix avoids
  collisions with pre-existing model parameters.
- One native Inventor **surface loft** named `DynamoVase`. Construction geometry is hidden.

The output code block is:

```text
InventorVase.ByCircles(sections, "DynamoVase");
```

`sections` is wired to the same `List.Create` output used by the original `Surface.ByLoft` node.
The C# method validates those circles, so the native output follows the actual graph section data.

This is an **open surface**, just like the original example: it has neither a closed bottom nor wall
thickness and is not a printable/watertight solid. The native feature is reconstructed through the
same sections using Inventor's loft algorithm. Exact surface equality between sections and
Dynamo's `Surface.ByLoft` has not been established.

## Units and default dimensions

One graph coordinate unit is one **centimetre**, matching the current DynamoInventor geometry
conversion. There is no feet-to-centimetres multiplier. The original slider values are preserved.
All slider labels state cm; their minimum is raised from 0 to 0.1 to avoid degenerate geometry.

| Input/section | Graph value | Millimetres |
|---|---:|---:|
| Bottom radius | 0.9 cm | 9 mm |
| Circle 2 radius | 3.1 cm | 31 mm |
| Circle 3 radius | 1.3 cm | 13 mm |
| Top radius | 1.4 cm | 14 mm |
| Total height | 15 cm | 150 mm |
| Section heights | 0, 4.5, 10.5, 15 cm | 0, 45, 105, 150 mm |

The intermediate heights remain 30% and 70% of total height. Inventor's displayed document units
can be millimetres or inches; the node supplies database length values in centimetres explicitly.

## Updates and ownership

Change a slider and click Run again. The node finds the loft tagged with the instance key
`DynamoVase` in the **currently active part**, updates its eight parameters and regenerates the
document. It does not delete/recreate the loft on each evaluation. Ownership attributes are stored
in the part, so saving/reopening the part does not depend on Dynamo trace being retained.

The key is scoped to the active part, not the graph/node GUID. Two graphs using the same instance
name intentionally control the same vase in that part. To create a second independent vase, change
the string to a different key. Changing the name creates another vase and leaves the earlier one.
Deleting a Dynamo node does not delete the Inventor model. The node does not save documents.

The sample accepts exactly four circles parallel to XY, centred on the Z axis, at increasing heights.
It reports an error for parts of the input outside that scope. It requires an active part and will
not create a new part or insert a component into an assembly automatically.

Creation/updates are enclosed in one Inventor transaction. A failed regeneration attempts to roll
back the edit. Missing/renamed generated parameters produce an error rather than causing the node
to silently rebuild or replace unrelated geometry. Manually deleting the loft can leave its support
geometry/parameters; clean up those objects before deliberately recreating it.

## Validation and remaining runtime check

Validated here: JSON parsing, node/port IDs, connector endpoints, unchanged original geometry
dependencies, section order/default dimensions and altered input scenarios. The Inventor calls
were checked against Autodesk's API documentation. See `VALIDATION.md` for the exact checks.

**The C# node has not been compiled or executed against Inventor here.** This environment has no
Inventor installation or .NET SDK. The native output is a source-level implementation for your
Windows build; the checks above are not a substitute for that build and smoke test.

Recommended Windows smoke test:

1. In an otherwise empty part, Run the native graph. Expect one loft, four additional planes, four
   sketches and eight user parameters; height 150 mm and largest input section diameter 62 mm.
2. Set height to 18 and Run. Expect section heights 0, 54, 126, 180 mm and unchanged object counts.
3. Change Circle 2 Radius to 4 and Run. Expect that section's diameter to be 80 mm, with the same loft.
4. Save the part, close/reopen it, reopen the graph and Run. Confirm the same tagged loft is updated.
5. Activate an assembly or drawing and Run. Expect the explicit 'open a part' error, without changes.

## API references

- [LoftFeatures.CreateLoftDefinition](https://help.autodesk.com/cloudhelp/2024/ENU/Inventor-API/files/LoftFeatures_CreateLoftDefinition.htm)
- [Profiles.AddForSurface](https://help.autodesk.com/cloudhelp/2024/ENU/Inventor-API/files/Profiles_AddForSurface.htm)
- [WorkPlanes.AddByPlaneAndOffset](https://help.autodesk.com/cloudhelp/2024/ENU/Inventor-API/files/WorkPlanes_AddByPlaneAndOffset.htm)
- [UserParameters.AddByValue](https://help.autodesk.com/cloudhelp/2024/ENU/Inventor-API/files/UserParameters_AddByValue.htm)
- [DimensionConstraints.AddRadius](https://help.autodesk.com/cloudhelp/2024/ENU/Inventor-API/files/DimensionConstraints_AddRadius.htm)

The accessible API reference pages above are the 2024 documentation for these established APIs;
the target project is the user's Inventor 2027 / Dynamo 4.2 integration.
