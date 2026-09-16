DynamoInventor
==============

Dynamo (visual programming) for Autodesk Inventor: an Inventor node library plus a host that runs
Dynamo Core against Inventor.

Current target: **Inventor 2027 + Dynamo Core 4.2.x on .NET 10**. The pairing is dictated by the CLR
each side runs on (Dynamo 4.x = .NET 10 = Inventor 2027; Dynamo 3.x = .NET 8 = Inventor 2025/2026).
Change `InventorVersion` and `DynamoVersion` in `src/Directory.Build.props` together.

Architecture
------------

Dynamo runs **out of process**. This is deliberate, not a stopgap: Inventor loads WebView2 SDK
1.0.1210 into its default AssemblyLoadContext at startup, Dynamo 4.2 is built against 1.0.2478, and
Dynamo's own `Assembly.LoadFrom` calls pin its extensions to that same context, so an in-process
Dynamo window dies in Dynamo's crash reporter (verified 2026-09-16). A separate process gets exactly
the dependencies Dynamo shipped with.

| Project | Output | Role |
|---|---|---|
| `DynamoInventor` | `DynamoInventor.dll` + `.addin` | The Inventor add-in. One ribbon button that starts or focuses the host. **No Dynamo dependency** (Inventor's loader reflects over every type in the add-in before `Activate`, so it must not reference Dynamo). |
| `DynamoInventor.App` | `DynamoInventor.App.exe` | WPF host that runs Dynamo the way Dynamo Sandbox does: resolves the Dynamo runtime, preloads Inventor's own ASM from *Common Files\Autodesk Shared\Components*, preloads `InventorLibrary.dll`, shows `DynamoView`. Single instance. |
| `DynamoInventor.Host` | `DynamoInventor.Host.dll` | Runtime locator/resolver and the `DynamoModel` subclass + `IPathResolver`. |
| `InventorLibrary` | `InventorLibrary.dll` | ZeroTouch node library over the Inventor API. Reaches Inventor over COM (`Inventor.Application` in the Running Object Table); starts Inventor if none is running. |
| `InventorServices` | `InventorServices.dll` | Reference-key binding, trace ids (base64/JSON in Dynamo's trace slot), IoC. |

All projects build into one flat `bin\<Configuration>\` folder so the add-in finds the host beside it.

Dynamo assemblies are compile-time NuGet references only (`DynamoVisualProgramming.*`,
`ExcludeAssets=runtime`). At runtime the host finds Dynamo Core in this order: `%DYNAMO_INVENTOR_RUNTIME%`,
`dynamo-runtime.txt` beside the exe, `DynamoCore\` beside the exe, newest `%ProgramFiles%\Dynamo\Dynamo Core\4.*`.

Developer setup
---------------

1. Inventor 2027 installed (interop is referenced from `Bin\Public Assemblies`).
2. A Dynamo Core 4.2.x runtime, e.g. `DynamoCoreRuntime4.2.1.5887.zip` from the DynamoDS/Dynamo GitHub
   release, extracted to `C:\AFAutomations\Tools\DynamoCoreRuntime\4.2.1` (or anywhere; pass `-DynamoRuntime`).
3. Build and register the add-in for the current user:

       pwsh -File scripts\deploy-dev-addin.ps1

   Allow the unsigned add-in once in Inventor's Add-In Manager. `-Remove` unregisters.

Smoke test without touching the UI (Inventor running or not; it will be started if needed):

    bin\Debug\DynamoInventor.App.exe --smoke-test "InventorWorkPoint.ByPoint(Point.ByCoordinates(1,2,3));"

Add `--then "<code>"` to rewrite the code block after the first run and exercise the update-in-place
path, or `--open <file.dyn>` to open a saved graph and log its first run.

Nodes
-----

Hand-written nodes work in the active **part or assembly** (a new assembly is created when nothing is
open; drawings and presentations are rejected with a clear error). A Dynamo unit is an Inventor
centimetre. Each geometry node is trace-bound: re-running the same node moves the object it made
rather than creating another.

| Node | What it does |
|---|---|
| `InventorWorkPoint.ByPoint(point)` | Fixed work point (assemblies only allow fixed work points; parts allow all kinds). |
| `InventorWorkPlane.ByPlane(plane)` / `.ByOriginXAxisYAxis(origin, x, y)` | Fixed work plane. |
| `InventorParameter.Names()` / `.UserParameterNames()` | Parameter names of the active document. |
| `InventorParameter.Value(name)` / `.Expression(name)` / `.Units(name)` | Read a parameter (value in its own units). |
| `InventorParameter.SetValue(name, value)` / `.SetExpression(name, expr)` | Drive a parameter and update the document. |
| `InventorParameter.AddUserParameter(name, value, units)` | Add (or set) a user parameter. |

The `InventorLibrary.API.Inv*` classes are generated wrappers over the interop (hidden from the
library, usable from code blocks). Everything is reachable over COM from the out-of-process host.

Wrapper generator
-----------------

`src\DynamoInventor.Generator` replaces the old IronPython 2.7 NodeGeneration project. It reflects
over the installed `Autodesk.Inventor.Interop.dll` and rewrites every `Inv*.cs` in
`src\Libraries\Inventor\DSInventorNodes\API` from the type list in `API\wrappers.json`:

    dotnet run --project src\DynamoInventor.Generator -- --config src\Libraries\Inventor\DSInventorNodes\API\wrappers.json

Members whose types are not in the list are emitted as `// skipped:` comments rather than as raw
interop types, so the output always compiles and never leaks `Inventor.*` names into DesignScript;
widen the list to expose more. Hand additions belong in `Inv<Name>.Custom.cs` partials, never in
generated files.

Logs: `%LOCALAPPDATA%\DynamoInventor\DynamoInventor.log` (add-in) and `DynamoInventor.App.log` (host);
Dynamo's own log under `%APPDATA%\Dynamo\Dynamo Inventor\4.2\Logs`.

Known issues
------------

- Naming: hand-written node classes must not reuse a generated wrapper's `Inv*` name or an Inventor
  interop type name (many wrapper members expose interop types publicly, so DesignScript can see
  both). `InvWorkPoint` (node) was renamed `InventorWorkPoint` for this reason.
- Three files in `InventorLibrary` were never part of the build and stay excluded until reviewed
  (`API\InvOGSSceneNode.cs`, `InvWorkPoint.cs`, `ModulePlacement\DSAssemblyComponent.cs`).
- Units: conversions are raw centimetres (the Inventor API's internal unit); Dynamo 4 has no host-unit setting.
- The ModulePlacement nodes predate the part/assembly rework and still assume an assembly.
