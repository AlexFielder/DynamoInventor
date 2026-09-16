using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Loader;
using Inventor;
using IOPath = System.IO.Path;

namespace AsmSpike
{
    /// <summary>
    /// Loads on Inventor startup, preloads Dynamo 4.2.1's libG/ASM 232, starts a headless
    /// DynamoModel, evaluates some ProtoGeometry, then exercises Inventor's own kernel
    /// (TransientBRep) to prove the two ASMs coexist. Everything is written to the log file;
    /// nothing is shown in the UI.
    /// </summary>
    [ComVisible(true)]
    [Guid("6B8C2E4A-9D31-4F6E-A0C7-5E2D1B9F3A11")]
    public class AsmSpikeAddIn : ApplicationAddInServer
    {
        // Same defaults as AsmSpike.csproj; overridable via environment for a second run.
        private static readonly string DynamoRuntime =
            System.Environment.GetEnvironmentVariable("ASMSPIKE_DYNAMO_RUNTIME") ?? @"C:\AFAutomations\Tools\DynamoCoreRuntime\4.2.1";
        private static readonly string AsmDir =
            System.Environment.GetEnvironmentVariable("ASMSPIKE_ASM_DIR") ?? @"C:\Program Files\Autodesk\CFD 2027";
        private static readonly string LogPath =
            System.Environment.GetEnvironmentVariable("ASMSPIKE_LOG") ?? @"C:\AFAutomations\Tools\DynamoCoreRuntime\AsmSpike.log";

        private Inventor.Application app;
        private object dynamoModel; // kept as object so this class has no static dependency on DynamoCore

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            app = addInSiteObject.Application;
            try { System.IO.File.Delete(LogPath); } catch { }
            Log("=== AsmSpike start " + DateTime.Now.ToString("s"));
            Log("Inventor " + app.SoftwareVersion.DisplayVersion + " (" + app.SoftwareVersion.DisplayVersion + ")  build " + app.SoftwareVersion.BuildIdentifier);
            Log("CLR " + System.Environment.Version + "  process x64=" + System.Environment.Is64BitProcess);
            var ownAlc = AssemblyLoadContext.GetLoadContext(typeof(AsmSpikeAddIn).Assembly);
            Log("Add-in ALC: '" + (ownAlc?.Name ?? "<null>") + "'  isDefault=" + (ownAlc == AssemblyLoadContext.Default));
            Log("DynamoRuntime=" + DynamoRuntime);
            Log("AsmDir=" + AsmDir);

            // Resolve Dynamo assemblies from the runtime folder in BOTH contexts: the one Inventor
            // gave us and the default one (Assembly.LoadFrom inside Dynamo always lands in Default).
            AssemblyLoadContext.Default.Resolving += ResolveFromDynamoRuntime;
            if (ownAlc != null && ownAlc != AssemblyLoadContext.Default)
                ownAlc.Resolving += ResolveFromDynamoRuntime;

            try
            {
                RunSpike();
                Log("SPIKE OK");
            }
            catch (Exception ex)
            {
                Log("SPIKE FAILED: " + ex);
                var inner = ex.InnerException;
                while (inner != null) { Log("  inner: " + inner); inner = inner.InnerException; }
            }
            finally
            {
                LogAsmModules("after");
                Log("=== AsmSpike end " + DateTime.Now.ToString("s"));
            }
        }

        private static Assembly ResolveFromDynamoRuntime(AssemblyLoadContext ctx, AssemblyName name)
        {
            var candidate = IOPath.Combine(DynamoRuntime, name.Name + ".dll");
            if (!System.IO.File.Exists(candidate)) return null;
            try
            {
                var asm = ctx.LoadFromAssemblyPath(candidate);
                Log("  resolved " + name.Name + " -> " + candidate + " (ctx '" + ctx.Name + "')");
                return asm;
            }
            catch (Exception ex)
            {
                Log("  resolve FAILED " + name.Name + ": " + ex.Message);
                return null;
            }
        }

        // Separate method so the JIT does not touch Dynamo types before the Resolving handlers exist.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private void RunSpike()
        {
            LogAsmModules("before");

            // 1. Which ASM build are we about to load?
            var asmVersion = DynamoShapeManager.Utilities.GetVersionFromPath(AsmDir);
            Log("ASM at AsmDir reports version " + asmVersion);

            // 2. Start a headless Dynamo model exactly the way DynamoCLI does; this preloads libG + ASM.
            var userData = IOPath.Combine(IOPath.GetDirectoryName(LogPath), "spike-data", "user");
            var commonData = IOPath.Combine(IOPath.GetDirectoryName(LogPath), "spike-data", "common");
            Directory.CreateDirectory(userData);
            Directory.CreateDirectory(commonData);

            var sw = Stopwatch.StartNew();
            var model = Dynamo.Applications.StartupUtils.MakeCLIModel(
                AsmDir, userData, commonData,
                new Dynamo.Models.HostAnalyticsInfo { HostName = "Dynamo Inventor", HostProductName = "Inventor" });
            dynamoModel = model;
            Log("DynamoModel started in " + sw.ElapsedMilliseconds + " ms; Dynamo " + Dynamo.Models.DynamoModel.Version);
            LogAsmModules("after Dynamo start");

            // 3. Dynamo-side geometry through libG/ASM 232.
            var p = Autodesk.DesignScript.Geometry.Point.ByCoordinates(1, 2, 3);
            var d = p.DistanceTo(Autodesk.DesignScript.Geometry.Point.Origin());
            Log("Dynamo Point(1,2,3).DistanceTo(origin) = " + d + "   (expect 3.7417)");

            var box = Autodesk.DesignScript.Geometry.Cuboid.ByLengths(10, 20, 30);
            Log("Dynamo Cuboid(10,20,30).Volume = " + box.Volume + "   (expect 6000)");

            var sphere = Autodesk.DesignScript.Geometry.Sphere.ByCenterPointRadius(
                Autodesk.DesignScript.Geometry.Point.ByCoordinates(5, 10, 15), 8);
            var union = box.Union(sphere);
            Log("Dynamo Cuboid.Union(Sphere).Volume = " + union.Volume + "   (expect > 6000, boolean through ASM 232)");

            // 4. Inventor-side geometry through Inventor's own kernel, after Dynamo's ASM is resident.
            var tg = app.TransientGeometry;
            var tb = app.TransientBRep;
            var bx = tg.CreateBox();
            bx.MinPoint = tg.CreatePoint(0, 0, 0);
            bx.MaxPoint = tg.CreatePoint(1, 2, 3);
            var invBox = tb.CreateSolidBlock(bx);
            Log("Inventor TransientBRep block volume = " + invBox.Volume[0.0001] + " cm^3   (expect 6)");
            var invSphere = tb.CreateSolidSphere(tg.CreatePoint(0.5, 1, 1.5), 0.8);
            tb.DoBoolean(invBox, invSphere, BooleanTypeEnum.kBooleanTypeUnion);
            Log("Inventor block UNION sphere volume = " + invBox.Volume[0.0001] + " cm^3   (expect > 6, boolean through Inventor's ASM)");

            // 5. Back to Dynamo once more, to make sure Inventor's kernel work did not upset libG.
            var p2 = Autodesk.DesignScript.Geometry.Point.ByCoordinates(3, 4, 0);
            Log("Dynamo Point(3,4,0).DistanceTo(origin) = " + p2.DistanceTo(Autodesk.DesignScript.Geometry.Point.Origin()) + "   (expect 5)");
        }

        private static void LogAsmModules(string phase)
        {
            try
            {
                var mods = Process.GetCurrentProcess().Modules.Cast<ProcessModule>()
                    .Where(m => {
                        var n = m.ModuleName.ToUpperInvariant();
                        return n.StartsWith("ASM") || n.StartsWith("RXASM") || n.StartsWith("LIBG") || n.StartsWith("TBB");
                    })
                    .Select(m => m.ModuleName + "  <-  " + m.FileName)
                    .OrderBy(s => s).ToList();
                Log("Kernel modules (" + phase + "): " + mods.Count);
                foreach (var m in mods) Log("    " + m);
            }
            catch (Exception ex) { Log("module scan failed: " + ex.Message); }
        }

        private static readonly object logLock = new object();
        private static void Log(string line)
        {
            lock (logLock)
            {
                System.IO.File.AppendAllText(LogPath, DateTime.Now.ToString("HH:mm:ss.fff") + "  " + line + System.Environment.NewLine);
            }
        }

        public void Deactivate()
        {
            try { (dynamoModel as IDisposable)?.Dispose(); } catch { }
            dynamoModel = null;
            app = null;
        }

        public void ExecuteCommand(int commandID) { }
        public object Automation => null;
    }
}
