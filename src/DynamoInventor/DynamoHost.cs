using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Loader;
using System.Windows.Interop;

namespace DynamoInventor
{
    /// <summary>
    /// Locates the Dynamo Core runtime on disk and makes its assemblies resolvable from Inventor's
    /// default AssemblyLoadContext. Nothing in this class may touch a Dynamo type directly: it runs
    /// before the resolver exists.
    /// </summary>
    internal static class DynamoRuntime
    {
        public const string RuntimeEnvironmentVariable = "DYNAMO_INVENTOR_RUNTIME";

        private static readonly object gate = new object();
        private static bool resolverRegistered;
        private static string rootFolder;

        /// <summary>
        /// Folder containing DynamoCore.dll, libg_xxx_0_0\ etc. Probe order:
        /// 1. %DYNAMO_INVENTOR_RUNTIME%
        /// 2. DynamoCore\ next to the add-in
        /// 3. the newest 4.x "Dynamo Core" install under %ProgramFiles%\Dynamo
        /// </summary>
        public static string RootFolder
        {
            get
            {
                if (rootFolder != null) return rootFolder;

                var candidates = new System.Collections.Generic.List<string>();
                var fromEnv = Environment.GetEnvironmentVariable(RuntimeEnvironmentVariable);
                if (!string.IsNullOrEmpty(fromEnv)) candidates.Add(fromEnv);

                var addinFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(addinFolder)) candidates.Add(Path.Combine(addinFolder, "DynamoCore"));

                var programFilesDynamo = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Dynamo", "Dynamo Core");
                if (Directory.Exists(programFilesDynamo))
                {
                    candidates.AddRange(Directory.GetDirectories(programFilesDynamo, "4.*").OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase));
                }

                rootFolder = candidates.FirstOrDefault(c => File.Exists(Path.Combine(c, "DynamoCore.dll")));
                if (rootFolder == null)
                {
                    throw new DirectoryNotFoundException(
                        "Dynamo Core runtime not found. Set " + RuntimeEnvironmentVariable +
                        " to a folder containing DynamoCore.dll, or place a DynamoCore folder next to DynamoInventor.dll. Probed: " +
                        string.Join("; ", candidates));
                }
                Log("Dynamo runtime: " + rootFolder);
                return rootFolder;
            }
        }

        public static void EnsureResolver()
        {
            lock (gate)
            {
                if (resolverRegistered) return;
                var probeFolders = new[] { RootFolder, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) };
                AssemblyLoadContext.Default.Resolving += (ctx, name) =>
                {
                    foreach (var folder in probeFolders)
                    {
                        var candidate = Path.Combine(folder, name.Name + ".dll");
                        if (File.Exists(candidate)) return ctx.LoadFromAssemblyPath(candidate);
                    }
                    return null;
                };
                resolverRegistered = true;
            }
        }

        /// <summary>
        /// Folder of the ASM build Inventor has already loaded into this process (Inventor 2027 loads
        /// ASM 232 from Autodesk Shared Components at startup). Dynamo's libG binds to those modules,
        /// so no second kernel is ever loaded.
        /// </summary>
        public static string FindLoadedAsmFolder()
        {
            var asm = Process.GetCurrentProcess().Modules.Cast<ProcessModule>()
                .FirstOrDefault(m => m.ModuleName.StartsWith("ASMAHL", StringComparison.OrdinalIgnoreCase));
            if (asm == null)
            {
                throw new InvalidOperationException("Inventor has not loaded an ASMAHL*.dll; cannot pair Dynamo's geometry library with the host kernel.");
            }
            return Path.GetDirectoryName(asm.FileName);
        }

        public static string LogPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DynamoInventor", "DynamoInventor.log");

        public static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine);
            }
            catch
            {
                // Logging must never take the add-in down.
            }
        }
    }

    /// <summary>
    /// Owns the single Dynamo model/view-model/window pair for this Inventor session.
    /// Every member that touches Dynamo types is NoInlining so the JIT does not resolve
    /// Dynamo assemblies before <see cref="DynamoRuntime.EnsureResolver"/> has run.
    /// </summary>
    internal static class DynamoSession
    {
        private static Dynamo.Controls.DynamoView view;
        private static Dynamo.ViewModels.DynamoViewModel viewModel;
        private static Host.InventorDynamoModel model;

        public static bool IsOpen => view != null;

        public static void Open(Inventor.Application app)
        {
            DynamoRuntime.EnsureResolver();
            OpenCore(app);
        }

        public static void Close()
        {
            if (view == null) return;
            CloseCore();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void OpenCore(Inventor.Application app)
        {
            if (view != null)
            {
                view.Activate();
                return;
            }

            try
            {
                if (!Dynamo.Wpf.Utilities.WebView2Utilities.ValidateWebView2RuntimeInstalled())
                {
                    return; // Dynamo has already told the user what to install.
                }
            }
            catch (MissingMethodException ex)
            {
                // Inventor 2027 has WebView2 SDK 1.0.1210 loaded in the default load context; Dynamo 4.2 was
                // built against 1.0.2478 and its check uses an overload the older SDK lacks. Inventor itself
                // requires the WebView2 runtime, so treat it as present and carry on (diagnostic mode).
                DynamoRuntime.Log("WebView2 SDK version conflict (continuing): " + ex.Message);
            }

            // Geometry: pair Dynamo's libG with the ASM Inventor already has in memory.
            var asmFolder = DynamoRuntime.FindLoadedAsmFolder();
            var asmVersion = DynamoShapeManager.Utilities.GetVersionFromPath(asmFolder);
            var preloaderLocation = DynamoShapeManager.Utilities.GetLibGPreloaderLocation(asmVersion, DynamoRuntime.RootFolder);
            DynamoShapeManager.Utilities.PreloadAsmFromPath(preloaderLocation, asmFolder);
            var geometryFactoryPath = Path.Combine(preloaderLocation, DynamoShapeManager.Utilities.GeometryFactoryAssembly);
            DynamoRuntime.Log("ASM " + asmVersion + " from " + asmFolder + "; libG " + preloaderLocation);

            var sw = Stopwatch.StartNew();
            var config = Host.InventorDynamoModel.CreateConfiguration(
                geometryFactoryPath, preloaderLocation, app.SoftwareVersion.DisplayVersion);
            model = Host.InventorDynamoModel.Start(config);
            DynamoRuntime.Log("DynamoModel " + Dynamo.Models.DynamoModel.Version + " started in " + sw.ElapsedMilliseconds + " ms");

            viewModel = Dynamo.ViewModels.DynamoViewModel.Start(new Dynamo.ViewModels.DynamoViewModel.StartConfiguration
            {
                DynamoModel = model,
                Watch3DViewModel = Dynamo.Wpf.ViewModels.Watch3D.HelixWatch3DViewModel.TryCreateHelixWatch3DViewModel(
                    null, new Dynamo.Wpf.ViewModels.Watch3D.Watch3DViewModelStartupParams(model), model.Logger),
                ShowLogin = false
            });

            view = new Dynamo.Controls.DynamoView(viewModel);
            new WindowInteropHelper(view).Owner = new IntPtr(app.MainFrameHWND);
            view.Closed += OnViewClosed;
            view.Show();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void CloseCore()
        {
            try { view.Close(); }
            catch (Exception ex) { DynamoRuntime.Log("Close failed: " + ex); }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void OnViewClosed(object sender, EventArgs e)
        {
            // DynamoView's own closing sequence shuts the view-model and model down; just drop our references.
            view.Closed -= OnViewClosed;
            view = null;
            viewModel = null;
            model = null;
            DynamoRuntime.Log("Dynamo window closed");
        }
    }
}
