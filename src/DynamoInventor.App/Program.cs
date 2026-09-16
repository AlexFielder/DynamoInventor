using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using DynamoInventor.Host;

namespace DynamoInventor.App
{
    /// <summary>
    /// Entry point. Mirrors DynamoSandbox's DynamoCoreSetup: preload ASM/libG, start the model,
    /// start the view-model, show DynamoView, run the WPF message loop.
    ///
    /// Arguments:
    ///   --inventor-version 2027.1   host version for analytics and ASM year selection (optional)
    ///   --no-network                start Dynamo in no-network mode (optional)
    /// </summary>
    internal static class Program
    {
        private const string SingleInstanceMutexName = @"Local\DynamoInventor.App";
        private const string HostName = "Dynamo Inventor";

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        [STAThread]
        private static int Main(string[] args)
        {
            using var mutex = new Mutex(true, SingleInstanceMutexName, out var firstInstance);
            if (!firstInstance)
            {
                FocusRunningInstance();
                return 0;
            }

            DynamoRuntime.LogFileName = "DynamoInventor.App.log";
            DynamoRuntime.Log("=== DynamoInventor.App start; args: " + string.Join(" ", args));

            try
            {
                DynamoRuntime.EnsureResolver();
            }
            catch (Exception ex)
            {
                DynamoRuntime.Log("Runtime not found: " + ex);
                System.Windows.MessageBox.Show(ex.Message, "Dynamo for Inventor",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return 1;
            }

            return Run(args);
        }

        // Separate, non-inlined method: nothing in Main may JIT-resolve a Dynamo type before EnsureResolver ran.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int Run(string[] args)
        {
            var app = new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose };
            Dynamo.ViewModels.DynamoViewModel viewModel = null;

            try
            {
                if (!Dynamo.Wpf.Utilities.WebView2Utilities.ValidateWebView2RuntimeInstalled())
                {
                    return 2; // Dynamo has already told the user what to install.
                }

                var hostVersion = ArgValue(args, "--inventor-version") ?? QueryRunningInventorVersion() ?? string.Empty;
                var noNetwork = args.Any(a => string.Equals(a, "--no-network", StringComparison.OrdinalIgnoreCase));
                var year = ParseYear(hostVersion, fallback: 2027);

                // Geometry: use the ASM build Inventor itself ships (Autodesk Shared Components), so Dynamo
                // geometry round-trips with Inventor's kernel version-for-version.
                var asmFolder = DynamoRuntime.FindInventorAsmFolder(year)
                    ?? throw new DirectoryNotFoundException("No Inventor ASM found under Common Files\\Autodesk Shared\\Components.");
                var asmVersion = DynamoShapeManager.Utilities.GetVersionFromPath(asmFolder);
                var preloaderLocation = DynamoShapeManager.Utilities.GetLibGPreloaderLocation(asmVersion, DynamoRuntime.RootFolder);
                DynamoShapeManager.Utilities.PreloadAsmFromPath(preloaderLocation, asmFolder);
                var geometryFactoryPath = Path.Combine(preloaderLocation, DynamoShapeManager.Utilities.GeometryFactoryAssembly);
                DynamoRuntime.Log("ASM " + asmVersion + " from " + asmFolder + "; libG " + preloaderLocation);

                var userData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Dynamo", HostName);
                var commonData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Dynamo", HostName);

                var sw = Stopwatch.StartNew();
                var config = InventorDynamoModel.CreateConfiguration(geometryFactoryPath, preloaderLocation, hostVersion, userData, commonData);
                config.NoNetworkMode = noNetwork;
                var model = InventorDynamoModel.Start(config);
                DynamoRuntime.Log("DynamoModel " + Dynamo.Models.DynamoModel.Version + " started in " + sw.ElapsedMilliseconds + " ms (host '" + hostVersion + "')");
                LogLibraryDiagnostics(model);

                viewModel = Dynamo.ViewModels.DynamoViewModel.Start(new Dynamo.ViewModels.DynamoViewModel.StartConfiguration
                {
                    DynamoModel = model,
                    Watch3DViewModel = Dynamo.Wpf.ViewModels.Watch3D.HelixWatch3DViewModel.TryCreateHelixWatch3DViewModel(
                        null, new Dynamo.Wpf.ViewModels.Watch3D.Watch3DViewModelStartupParams(model), model.Logger),
                    ShowLogin = false
                });

                var view = new Dynamo.Controls.DynamoView(viewModel);
                DynamoRuntime.Log("Showing DynamoView");

                // Nodes must be added after the view-model exists (NodeViewModel creation needs it) and,
                // like a user placing a node, on the UI thread once the window is up.
                var smokeTestCode = ArgValue(args, "--smoke-test");
                if (smokeTestCode != null)
                {
                    view.Loaded += (s, e) => StartSmokeTest(model, smokeTestCode);
                }
                app.Run(view);
                DynamoRuntime.Log("=== DynamoInventor.App exit");
                return 0;
            }
            catch (Exception ex)
            {
                DynamoRuntime.Log("FATAL: " + ex);
                try { viewModel?.Exit(allowCancel: false); } catch { }
                System.Windows.MessageBox.Show(
                    ex.Message + Environment.NewLine + Environment.NewLine + "Details: " + DynamoRuntime.LogPath,
                    "Dynamo for Inventor", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return 1;
            }
        }

        /// <summary>Logs the imported libraries and how many Inventor nodes made it into the search index.</summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void LogLibraryDiagnostics(Dynamo.Models.DynamoModel model)
        {
            try
            {
                var libraries = model.LibraryServices.ImportedLibraries.ToList();
                DynamoRuntime.Log("Imported libraries (" + libraries.Count + "): " + string.Join(", ", libraries.Select(Path.GetFileName)));

                var entries = model.SearchModel.Entries.ToList();
                var inventorEntries = entries.Where(e => e.FullName.StartsWith("InventorLibrary", StringComparison.OrdinalIgnoreCase)).ToList();
                DynamoRuntime.Log("Search entries: " + entries.Count + " total, " + inventorEntries.Count + " from InventorLibrary" +
                    (inventorEntries.Count > 0 ? ", e.g. " + string.Join(", ", inventorEntries.Take(5).Select(e => e.FullName)) : ""));
            }
            catch (Exception ex)
            {
                DynamoRuntime.Log("Library diagnostics failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Drops a code block containing <paramref name="code"/> into the home workspace and logs the
        /// result of the automatic run. Used to exercise the Inventor nodes end to end over COM, e.g.
        ///   --smoke-test "InventorWorkPoint.ByPoint(Point.ByCoordinates(1,2,3));"
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void StartSmokeTest(Dynamo.Models.DynamoModel model, string code)
        {
            DynamoRuntime.Log("SMOKE TEST: " + code);
            var node = new Dynamo.Graph.Nodes.CodeBlockNodeModel(code, 100, 100, model.LibraryServices, model.CurrentWorkspace.ElementResolver);

            model.EvaluationCompleted += (s, e) =>
            {
                try
                {
                    DynamoRuntime.Log("SMOKE TEST evaluation completed: succeeded=" + e.EvaluationSucceeded +
                        (e.EvaluationSucceeded ? "" : "; error=" + e.Error?.Message));
                    foreach (var n in model.CurrentWorkspace.Nodes)
                    {
                        var infos = string.Join(" | ", n.NodeInfos.Select(i => i.State + ": " + i.Message.Replace(Environment.NewLine, " ")));
                        DynamoRuntime.Log("  node " + n.Name + ": state=" + n.State +
                            (infos.Length == 0 ? "" : "; " + infos) +
                            "; value=" + DescribeValue(n));
                    }
                }
                catch (Exception ex)
                {
                    DynamoRuntime.Log("SMOKE TEST reporting failed: " + ex.Message);
                }
            };

            model.ExecuteCommand(new Dynamo.Models.DynamoModel.CreateNodeCommand(node, 100, 100, false, false));
        }

        private static string DescribeValue(Dynamo.Graph.Nodes.NodeModel n)
        {
            try
            {
                var v = n.CachedValue;
                if (v == null) return "<null>";
                return v.IsCollection ? "collection[" + v.GetElements().Count() + "]" : (v.Data?.ToString() ?? "<null data>");
            }
            catch (Exception ex)
            {
                return "<" + ex.GetType().Name + ">";
            }
        }

        private static void FocusRunningInstance()
        {
            var me = Process.GetCurrentProcess();
            var other = Process.GetProcessesByName(me.ProcessName).FirstOrDefault(p => p.Id != me.Id && p.MainWindowHandle != IntPtr.Zero);
            if (other == null) return;
            ShowWindow(other.MainWindowHandle, SW_RESTORE);
            SetForegroundWindow(other.MainWindowHandle);
        }

        /// <summary>Asks a running Inventor for its version over COM; null when none is running.</summary>
        private static string QueryRunningInventorVersion()
        {
            try
            {
                var inv = (Inventor.Application)InventorServices.Utilities.ComInterop.GetActiveObject("Inventor.Application");
                return inv.SoftwareVersion.DisplayVersion;
            }
            catch
            {
                return null;
            }
        }

        private static string ArgValue(string[] args, string name)
        {
            for (var i = 0; i < args.Length - 1; i++)
            {
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
            }
            return null;
        }

        private static int ParseYear(string displayVersion, int fallback)
        {
            var head = (displayVersion ?? string.Empty).Split('.').FirstOrDefault();
            return int.TryParse(head, out var y) && y > 2000 ? y : fallback;
        }
    }
}
