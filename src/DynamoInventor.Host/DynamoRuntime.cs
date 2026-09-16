using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;

namespace DynamoInventor.Host
{
    /// <summary>
    /// Locates the Dynamo Core runtime and Inventor's ASM on disk and makes Dynamo's assemblies
    /// resolvable from the default AssemblyLoadContext. Nothing here may touch a Dynamo type:
    /// it runs before the resolver exists.
    /// </summary>
    public static class DynamoRuntime
    {
        public const string RuntimeEnvironmentVariable = "DYNAMO_INVENTOR_RUNTIME";
        public const string RuntimePointerFile = "dynamo-runtime.txt";

        private static readonly object gate = new object();
        private static bool resolverRegistered;
        private static string rootFolder;

        /// <summary>
        /// Folder containing DynamoCore.dll, libg_xxx_0_0\ etc. Probe order:
        /// 1. %DYNAMO_INVENTOR_RUNTIME%
        /// 2. the path written in dynamo-runtime.txt next to this assembly (dev convenience)
        /// 3. DynamoCore\ next to this assembly
        /// 4. the newest 4.x "Dynamo Core" install under %ProgramFiles%\Dynamo
        /// </summary>
        public static string RootFolder
        {
            get
            {
                if (rootFolder != null) return rootFolder;

                var candidates = new List<string>();
                var fromEnv = Environment.GetEnvironmentVariable(RuntimeEnvironmentVariable);
                if (!string.IsNullOrEmpty(fromEnv)) candidates.Add(fromEnv);

                var here = OwnFolder;
                if (here != null)
                {
                    var pointer = Path.Combine(here, RuntimePointerFile);
                    if (File.Exists(pointer))
                    {
                        var p = File.ReadAllText(pointer).Trim();
                        if (!string.IsNullOrEmpty(p)) candidates.Add(p);
                    }
                    candidates.Add(Path.Combine(here, "DynamoCore"));
                }

                var programFilesDynamo = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Dynamo", "Dynamo Core");
                if (Directory.Exists(programFilesDynamo))
                {
                    candidates.AddRange(Directory.GetDirectories(programFilesDynamo, "4.*").OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase));
                }

                rootFolder = candidates.FirstOrDefault(c => File.Exists(Path.Combine(c, "DynamoCore.dll")));
                if (rootFolder == null)
                {
                    throw new DirectoryNotFoundException(
                        "Dynamo Core runtime not found. Set " + RuntimeEnvironmentVariable + ", write its path into " +
                        RuntimePointerFile + " next to the application, or place a DynamoCore folder there. Probed: " +
                        string.Join("; ", candidates));
                }
                Log("Dynamo runtime: " + rootFolder);
                return rootFolder;
            }
        }

        public static string OwnFolder => Path.GetDirectoryName(typeof(DynamoRuntime).Assembly.Location);

        /// <summary>
        /// Registers assembly resolution for the runtime folder and this assembly's folder, and puts
        /// the runtime folder on the process PATH so native companions (WebView2Loader.dll, libG) load.
        /// Idempotent.
        /// </summary>
        public static void EnsureResolver()
        {
            lock (gate)
            {
                if (resolverRegistered) return;
                var probeFolders = new[] { RootFolder, OwnFolder }.Where(f => !string.IsNullOrEmpty(f)).ToArray();

                AssemblyLoadContext.Default.Resolving += (ctx, name) =>
                {
                    foreach (var folder in probeFolders)
                    {
                        var candidate = Path.Combine(folder, name.Name + ".dll");
                        if (File.Exists(candidate)) return ctx.LoadFromAssemblyPath(candidate);
                    }
                    return null;
                };

                var path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process) ?? string.Empty;
                if (path.IndexOf(RootFolder, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    Environment.SetEnvironmentVariable("PATH", RootFolder + ";" + path, EnvironmentVariableTarget.Process);
                }
                resolverRegistered = true;
            }
        }

        /// <summary>
        /// Folder holding the ASM build that Inventor itself uses. Inventor 2025+ installs it under
        /// Common Files\Autodesk Shared\Components\{year}\{component version}\ASMAHL*.dll rather than
        /// in its own Bin. Prefers <paramref name="preferredYear"/>, then the newest year present.
        /// Returns null when no ASM is installed that way.
        /// </summary>
        public static string FindInventorAsmFolder(int preferredYear)
        {
            var componentsRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles), "Autodesk Shared", "Components");
            if (!Directory.Exists(componentsRoot)) return null;

            var years = Directory.GetDirectories(componentsRoot)
                .Select(d => new { Path = d, Year = int.TryParse(Path.GetFileName(d), out var y) ? y : 0 })
                .Where(x => x.Year > 0)
                .OrderByDescending(x => x.Year == preferredYear)
                .ThenByDescending(x => x.Year)
                .ToList();

            foreach (var year in years)
            {
                var hit = Directory.GetDirectories(year.Path)
                    .OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault(d => Directory.GetFiles(d, "ASMAHL*.dll").Length > 0);
                if (hit != null) return hit;
            }
            return null;
        }

        public static string LogFileName { get; set; } = "DynamoInventor.log";

        public static string LogPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DynamoInventor", LogFileName);

        public static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine);
            }
            catch
            {
                // Logging must never take the host down.
            }
        }
    }
}
