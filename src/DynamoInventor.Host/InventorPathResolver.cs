using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Dynamo.Interfaces;

namespace DynamoInventor.Host
{
    /// <summary>
    /// Tells Dynamo where the Inventor-specific bits live: the libG preloader folder for assembly
    /// resolution and InventorLibrary.dll as a preloaded node library. Mirrors what DynamoRevit's
    /// path resolver does for Revit.
    /// </summary>
    public sealed class InventorPathResolver : IPathResolver
    {
        private readonly List<string> additionalResolutionPaths = new List<string>();
        private readonly List<string> additionalNodeDirectories = new List<string>();
        private readonly List<string> preloadedLibraryPaths = new List<string>();

        public InventorPathResolver(string preloaderLocation, string userDataFolder = null, string commonDataFolder = null)
        {
            var addinFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (!string.IsNullOrEmpty(preloaderLocation) && Directory.Exists(preloaderLocation))
            {
                additionalResolutionPaths.Add(preloaderLocation);
            }
            // Dynamo does NOT merge this list with a built-in one: whatever the resolver returns IS the set
            // of libraries the DesignScript VM starts with. Omitting the core set leaves the VM with no
            // executable ("Assertion failed: 'exe != null'" in LiveRunner). Same list as Dynamo's own
            // SandboxPathResolver; bare names resolve against the Dynamo runtime folder.
            preloadedLibraryPaths.AddRange(new[]
            {
                "VMDataBridge.dll",
                "ProtoGeometry.dll",
                "DesignScriptBuiltin.dll",
                "DSCoreNodes.dll",
                "DSOffice.dll",
                "DSCPython.dll",
                "FunctionObject.ds",
                "BuiltIn.ds",
                "DynamoConversions.dll",
                "DynamoUnits.dll",
                "Tessellation.dll",
                "Analysis.dll",
                "GeometryColor.dll"
            });

            if (!string.IsNullOrEmpty(addinFolder))
            {
                additionalResolutionPaths.Add(addinFolder);

                var inventorLibrary = Path.Combine(addinFolder, "InventorLibrary.dll");
                if (File.Exists(inventorLibrary))
                {
                    preloadedLibraryPaths.Add(inventorLibrary);
                }
            }

            UserDataRootFolder = userDataFolder ?? string.Empty;
            CommonDataRootFolder = commonDataFolder ?? string.Empty;
        }

        public IEnumerable<string> AdditionalResolutionPaths => additionalResolutionPaths;
        public IEnumerable<string> AdditionalNodeDirectories => additionalNodeDirectories;
        public IEnumerable<string> PreloadedLibraryPaths => preloadedLibraryPaths;

        /// <summary>Empty means "use Dynamo's default per-user folder".</summary>
        public string UserDataRootFolder { get; }

        /// <summary>Empty means "use Dynamo's default common folder".</summary>
        public string CommonDataRootFolder { get; }

        /// <summary>
        /// User-data folders of every Dynamo product on this machine, used to migrate preferences
        /// and packages. Here that is every "Dynamo Inventor" folder under %AppData%\Dynamo.
        /// </summary>
        public IEnumerable<string> GetDynamoUserDataLocations()
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Dynamo", "Dynamo Inventor");
            if (!Directory.Exists(root))
            {
                return Array.Empty<string>();
            }
            return Directory.GetDirectories(root);
        }
    }
}
