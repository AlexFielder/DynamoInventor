using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Dynamo.Interfaces;

namespace DynamoInventor.Models
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
