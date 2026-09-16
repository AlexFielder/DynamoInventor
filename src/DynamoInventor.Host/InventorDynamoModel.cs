using System;
using System.IO;
using Dynamo.Configuration;
using Dynamo.Core;
using Dynamo.Models;
using Dynamo.Scheduler;

namespace DynamoInventor.Host
{
    /// <summary>
    /// Inventor-hosted DynamoModel. Configuration is built on Dynamo's own
    /// <see cref="DynamoModel.DefaultStartConfiguration"/>, the same struct DynamoSandbox and the
    /// CLI use, so it tracks whatever members Dynamo adds or removes between releases.
    /// </summary>
    public class InventorDynamoModel : DynamoModel
    {
        /// <summary>
        /// Builds the start configuration for hosting inside Inventor.
        /// </summary>
        /// <param name="geometryFactoryPath">Full path to LibG.ProtoInterface.dll for the ASM build Inventor has loaded.</param>
        /// <param name="preloaderLocation">The libg_xxx folder that geometryFactoryPath lives in.</param>
        /// <param name="hostVersion">Inventor's display version, e.g. "2027.1", for analytics and the title bar.</param>
        /// <param name="userDataFolder">Optional override for Dynamo's per-user data folder; null keeps Dynamo's default.</param>
        public static DefaultStartConfiguration CreateConfiguration(
            string geometryFactoryPath, string preloaderLocation, string hostVersion, string userDataFolder = null)
        {
            if (string.IsNullOrEmpty(geometryFactoryPath) || !File.Exists(geometryFactoryPath))
            {
                throw new FileNotFoundException("Dynamo geometry factory not found.", geometryFactoryPath);
            }

            return new DefaultStartConfiguration
            {
                GeometryFactoryPath = geometryFactoryPath,
                ProcessMode = TaskProcessMode.Asynchronous,
                StartInTestMode = false,
                // Preferences left null: DynamoModel loads or creates the per-host preferences file itself.
                PathResolver = new InventorPathResolver(preloaderLocation, userDataFolder),
                HostAnalyticsInfo = new HostAnalyticsInfo
                {
                    HostName = "Dynamo Inventor",
                    HostProductName = "Inventor",
                    HostProductVersion = ParseVersion(hostVersion),
                    HostVersion = ParseVersion(Version)
                },
                // AuthProvider deliberately not set (Greg.IAuthProvider): sign-in is a package-manager concern
                // and wiring Dynamo.Core.IDSDKManager needs the Greg assembly, which the NuGet packages omit.
                NoNetworkMode = false
            };
        }

        public static InventorDynamoModel Start(DefaultStartConfiguration configuration)
        {
            return new InventorDynamoModel(configuration);
        }

        private InventorDynamoModel(IStartConfiguration configuration)
            : base(configuration)
        {
        }

        private static System.Version ParseVersion(string displayVersion)
        {
            if (System.Version.TryParse(displayVersion, out var v))
            {
                return v.Build < 0 ? new System.Version(v.Major, v.Minor, 0) : v;
            }
            return new System.Version(0, 0, 0);
        }
    }
}
