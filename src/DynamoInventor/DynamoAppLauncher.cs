using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DynamoInventor
{
    /// <summary>
    /// Starts DynamoInventor.App.exe (the out-of-process Dynamo host) or brings the running one forward.
    /// The add-in itself has no Dynamo dependency: everything Dynamo lives in that process.
    /// </summary>
    internal static class DynamoAppLauncher
    {
        public const string AppExeName = "DynamoInventor.App.exe";
        public const string AppPathEnvironmentVariable = "DYNAMO_INVENTOR_APP";

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public static void OpenOrFocus(Inventor.Application inventor)
        {
            // Any live host process counts, even one still starting up (its window appears a few seconds
            // in); the host's own single-instance mutex is the backstop if two launches race.
            var running = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(AppExeName)).FirstOrDefault();
            if (running != null)
            {
                if (running.MainWindowHandle != IntPtr.Zero)
                {
                    ShowWindow(running.MainWindowHandle, SW_RESTORE);
                    SetForegroundWindow(running.MainWindowHandle);
                }
                AddinLog.Log("Dynamo already running (pid " + running.Id + "); focused");
                return;
            }

            var exe = FindApp();
            var psi = new ProcessStartInfo(exe)
            {
                UseShellExecute = false,
                WorkingDirectory = Path.GetDirectoryName(exe),
                Arguments = "--inventor-version \"" + inventor.SoftwareVersion.DisplayVersion + "\""
            };
            var process = Process.Start(psi);
            AddinLog.Log("Started " + exe + " (pid " + (process?.Id.ToString() ?? "?") + ")");
        }

        /// <summary>%DYNAMO_INVENTOR_APP% (full exe path) or DynamoInventor.App.exe beside this DLL.</summary>
        private static string FindApp()
        {
            var fromEnv = Environment.GetEnvironmentVariable(AppPathEnvironmentVariable);
            if (!string.IsNullOrEmpty(fromEnv) && File.Exists(fromEnv)) return fromEnv;

            var beside = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty, AppExeName);
            if (File.Exists(beside)) return beside;

            throw new FileNotFoundException(
                AppExeName + " was not found next to the add-in. Set " + AppPathEnvironmentVariable + " to its full path.", beside);
        }
    }

    internal static class AddinLog
    {
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
}
