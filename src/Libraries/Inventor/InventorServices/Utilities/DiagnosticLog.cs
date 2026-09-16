using System;
using System.IO;

namespace InventorServices.Utilities
{
    /// <summary>
    /// Append-only diagnostic log for the node library. Dynamo only surfaces an exception's message
    /// on the node ("X operation failed. ..."); this keeps the full exception and stack.
    /// File: %LOCALAPPDATA%\DynamoInventor\InventorLibrary.log
    /// </summary>
    public static class DiagnosticLog
    {
        public static string LogPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DynamoInventor", "InventorLibrary.log");

        public static void Write(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message + Environment.NewLine);
            }
            catch
            {
                // Never let logging break a node.
            }
        }

        public static void Write(string context, Exception ex)
        {
            Write(context + ": " + ex);
        }
    }
}
