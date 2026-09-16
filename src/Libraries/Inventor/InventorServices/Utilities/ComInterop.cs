using System;
using System.Runtime.InteropServices;

namespace InventorServices.Utilities
{
    /// <summary>
    /// Replacement for <c>Marshal.GetActiveObject</c>, which does not exist on .NET Core / .NET 5+.
    /// Looks up a running COM server in the Running Object Table by ProgID.
    /// </summary>
    public static class ComInterop
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void CLSIDFromProgID(string progId, out Guid clsid);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);

        /// <summary>
        /// Returns the running instance registered for <paramref name="progId"/>, or throws
        /// <see cref="COMException"/> (MK_E_UNAVAILABLE) when none is running.
        /// </summary>
        public static object GetActiveObject(string progId)
        {
            CLSIDFromProgID(progId, out var clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
            return obj;
        }
    }
}
