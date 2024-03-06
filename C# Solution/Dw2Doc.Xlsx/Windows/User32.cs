using System.Runtime.InteropServices;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Windows
{
    internal static class User32
    {
        [DllImport("user32.dll")]
        internal static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);
    }
}
