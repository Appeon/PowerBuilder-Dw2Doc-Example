namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Windows

{
    internal static class ScreenTools
    {
        public enum MeasureDirection
        {
            Horizontal,
            Vertical
        }

        public static int GetPPI(MeasureDirection direction)
        {
            int ppi;
            IntPtr dc = User32.GetDC(IntPtr.Zero);

            if (direction == MeasureDirection.Horizontal)
                ppi = Gdi32.GetDeviceCaps(dc, 88); //DEVICECAP LOGPIXELSX
            else
                ppi = Gdi32.GetDeviceCaps(dc, 90); //DEVICECAP LOGPIXELSY

            User32.ReleaseDC(IntPtr.Zero, dc);
            return ppi;
        }
    }
}
