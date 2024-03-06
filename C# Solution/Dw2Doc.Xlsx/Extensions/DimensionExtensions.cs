using Appeon.DotnetDemo.Dw2Doc.Xlsx.Windows;
using static Appeon.DotnetDemo.Dw2Doc.Xlsx.Windows.ScreenTools;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions
{
    internal static class DimensionExtensions
    {
        public static int TwipsToPixels(this int twips, MeasureDirection direction)
        {
            return (int)(twips * ((double)GetPPI(direction)) / 1440.0);
        }

        public static int PixelsToTwips(this int pixels, MeasureDirection direction)
        {
            return (int)(pixels * 1440.0 / GetPPI(direction));
        }
    }
}
