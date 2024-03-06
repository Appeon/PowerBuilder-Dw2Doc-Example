using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions
{
    public static class ColorExtensions
    {
        public static int ToRgb(this Color color)
        {
            return color.ToArgb() & 16777215; /// this mask removes the highest 8 bits
        }
    }
}
