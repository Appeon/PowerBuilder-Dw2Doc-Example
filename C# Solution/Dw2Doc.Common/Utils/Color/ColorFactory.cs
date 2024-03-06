using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Utils.Color
{
    public class ColorFactory
    {
        public static DwColorWrapper FromRgb(byte red, byte green, byte blue)
            => new()
            {
                Value = System.Drawing.Color.FromArgb(red, green, blue)
            };

        public static DwColorWrapper FromArgb(byte alpha, byte red, byte green, byte blue)
            => new()
            {
                Value = System.Drawing.Color.FromArgb(alpha, red, green, blue)
            };
    }
}
