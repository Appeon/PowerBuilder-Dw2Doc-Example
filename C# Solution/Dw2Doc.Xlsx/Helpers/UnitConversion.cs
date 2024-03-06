using NPOI.Util;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Helpers;

internal static class UnitConversion
{
    public static int PixelsToXlsxCellUnits(int pixels) => pixels * 256 / 6;
    //public static int PixelsToXlsxCellUnits(int pixels) => pixels * 256 / 6;
    //public static double PixelsToCharacterFraction(int pixels) => pixels / 7.001699924468994 / 255;
    public static double PixelsToColumnWidth(int pixels) => Units.PixelToEMU(pixels) * 255.0 / 66691.0;
    public static int PixelsToEMU(int pixels) => Units.PixelToEMU(pixels);
    public static double PixelsToPoints(int pixels) => Units.PixelToPoints(pixels);
}