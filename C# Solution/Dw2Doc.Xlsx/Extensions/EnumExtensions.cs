using Appeon.DotnetDemo.Dw2Doc.Common.Enums;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions
{
    public static class EnumExtensions
    {
        public static NPOI.SS.UserModel.LineStyle DwLineStyleToNpoiLineStyle(this LineStyle style) => style switch
        {
            LineStyle.Dotted => NPOI.SS.UserModel.LineStyle.DotSys,
            LineStyle.Solid => NPOI.SS.UserModel.LineStyle.Solid,
            LineStyle.Dash => NPOI.SS.UserModel.LineStyle.DashSys,
            LineStyle.DashDot => NPOI.SS.UserModel.LineStyle.DashDotSys,
            LineStyle.DashDotDot => NPOI.SS.UserModel.LineStyle.DashDotDotSys,
            LineStyle.NoLine => NPOI.SS.UserModel.LineStyle.None,
            _ => throw new NotImplementedException("Unsupported LineStyle value"),
        };
    }
}
