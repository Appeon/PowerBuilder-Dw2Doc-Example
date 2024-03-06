using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.PixelFormats;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers.Xlsx
{
    [RendererFor(typeof(DwTextAttributes), typeof(XlsxTextRenderer))]
    internal class XlsxTextRenderer : AbstractXlsxRenderer
    {
        private const double TextBoxWidthAdjustment = 5;

        private static double ConvertFontSize(DwTextAttributes attributes)
        {
            return attributes.FontSize - 1;
        }

        public override ExportedCellBase? Render(
            ISheet sheet,
            VirtualCell cell,
            DwObjectAttributesBase attribute,
            ICell renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);
            XSSFCellStyle style = renderTarget.Row.Sheet.Workbook.CreateCellStyle() as XSSFCellStyle
                ?? throw new InvalidCastException("Could not get XSSFCellStyle");
            var font = (renderTarget.Row.Sheet.Workbook.CreateFont() as XSSFFont)!;
            font.FontHeightInPoints = ConvertFontSize(textAttribute);
            font.FontName = textAttribute.FontFace;
            font.SetColor(new XSSFColor(new byte[] {
                textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B}));
            font.IsBold = textAttribute.FontWeight >= 700;
            font.IsItalic = textAttribute.Italics;
            font.IsStrikeout = textAttribute.Strikethrough;
            font.Underline = textAttribute.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
            style.SetFont(font);
            style.SetFillForegroundColor(new XSSFColor(new byte[3] {
                textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B,
            }));
            style.FillPattern = textAttribute.BackgroundColor.Value.A == 0 ? FillPattern.NoFill : FillPattern.SolidForeground;
            style.Alignment = textAttribute.Alignment.ToNpoiHorizontalAlignment();
            style.VerticalAlignment = VerticalAlignment.Top;

            renderTarget.CellStyle = style;
            renderTarget.SetCellValue(textAttribute.Text);

            return new ExportedCell(cell, attribute)
            {
                OutputCell = renderTarget,
            };
        }

        public override ExportedCellBase Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);

            var anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y,
                widthAdjustment: TextBoxWidthAdjustment);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var textBox = renderTarget.draw.CreateTextbox(anchor);
            textBox.BottomInset = 0;
            textBox.LeftInset = 0;
            textBox.TopInset = 0;
            textBox.RightInset = 0;
            if (textAttribute.BackgroundColor.Value.A != 0)
            {
                textBox.SetFillColor(textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B);
            }


            var paragraph = textBox.TextParagraphs[0];
            var run = paragraph.AddNewTextRun();
            run.FontColor = new Rgb24(textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B);
            run.Text = textAttribute.Text;
            run.FontSize = ConvertFontSize(textAttribute);
            run.IsItalic = textAttribute.Italics;
            run.IsStrikethrough = textAttribute.Strikethrough;
            run.IsUnderline = textAttribute.Underline;
            run.SetFont(textAttribute.FontFace);
            run.IsBold = textAttribute.FontWeight >= 700;
            paragraph.TextAlign = textAttribute.Alignment.ToNpoiTextAlignment();
            textBox.LineStyle = LineStyle.None;
            //textBox.SetText();

            return new ExportedFloatingCell(cell, attribute)
            {
                OutputShape = textBox,
            };
        }
    }
}
