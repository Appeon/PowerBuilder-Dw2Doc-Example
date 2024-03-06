using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers
{
    [RendererFor(typeof(DwButtonAttributes), typeof(XlsxButtonRenderer))]
    internal class XlsxButtonRenderer : AbstractXlsxRenderer
    {
        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            throw new NotImplementedException();
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var buttonAttributes = CheckAttributeType<DwButtonAttributes>(attribute);

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var shape = renderTarget.draw.CreateSimpleShape(anchor);
            shape.FillColor = Color.White.ToRgb();
            shape.LineStyle = LineStyle.Solid;
            shape.LineStyleColor = Color.Gray.ToRgb();
            shape.LineWidth = 1;
            shape.ShapeType = (int)ShapeTypes.RoundedRectangle;
            var paragraph = shape.TextParagraphs[0];
            paragraph.TextAlign = TextAlign.CENTER;
            var run = paragraph.AddNewTextRun();
            run.FontSize = buttonAttributes.FontSize;
            run.Text = buttonAttributes.Text;

            return new ExportedFloatingCell(cell, attribute)
            {
                OutputShape = shape,
            };
        }
    }
}
