using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers
{
    [RendererFor(typeof(DwLineAttributes), typeof(XlsxLineRenderer))]
    internal class XlsxLineRenderer : AbstractXlsxRenderer
    {
        public override ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget)
        {
            throw new NotImplementedException();
        }

        public override ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var lineAttribute = CheckAttributeType<DwLineAttributes>(attribute);

            XSSFClientAnchor anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y);

            anchor.AnchorType = AnchorType.MoveDontResize;

            XSSFShape shapeBase;
            if (cell.Object.Height == 0 || cell.Object.Width == 0)
            { // Line is straight 
                var shape = renderTarget.draw.CreateConnector(anchor);
                shapeBase = shape;
                shape.ShapeType = NPOI.OpenXmlFormats.Dml.ST_ShapeType.line;
            }
            else
            {
                var shape = renderTarget.draw.CreateSimpleShape(anchor);
                shapeBase = shape;
                shape.ShapeType = (int)ShapeTypes.Line;
            }

            shapeBase.LineStyle = lineAttribute.LineStyle.DwLineStyleToNpoiLineStyle();
            shapeBase.LineStyleColor = lineAttribute.LineColor.Value.ToRgb();
            shapeBase.LineWidth = lineAttribute.LineWidth;

            return new ExportedFloatingCell(cell, attribute)
            {
                OutputShape = shapeBase,
            };
        }
    }
}
