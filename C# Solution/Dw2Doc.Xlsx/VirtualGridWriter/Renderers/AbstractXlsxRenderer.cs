using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Common;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Helpers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers
{
    internal abstract class AbstractXlsxRenderer
        : IRenderer<ISheet, VirtualCell, DwObjectAttributesBase, ICell>,
            IRenderer<ISheet, FloatingVirtualCell, DwObjectAttributesBase, (int x, int y, XSSFDrawing draw)>
    {
        public abstract ExportedCellBase? Render(ISheet sheet, VirtualCell cell, DwObjectAttributesBase attribute, ICell renderTarget);
        public abstract ExportedCellBase? Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget);

        protected static T CheckAttributeType<T>(DwObjectAttributesBase attribute)
            where T : DwObjectAttributesBase
        {
            var instanceType = typeof(T);
            if (attribute.GetType() != instanceType)
                throw new ArgumentException($"Attribute {nameof(attribute)} is not of type {instanceType.Name}");

            return (T)attribute;
        }

        protected static XSSFClientAnchor GetAnchor(XSSFDrawing drawing,
                                                    FloatingVirtualCell floatingCell,
                                                    int startColumn,
                                                    int startRow,
                                                    double widthAdjustment = 0.0,
                                                    double heightAdjustment = 0.0) =>
            (drawing.CreateAnchor(
                UnitConversion.PixelsToEMU((floatingCell.Offset.StartOffset.X)),
                UnitConversion.PixelsToEMU(floatingCell.Offset.StartOffset.Y),
                UnitConversion.PixelsToEMU((int)(-floatingCell.Offset.EndOffset.X + widthAdjustment)),
                UnitConversion.PixelsToEMU((int)(-floatingCell.Offset.EndOffset.Y + heightAdjustment)),
                startColumn,
                startRow,
                startColumn + floatingCell.Offset.ColSpan,
                startRow + floatingCell.Offset.RowSpan) as XSSFClientAnchor)!;

    }
}
