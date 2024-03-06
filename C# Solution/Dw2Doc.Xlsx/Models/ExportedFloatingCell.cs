using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;

public class ExportedFloatingCell : ExportedCellBase
{
    public XSSFShape? OutputShape { get; set; }

    public ExportedFloatingCell(
        FloatingVirtualCell cell,
        DwObjectAttributesBase attribute) : base(cell, attribute)
    {
    }
}
