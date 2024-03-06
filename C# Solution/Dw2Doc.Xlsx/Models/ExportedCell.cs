using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using NPOI.SS.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;

public class ExportedCell : ExportedCellBase
{
    public ICell? OutputCell { get; set; }

    public ExportedCell(VirtualCell cell,
        DwObjectAttributesBase attribute) : base(cell, attribute)
    {
    }
}
